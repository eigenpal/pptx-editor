// Diagram editing: typing in a diagram shape writes the drawing shape and its data point, and
// adding and deleting shapes changes the data model and lays the saved drawing out again. Undo,
// refusals, and save and reopen.

import { describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { readPptx, serializeXmlPart } from '../../store/index.ts';
import { createPptxEditor, type PptxEditorInstance } from '../index.ts';
import { editorCaretOf, toolbarContextOf } from '../adapter-support.ts';
import { CANVAS_CONTEXT_MENUS, canvasMenuKind } from '../chrome-menus.ts';
import { runToolbarCommand, toolbarCommandState } from '../chrome-controls.ts';
import { listenPaneRequests, visibleSidePanelTabs } from '../pane-requests.ts';
import { runChromeRow } from '../chrome-actions.ts';
import { ribbonContextActive } from '../ribbon.ts';
import { firstCellCaret } from '../table-keys.ts';
import { diagramPress, diagramTab, diagramTextItems } from '../surface-diagram.ts';

const FIXTURES = join(import.meta.dir, '../../../../../fixtures');
const fixture = (name: string): Uint8Array =>
  new Uint8Array(readFileSync(join(FIXTURES, `${name}.pptx`)));

async function open(
  name = 'diagrams',
  slide = 0,
  bytes = fixture(name)
): Promise<PptxEditorInstance> {
  const e = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await e.load(bytes)).ok).toBe(true);
  e.goToSlide(slide);
  return e;
}

const frameOf = (e: PptxEditorInstance): string =>
  e.snapshot().shapes!.find((s) => s.name.startsWith('Diagram'))!.nodeId;

function partText(e: PptxEditorInstance, name: string): string {
  const part = e.session!.store.package.parts.get(name);
  if (!part) throw new Error(`no part ${name}`);
  return part.kind === 'xml' ? serializeXmlPart(part.xml) : new TextDecoder().decode(part.bytes);
}

/** The data point id whose drawing shape shows `text`. */
function nodeWithText(e: PptxEditorInstance, text: string): string {
  const layout = e.session!.layout(e.snapshot().currentSlide)!;
  const frame = layout.items.find((i) => i.element.nodeId === frameOf(e));
  if (frame?.kind !== 'frame') throw new Error('no frame');
  const walk = (items: readonly any[]): string | undefined => {
    for (const it of items) {
      const t = it.element.text?.paragraphs
        ?.map((p: any) => p.runs?.map((r: any) => r.text ?? '').join(''))
        .join('\n');
      if (it.element.diagramNode && t === text) return it.element.diagramNode;
      if (it.children) {
        const inner = walk(it.children);
        if (inner) return inner;
      }
    }
    return undefined;
  };
  const node = walk(frame.drawing ?? []);
  if (!node) throw new Error(`no node shows ${text}`);
  return node;
}

function caretIn(e: PptxEditorInstance, node: string, offset = 0): void {
  const at = { paragraph: 0, offset };
  e.setSelection({ kind: 'text', nodeId: frameOf(e), diagramNode: node, anchor: at, focus: at });
}

describe('diagram text', () => {
  test('typing writes the drawing shape and the data point, keeping run formatting', async () => {
    const e = await open();
    const node = nodeWithText(e, 'Plan');
    caretIn(e, node, 4);
    expect(e.snapshot().diagram?.node).toBe(node);
    for (const ch of ' & ship') expect(e.input('text', ch).ok).toBe(true);
    const data = partText(e, '/ppt/diagrams/data1.xml');
    const drawing = partText(e, '/ppt/diagrams/drawing1.xml');
    expect(data).toContain('<a:t>Plan &amp; ship</a:t>');
    expect(drawing).toContain(
      '<a:rPr lang="en-US" sz="2000" kern="1200"/><a:t>Plan &amp; ship</a:t>'
    );
    expect(nodeWithText(e, 'Plan & ship')).toBe(node);
    // One undo removes the typed burst.
    expect(e.exec({ type: 'undo' }).ok).toBe(true);
    expect(partText(e, '/ppt/diagrams/data1.xml')).toContain('<a:t>Plan</a:t>');
    e.destroy();
  });
});

describe('diagram structure', () => {
  test('add after, before, and below; delete', async () => {
    const e = await open();
    const plan = nodeWithText(e, 'Plan');
    caretIn(e, plan);
    expect(e.snapshot().diagram?.structure).toBe(true);
    expect(e.exec({ type: 'addDiagramShape', position: 'after' }).ok).toBe(true);
    const sel = e.snapshot().selection;
    expect(sel.kind === 'text' && sel.diagramNode).toBeTruthy();
    for (const ch of 'Test') e.input('text', ch);
    expect(nodeWithText(e, 'Test')).toBeTruthy();
    const data = partText(e, '/ppt/diagrams/data1.xml');
    expect(data).toContain('<a:t>Test</a:t>');
    // Four top-level columns: the new one second.
    const drawing = partText(e, '/ppt/diagrams/drawing1.xml');
    expect((drawing.match(/<dsp:sp /g) ?? []).length).toBe(10);
    expect(e.exec({ type: 'addDiagramShape', position: 'below' }).ok).toBe(true);
    for (const ch of 'Kid') e.input('text', ch);
    expect(nodeWithText(e, 'Kid')).toBeTruthy();
    expect(e.exec({ type: 'deleteDiagramShape', node: nodeWithText(e, 'Test') }).ok).toBe(true);
    const after = partText(e, '/ppt/diagrams/data1.xml');
    expect(after).not.toContain('Test');
    expect(after).not.toContain('Kid');
    // Save and reopen: both parts read back.
    const saved = e.save()!;
    const pkg = readPptx(saved);
    expect(pkg.ok).toBe(true);
    e.destroy();
  });

  test('the last top-level shape stays', async () => {
    const e = await open();
    caretIn(e, nodeWithText(e, 'Plan'));
    e.exec({ type: 'deleteDiagramShape' });
    caretIn(e, nodeWithText(e, 'Write'));
    e.exec({ type: 'deleteDiagramShape' });
    caretIn(e, nodeWithText(e, 'Ship'));
    const r = e.exec({ type: 'deleteDiagramShape' });
    expect(r.ok).toBe(false);
    if (r.ok) return;
    expect(r.reason).toBe('errors.diagram.lastShape');
    e.destroy();
  });

  test('a delete removes the presentation points and connections of the subtree', async () => {
    const e = await open();
    caretIn(e, nodeWithText(e, 'Write'));
    expect(e.exec({ type: 'deleteDiagramShape' }).ok).toBe(true);
    const data = partText(e, '/ppt/diagrams/data1.xml');
    const points = new Set([...data.matchAll(/<dgm:pt modelId="([^"]+)"/g)].map((m) => m[1]));
    const named = [...data.matchAll(/(?:srcId|destId|presAssocID)="([^"]+)"/g)].map((m) => m[1]);
    expect(named.filter((id) => !points.has(id))).toEqual([]);
    for (const text of ['Write', 'Draft', 'Art']) expect(data).not.toContain(`<a:t>${text}<`);
    // Undo brings both parts back.
    expect(e.exec({ type: 'undo' }).ok).toBe(true);
    expect(partText(e, '/ppt/diagrams/data1.xml')).toContain('<a:t>Draft</a:t>');
    e.destroy();
  });
});

describe('diagram chrome and keys', () => {
  test('the diagram group, menu, and tab follow the selection', async () => {
    const e = await open('diagram-editing');
    e.setSelection({ kind: 'shapes', nodeIds: [frameOf(e)] });
    const snap = e.snapshot();
    expect(snap.diagram?.structure).toBe(true);
    expect(toolbarContextOf(snap)).toBe('diagram');
    expect(canvasMenuKind(snap)).toBe('diagram');
    expect(ribbonContextActive(snap, 'diagram')).toBe(true);
    expect(ribbonContextActive(snap, 'shape')).toBe(false);
    // With only the frame selected, Add after and before run; Below and Delete need a node.
    expect(toolbarCommandState(e, 'diagram.addAfter').enabled).toBe(true);
    expect(toolbarCommandState(e, 'diagram.addBefore').enabled).toBe(true);
    expect(toolbarCommandState(e, 'diagram.delete').disabledReason).toBe('errors.diagram.noNode');
    caretIn(e, nodeWithText(e, 'Build'));
    expect(toolbarCommandState(e, 'diagram.delete').enabled).toBe(true);
    // The process shows no children: Add shape below is refused before it writes.
    expect(toolbarCommandState(e, 'diagram.addBelow').disabledReason).toBe(
      'errors.diagram.noChildren'
    );
    e.destroy();
  });

  test('Enter enters the first node; Tab moves between nodes', async () => {
    const e = await open('diagram-editing');
    const layout = () => e.session!.layout(0);
    const first = firstCellCaret(layout(), frameOf(e));
    expect(first?.kind === 'text' && first.diagramNode).toBe(nodeWithText(e, 'Plan'));
    const next = diagramTab(layout(), first!, false);
    expect(next?.kind === 'text' && next.diagramNode).toBe(nodeWithText(e, 'Build'));
    const back = diagramTab(layout(), first!, true);
    expect(back?.kind === 'text' && back.diagramNode).toBe(nodeWithText(e, 'Ship'));
    e.destroy();
  });

  test('a press on a node of the selected diagram places the caret', async () => {
    const e = await open('diagram-editing');
    const frame = frameOf(e);
    const lay = e.session!.layout(0)!;
    const item = diagramTextItems(lay, frame)[1]!;
    const { cx, cy } = item.element.xfrm;
    const m = item.matrix;
    const inside = {
      x: m[0] * (cx / 2) + m[2] * (cy / 2) + m[4],
      y: m[1] * (cx / 2) + m[3] * (cy / 2) + m[5],
    };
    // Not selected yet: the press selects the frame first.
    expect(diagramPress(lay, frame, { kind: 'none' }, inside, false)).toBeNull();
    const sel = diagramPress(lay, frame, { kind: 'shapes', nodeIds: [frame] }, inside, false);
    expect(sel?.kind === 'text' && sel.diagramNode).toBe(nodeWithText(e, 'Build'));
    e.destroy();
  });

  test('a shape that shows the text of two nodes is read-only', async () => {
    const files = unzipSync(fixture('diagram-editing'));
    const data = strFromU8(files['ppt/diagrams/data1.xml']!);
    // A second presOf connection to the first box: the box shows the text of two nodes.
    const box = /<dsp:sp modelId="([^"]+)"/.exec(
      strFromU8(files['ppt/diagrams/drawing1.xml']!)
    )![1];
    const other = '{00000003-0000-4000-8000-000000000006}';
    const extra = `<dgm:cxn modelId="{AAAAAAAA-0000-4000-8000-000000000001}" type="presOf" srcId="${other}" destId="${box}" srcOrd="1" destOrd="0"/>`;
    files['ppt/diagrams/data1.xml'] = strToU8(data.replace('<dgm:cxnLst>', `<dgm:cxnLst>${extra}`));
    const e = await open('diagram-editing', 0, zipSync(files));
    const lay = e.session!.layout(0)!;
    const nodes = diagramTextItems(lay, frameOf(e)).map((i) =>
      i.element.kind === 'shape' ? i.element.diagramNode : undefined
    );
    // Three of the four boxes still take the caret; the shared one does not.
    expect(nodes.length).toBe(3);
    // Structure edits refuse: the drawing no longer reads as one shape per node.
    e.setSelection({ kind: 'shapes', nodeIds: [frameOf(e)] });
    expect(e.snapshot().diagram?.structure).toBe(false);
    e.destroy();
  });
});

describe('diagram caret', () => {
  test('the caret names the diagram node; a move to another node is a new caret', async () => {
    const e = await open('diagram-editing');
    caretIn(e, nodeWithText(e, 'Plan'));
    const first = editorCaretOf(e.snapshot(), null);
    expect(first?.diagramNode).toBe(nodeWithText(e, 'Plan'));
    caretIn(e, nodeWithText(e, 'Build'));
    const next = editorCaretOf(e.snapshot(), first);
    expect(next).not.toBe(first);
    expect(next?.diagramNode).toBe(nodeWithText(e, 'Build'));
    e.destroy();
  });
});

describe('diagram text pane', () => {
  test('the snapshot lists the nodes in outline order with their text and level', async () => {
    const e = await open();
    e.setSelection({ kind: 'shapes', nodeIds: [frameOf(e)] });
    const nodes = e.snapshot().diagram!.nodes;
    const top = nodes.filter((n) => n.level === 0).map((n) => n.text);
    expect(top).toContain('Plan');
    expect(top).toContain('Write');
    // A child follows its parent, one level down.
    const write = nodes.findIndex((n) => n.text === 'Write');
    expect(nodes[write + 1]).toMatchObject({ text: 'Draft', level: 1 });
    expect(nodes.every((n) => n.editable)).toBe(true);
    e.destroy();
  });

  test('setDiagramText replaces the text of both parts as one undo step', async () => {
    const e = await open();
    const frame = frameOf(e);
    const plan = nodeWithText(e, 'Plan');
    e.setSelection({ kind: 'shapes', nodeIds: [frame] });
    const r = e.exec({ type: 'setDiagramText', nodeId: frame, node: plan, text: 'Kickoff\nGoals' });
    expect(r.ok).toBe(true);
    const data = partText(e, '/ppt/diagrams/data1.xml');
    const drawing = partText(e, '/ppt/diagrams/drawing1.xml');
    expect(data).toContain('<a:t>Kickoff</a:t>');
    expect(data).toContain('<a:t>Goals</a:t>');
    expect(data).not.toContain('<a:t>Plan</a:t>');
    // The first run's formatting stays on the new text.
    expect(drawing).toContain('<a:rPr lang="en-US" sz="2000" kern="1200"/><a:t>Kickoff</a:t>');
    expect(e.snapshot().diagram!.nodes.find((n) => n.node === plan)?.text).toBe('Kickoff\nGoals');
    // The frame stays selected.
    expect(e.snapshot().selection).toEqual({ kind: 'shapes', nodeIds: [frame] });
    expect(e.exec({ type: 'undo' }).ok).toBe(true);
    expect(partText(e, '/ppt/diagrams/data1.xml')).toContain('<a:t>Plan</a:t>');
    // Empty text leaves the shape empty.
    expect(e.exec({ type: 'setDiagramText', nodeId: frame, node: plan, text: '' }).ok).toBe(true);
    expect(partText(e, '/ppt/diagrams/drawing1.xml')).not.toContain('<a:t>Plan</a:t>');
    e.destroy();
  });

  test('setDiagramText replaces only the changed text', async () => {
    const e = await open();
    const frame = frameOf(e);
    const plan = nodeWithText(e, 'Plan');
    const set = (text: string) =>
      e.exec({ type: 'setDiagramText', nodeId: frame, node: plan, text });
    expect(set('Plan\nGoals').ok).toBe(true);
    // The first paragraph keeps its run when only a second paragraph is new.
    const body = (): string =>
      /<dsp:sp modelId="[^"]+">[\s\S]*?<\/dsp:sp>/.exec(
        partText(e, '/ppt/diagrams/drawing1.xml')
      )![0];
    expect(body()).toContain('<a:t>Plan</a:t>');
    expect(e.exec({ type: 'undo' }).ok).toBe(true);
    expect(set('Plan B').ok).toBe(true);
    // The kept text and the new text share the original run and its size.
    expect(body()).toContain('<a:rPr lang="en-US" sz="2000" kern="1200"/><a:t>Plan B</a:t>');
    const long = set('x'.repeat(10_001));
    expect(long.ok).toBe(false);
    if (!long.ok) expect(long.reason).toBe('errors.diagram.textTooLong');
    e.destroy();
  });

  test('a deleted paragraph leaves the next paragraph its own properties', async () => {
    const files = unzipSync(fixture('diagrams'));
    const more =
      '<a:p><a:pPr algn="r"/><a:r><a:rPr lang="en-US"/><a:t>B</a:t></a:r></a:p>' +
      '<a:p><a:pPr marL="9" algn="l"/><a:r><a:rPr lang="en-US"/><a:t>C</a:t></a:r></a:p>';
    for (const name of ['ppt/diagrams/drawing1.xml', 'ppt/diagrams/data1.xml']) {
      const xml = strFromU8(files[name]!);
      const at = xml.indexOf('</a:p>', xml.indexOf('<a:t>Plan</a:t>')) + '</a:p>'.length;
      files[name] = strToU8(xml.slice(0, at) + more + xml.slice(at));
    }
    const e = await open('diagrams', 0, zipSync(files));
    const frame = frameOf(e);
    e.setSelection({ kind: 'shapes', nodeIds: [frame] });
    const node = e.snapshot().diagram!.nodes[0]!;
    expect(node.text).toBe('Plan\nB\nC');
    const r = e.exec({ type: 'setDiagramText', nodeId: frame, node: node.node, text: 'Plan\nC' });
    expect(r.ok).toBe(true);
    const drawing = partText(e, '/ppt/diagrams/drawing1.xml');
    expect(drawing).toContain('<a:pPr marL="9" algn="l"/><a:r><a:rPr lang="en-US"/><a:t>C</a:t>');
    expect(drawing).not.toContain('<a:t>B</a:t>');
    e.destroy();
  });

  test('a change never splits an emoji or a field', async () => {
    // The first node gets a line break, a bold run, and a date field after its text.
    const files = unzipSync(fixture('diagrams'));
    const extra =
      '<a:br><a:rPr lang="en-US"/></a:br><a:r><a:rPr lang="en-US" b="1"/><a:t>Bold</a:t></a:r>' +
      '<a:fld id="{11111111-1111-4111-8111-111111111111}" type="datetime1"><a:rPr lang="en-US"/>' +
      '<a:t>11/8/2026</a:t></a:fld>';
    for (const name of ['ppt/diagrams/drawing1.xml', 'ppt/diagrams/data1.xml'])
      files[name] = strToU8(
        strFromU8(files[name]!).replace('<a:t>Plan</a:t></a:r>', `<a:t>Plan</a:t></a:r>${extra}`)
      );
    const e = await open('diagrams', 0, zipSync(files));
    const frame = frameOf(e);
    e.setSelection({ kind: 'shapes', nodeIds: [frame] });
    const node = e.snapshot().diagram!.nodes[0]!;
    expect(node.text).toBe('Plan\nBold11/8/2026');
    const set = (text: string) => {
      const command = { type: 'setDiagramText', nodeId: frame, node: node.node, text } as const;
      expect(e.can(command).ok).toBe(true);
      expect(e.exec(command).ok).toBe(true);
      return e.snapshot().diagram!.nodes[0]!.text;
    };
    // An edit before the line break keeps the break and the bold run.
    expect(set('Plans\nBold11/8/2026')).toBe('Plans\nBold11/8/2026');
    expect(partText(e, '/ppt/diagrams/drawing1.xml')).toContain(
      '<a:rPr lang="en-US" b="1"/><a:t>Bold</a:t>'
    );
    // An edit inside the field keeps its text, as plain text.
    expect(set('Plans\nBold11/9/2026')).toBe('Plans\nBold11/9/2026');
    expect(partText(e, '/ppt/diagrams/drawing1.xml')).not.toContain('<a:fld');
    // Emoji that share their first code unit change whole.
    expect(set('x\u{1F600}y')).toBe('x\u{1F600}y');
    expect(set('x\u{1F603}y')).toBe('x\u{1F603}y');
    expect(set('\u{1F600}')).toBe('\u{1F600}');
    e.destroy();
  });

  test('a shape with no paragraph is not editable, and setDiagramText refuses it', async () => {
    const files = unzipSync(fixture('diagram-editing'));
    const drawing = strFromU8(files['ppt/diagrams/drawing1.xml']!);
    // The first shape loses its paragraphs: a text body with none is invalid input.
    const first = /<dsp:txBody>([\s\S]*?)<\/dsp:txBody>/.exec(drawing)!;
    const bare = first[0].replace(/<a:p>[\s\S]*<\/a:p>/, '');
    files['ppt/diagrams/drawing1.xml'] = strToU8(drawing.replace(first[0], bare));
    const e = await open('diagram-editing', 0, zipSync(files));
    const frame = frameOf(e);
    e.setSelection({ kind: 'shapes', nodeIds: [frame] });
    const node = e.snapshot().diagram!.nodes[0]!;
    expect(node.editable).toBe(false);
    const command = { type: 'setDiagramText', nodeId: frame, node: node.node, text: '' } as const;
    expect(e.can(command).ok).toBe(false);
    const r = e.exec(command);
    expect(r.ok).toBe(false);
    if (!r.ok) expect(r.reason).toBe('errors.diagram.noText');
    e.destroy();
  });

  test('setDiagramText refuses an unknown node, unknown keys, and a frame that is no diagram', async () => {
    const e = await open();
    const frame = frameOf(e);
    const bad = e.exec({ type: 'setDiagramText', nodeId: frame, node: 'nope', text: 'x' });
    expect(bad.ok).toBe(false);
    if (!bad.ok) expect(bad.reason).toBe('errors.diagram.noText');
    const extra = e.exec({
      type: 'setDiagramText',
      nodeId: frame,
      node: nodeWithText(e, 'Plan'),
      text: 'x',
      color: 'red',
    } as never);
    expect(extra.ok).toBe(false);
    const other = e.snapshot().shapes!.find((s) => !s.name.startsWith('Diagram'))!.nodeId;
    const noDiagram = e.exec({ type: 'setDiagramText', nodeId: other, node: 'x', text: 'x' });
    expect(noDiagram.ok).toBe(false);
    if (!noDiagram.ok) expect(noDiagram.reason).toBe('errors.diagram.noDiagram');
    e.destroy();
  });

  test('Edit text leads the diagram menu and asks for the Diagram tab', async () => {
    const e = await open('diagram-editing');
    e.setSelection({ kind: 'shapes', nodeIds: [frameOf(e)] });
    expect(CANVAS_CONTEXT_MENUS.diagram[0]).toMatchObject({
      kind: 'item',
      slot: 'diagram.editText',
    });
    // Without a side panel, the row says why it is off.
    expect(toolbarCommandState(e, 'diagram.editText').disabledReason).toBe('disabled.noSidePanel');
    const asked: string[] = [];
    const stop = listenPaneRequests(e, (pane) => asked.push(pane));
    expect(toolbarCommandState(e, 'diagram.editText').enabled).toBe(true);
    expect(runToolbarCommand(e, 'diagram.editText').ok).toBe(true);
    expect(asked).toEqual(['diagram']);
    expect(visibleSidePanelTabs(false, 'format', true)).toContain('diagram');
    expect(visibleSidePanelTabs(false, 'format', false)).not.toContain('diagram');
    expect(visibleSidePanelTabs(false, 'diagram', false)).toContain('diagram');
    // The context menu row asks for the tab too.
    const noop = (): void => undefined;
    const rowHost = {
      openDialog: noop,
      startSlideshow: noop,
      openFind: noop,
      openLink: noop,
      pickImage: noop,
    };
    runChromeRow(e, rowHost, 'diagram.editText');
    expect(asked).toEqual(['diagram', 'diagram']);
    e.setSelection({ kind: 'none' });
    expect(toolbarCommandState(e, 'diagram.editText').disabledReason).toBe('disabled.noDiagram');
    stop();
    e.destroy();
  });
});
