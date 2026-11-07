// Validate saved PresentationML parts against the ECMA-376 Transitional schemas (pml.xsd,
// dml-main.xsd) with xmllint.
//
//   bun scripts/validate-xsd.ts            # fixtures, sample deck, and showcase deck
//   bun scripts/validate-xsd.ts --corpus   # also the private corpus (local only)
//
// Before validation each part goes through Markup Compatibility preprocessing (ECMA-376
// Part 3): ignorable namespaces are dropped and `mc:AlternateContent` keeps its Fallback, as a
// consumer that understands no extension must. For every deck it validates the ORIGINAL parts,
// the parts after save, and the parts after a scripted edit session. The gate: saving and
// editing introduce no schema error the original did not already have.

import { execFileSync } from 'node:child_process';
import { mkdtempSync, readdirSync, readFileSync, rmSync, writeFileSync, existsSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { PNG } from 'pngjs';
import {
  NS,
  readPptx,
  serializeXmlPart,
  writePptx,
  type Element,
  type PptxPackage,
} from '../packages/core/src/store/index.ts';
import { createPptxEditor } from '../packages/core/src/editor/index.ts';

const SCHEMAS = 'reference/ecma-376/part4/transitional-schemas';

const ANIMATION_PRESET_NAMES = [
  'fade',
  'pulse',
  'flyOut',
  'zoom',
  'spin',
  'wipeOut',
  'growShrink',
  'appear',
  'disappear',
  'wipe',
  'fadeOut',
  'flyIn',
  'zoomOut',
] as const;

function preprocess(el: Element, ignorable: Set<string>): Element {
  const children: Element['children'][number][] = [];
  for (const c of el.children) {
    if (c.type === 'text') {
      children.push(c);
      continue;
    }
    if (c.ns === NS.mc && c.local === 'AlternateContent') {
      const fallback = c.children.find((x) => x.type === 'element' && x.local === 'Fallback');
      if (fallback && fallback.type === 'element')
        for (const f of fallback.children)
          if (f.type === 'element') children.push(preprocess(f, ignorable));
      continue;
    }
    if (ignorable.has(c.ns)) continue;
    children.push(preprocess(c, ignorable));
  }
  return {
    ...el,
    attrs: el.attrs.filter((a) => !ignorable.has(a.ns) && !(a.ns === NS.mc)),
    children,
  };
}

function ignorableNamespaces(root: Element): Set<string> {
  const out = new Set<string>();
  const attr = root.attrs.find((a) => a.ns === NS.mc && a.local === 'Ignorable');
  for (const prefix of attr?.value.split(/\s+/) ?? []) {
    const decl = root.nsDecls.find((d) => d.prefix === prefix);
    if (decl) out.add(decl.uri);
  }
  return out;
}

const VALIDATED =
  /^\/ppt\/(slides\/slide|slideLayouts\/slideLayout|slideMasters\/slideMaster|notesSlides\/notesSlide|presentation|theme\/theme)\d*\.xml$/;

/** Validate a package's parts; returns error lines keyed by part, line numbers stripped. */
function validate(pkg: PptxPackage, dir: string, label: string): Map<string, Set<string>> {
  const byPart = new Map<string, Set<string>>();
  const files: { file: string; part: string; schema: string }[] = [];
  for (const part of pkg.parts.values()) {
    if (part.kind !== 'xml' || !VALIDATED.test(part.partName)) continue;
    const root = preprocess(part.xml.root, ignorableNamespaces(part.xml.root));
    const file = join(dir, `${label}-${part.partName.replace(/\//g, '_')}`);
    writeFileSync(file, serializeXmlPart({ ...part.xml, root }));
    files.push({
      file,
      part: part.partName,
      schema: part.partName.includes('/theme/') ? 'dml-main.xsd' : 'pml.xsd',
    });
  }
  for (const schema of ['pml.xsd', 'dml-main.xsd']) {
    const group = files.filter((f) => f.schema === schema);
    if (group.length === 0) continue;
    let out = '';
    try {
      execFileSync(
        'xmllint',
        ['--noout', '--schema', join(SCHEMAS, schema), ...group.map((g) => g.file)],
        { stdio: 'pipe' }
      );
    } catch (e) {
      out = String((e as { stderr?: Buffer }).stderr ?? '');
    }
    for (const line of out.split('\n')) {
      const g = group.find((x) => line.startsWith(`${x.file}:`));
      if (!g || line.endsWith('validates')) continue;
      const message = line
        .slice(g.file.length)
        .replace(/^:\d+:/, '')
        .trim();
      if (!byPart.has(g.part)) byPart.set(g.part, new Set());
      byPart.get(g.part)!.add(message);
    }
  }
  return byPart;
}

/** A `w`×`h` PNG, for the picture edits. */
function png(w: number, h: number): Uint8Array {
  const img = new PNG({ width: w, height: h });
  img.data.fill(160);
  return new Uint8Array(PNG.sync.write(img));
}

/** A scripted edit session through the real editor: the edits the editor supports. */
async function edited(bytes: Uint8Array): Promise<PptxPackage | null> {
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  if (!(await editor.load(bytes)).ok) return null;
  const selected = (): string[] => {
    const sel = editor.snapshot().selection;
    return sel.kind === 'shapes' ? [...sel.nodeIds] : [];
  };
  // Group two new shapes, move one inside the group, and copy the group.
  editor.exec({ type: 'insertShape', preset: 'rect' });
  const [a] = selected();
  editor.exec({ type: 'insertShape', preset: 'ellipse' });
  const [b] = selected();
  // Style the shapes: fill, outline, dash, and text fitting.
  editor.setSelection({ kind: 'shapes', nodeIds: [a, b] });
  editor.exec({ type: 'setShapeFill', color: '#FFC000' });
  editor.exec({ type: 'setLineWidth', points: 3 });
  editor.exec({ type: 'setLineColor', color: '#C00000' });
  editor.exec({ type: 'setLineDash', dash: 'dash' });
  editor.exec({ type: 'setAutofit', mode: 'shape' });
  editor.exec({ type: 'setVerticalAlign', anchor: 'ctr' });
  editor.exec({
    type: 'setShapeGradient',
    gradient: {
      kind: 'linear',
      angle: 45,
      stops: [
        { pos: 0, color: '#FFC000' },
        { pos: 1, color: '#7F3F00' },
      ],
      rotWithShape: true,
    },
  });
  editor.exec({ type: 'group' });
  const [group] = selected();
  editor.setSelection({ kind: 'shapes', nodeIds: [a], group });
  editor.exec({ type: 'nudge', dx: 95_250, dy: 0 });
  editor.setSelection({ kind: 'shapes', nodeIds: [group] });
  const copied = editor.copy();
  // A line and an elbow connector with arrowheads, and a shape from the gallery.
  editor.exec({ type: 'insertLine', kind: 'arrow' });
  editor.exec({ type: 'setLineEnds', head: 'oval', tail: 'stealth' });
  editor.exec({ type: 'insertLine', kind: 'elbow' });
  editor.exec({ type: 'setLineEnds', tail: 'diamond' });
  editor.exec({ type: 'insertShape', preset: 'upArrow' });
  // Shape Format: a gallery style, effects over its theme effect style, a flip, a turn, the
  // aspect lock, and the selection pane's rename and hide.
  editor.exec({ type: 'setShapeStyle', style: { row: 'intenseEffect', color: 'accent2' } });
  editor.exec({ type: 'setShapeEffect', effect: { kind: 'glow', size: 8 } });
  editor.exec({ type: 'setShapeEffect', effect: { kind: 'reflection', preset: 'half4pt' } });
  editor.exec({ type: 'setShapeEffect', effect: { kind: 'softEdge', size: 2.5 } });
  editor.exec({ type: 'setShapeEffect', effect: { kind: 'shadow', preset: 'innerTop' } });
  editor.exec({ type: 'flipObjects', axis: 'horizontal' });
  editor.exec({ type: 'rotateObjects', degrees: 90 });
  editor.exec({ type: 'setLockAspectRatio', lock: true });
  const [arrow] = selected();
  if (arrow) {
    editor.exec({ type: 'renameShape', nodeId: arrow, name: 'Up & <away>' });
    editor.exec({ type: 'setShapesVisible', nodeIds: [arrow], visible: false });
  }
  // A background on every slide, then a picture-free reset on the first.
  editor.exec({
    type: 'setSlideBackground',
    background: { kind: 'solid', color: '#F2F2F2' },
    allSlides: true,
  });
  editor.exec({ type: 'setSlideBackground', background: null });
  // A pasted copy on the same slide, ungrouped.
  if (copied) editor.exec({ type: 'paste', payload: copied });
  editor.exec({ type: 'ungroup' });
  // Picture edits: crop with padding, Fit, crop to shape, corrections, a style, Format
  // Picture, decorative alt text, Change Picture, and Reset Picture & Size.
  editor.exec({
    type: 'insertImage',
    bytes: png(64, 32),
    mime: 'image/png',
    width: 64,
    height: 32,
  });
  editor.exec({
    type: 'cropImage',
    crop: { kind: 'rect', left: 10, top: -5, right: 20, bottom: 0 },
  });
  editor.exec({ type: 'cropImage', crop: { kind: 'fit' } });
  editor.exec({ type: 'cropToShape', preset: 'ellipse' });
  editor.exec({
    type: 'setImageAdjustments',
    adjustments: { brightness: 0.2, contrast: -0.2, recolor: 'sepia', transparency: 0.3 },
  });
  editor.exec({ type: 'setImageStyle', style: 'simpleFrameWhite' });
  editor.exec({ type: 'setImageProperties', widthEmu: 1_828_800, rotationDegrees: 15 });
  editor.exec({ type: 'setAltText', description: 'A chart & <data>', title: 'Sales' });
  editor.exec({ type: 'setAltText', decorative: true });
  editor.exec({
    type: 'insertImage',
    bytes: png(32, 32),
    mime: 'image/png',
    width: 32,
    height: 32,
  });
  editor.exec({ type: 'replaceImage', bytes: png(48, 24), mime: 'image/png' });
  editor.exec({ type: 'resetImage', size: true });
  editor.exec({ type: 'addSlide' });
  // The group pasted onto another slide.
  if (copied) editor.exec({ type: 'paste', payload: copied });
  editor.exec({ type: 'insertTextBox' });
  editor.input('text', 'Edited & <saved>');
  // A link and cleared formatting on the new text.
  const box = editor.snapshot().selection;
  if (box.kind === 'text') {
    editor.setSelection({
      kind: 'text',
      nodeId: box.nodeId,
      anchor: { paragraph: 0, offset: 0 },
      focus: { paragraph: 0, offset: 6 },
    });
    editor.exec({ type: 'setLink', url: 'https://example.com/?a=1&b=<2>' });
    editor.exec({ type: 'clearFormatting' });
    // Lists: the a:pPr children keep the schema order (spacing, bullet color, font, kind). The
    // text box inherits no list, so list off removes the bullet (the a:buNone order is a unit
    // test). No bullet size: the transitional schema reads `buSzPct` as a percent string, and
    // saved decks hold 1/1000ths.
    editor.exec({ type: 'setLineSpacing', multiple: 1.5 });
    editor.exec({
      type: 'setBulletFormat',
      format: { kind: 'bullet', char: '§', font: 'Wingdings' },
    });
    editor.exec({
      type: 'setBulletFormat',
      format: { kind: 'number', scheme: 'romanLcParenR', startAt: 3, color: '#C00000' },
    });
    editor.exec({ type: 'setBulletFormat', format: { kind: 'bullet', color: null } });
    editor.exec({ type: 'toggleBullets' });
  }
  editor.exec({ type: 'toggleBold' });
  editor.exec({ type: 'insertShape', preset: 'roundRect' });
  // Animations: every preset, a reorder, timing, a removal, and a transition.
  for (const preset of ANIMATION_PRESET_NAMES) editor.exec({ type: 'addAnimation', preset });
  const effects = editor.snapshot().animations?.effects ?? [];
  if (effects.length > 3) {
    editor.exec({ type: 'moveAnimation', effectId: effects[1].id, to: 0 });
    editor.exec({
      type: 'setAnimationTiming',
      effectId: effects[2].id,
      trigger: 'afterPrevious',
      delayMs: 250,
      durationMs: 1500,
    });
    editor.exec({ type: 'removeAnimation', effectId: effects[3].id });
  }
  editor.exec({ type: 'setTransition', kind: 'split', durationMs: 1250, advanceAfterMs: 4000 });
  editor.exec({ type: 'nudge', dx: 9525, dy: 9525 });
  editor.exec({ type: 'duplicateSlide', index: 0 });
  editor.exec({ type: 'setNotes', index: 0, text: 'Edited notes' });
  editor.exec({ type: 'moveSlide', from: 0, to: 1 });
  // A delete updates sections and custom shows and removes the parts only that slide used.
  editor.exec({ type: 'duplicateSlide', index: 0 });
  editor.exec({ type: 'deleteSlide', index: 0 });
  // A table: text in cells, rows and columns, merge and split, and table and cell formatting.
  editor.exec({ type: 'insertTable', rows: 3, cols: 3 });
  const cellSel = editor.snapshot().selection;
  if (cellSel.kind === 'text' && cellSel.cell) {
    const nodeId = cellSel.nodeId;
    editor.input('text', 'Cell & <text>');
    editor.exec({ type: 'insertRow', where: 'below' });
    editor.exec({ type: 'insertColumn', where: 'right' });
    editor.setSelection({
      kind: 'cells',
      nodeId,
      anchor: { row: 1, column: 0 },
      focus: { row: 2, column: 1 },
    });
    editor.exec({ type: 'mergeCells' });
    editor.exec({ type: 'setCellFill', color: { theme: 'accent2', shift: 40 } });
    editor.exec({
      type: 'setTableBorders',
      scope: 'outside',
      spec: { color: '#C00000', width: 2, dash: 'dash' },
    });
    editor.exec({ type: 'setCellMargins', margins: { left: 0, top: 91_440 } });
    editor.exec({ type: 'setVerticalAlign', anchor: 'b' });
    editor.setSelection({ ...cellSel, cell: { row: 3, column: 3 } });
    editor.exec({ type: 'splitCell', rows: 2, cols: 2 });
    editor.exec({ type: 'setTableBorders', scope: 'none' });
    editor.exec({ type: 'setCellFill', color: null });
    editor.exec({ type: 'toggleTableStyleOption', option: 'lastRow' });
    editor.exec({ type: 'distributeColumns' });
    editor.exec({ type: 'setRowHeight', height: 457_200 });
    editor.exec({ type: 'deleteColumn' });
  }
  const r = readPptx(editor.save()!);
  return r.ok ? r.pkg : null;
}

async function main(): Promise<void> {
  const decks: { name: string; file: string }[] = [];
  for (const f of readdirSync('fixtures')
    .filter((x) => x.endsWith('.pptx'))
    .sort())
    decks.push({ name: f.replace('.pptx', ''), file: join('fixtures', f) });
  decks.push({ name: 'sample', file: 'examples/vite/public/sample.pptx' });
  decks.push({ name: 'showcase', file: 'examples/vite/public/showcase.pptx' });
  if (process.argv.includes('--corpus') && existsSync('demo-presentations'))
    readdirSync('demo-presentations')
      .filter((x) => x.endsWith('.pptx'))
      .sort()
      .forEach((f, i) => decks.push({ name: `deck${i + 1}`, file: join('demo-presentations', f) }));
  const dir = mkdtempSync(join(tmpdir(), 'pptx-xsd-'));
  let failed = false;
  try {
    for (const deck of decks) {
      const bytes = new Uint8Array(readFileSync(deck.file));
      const original = readPptx(bytes);
      if (!original.ok) {
        console.log(`${deck.name}: unreadable`);
        failed = true;
        continue;
      }
      const saved = readPptx(writePptx(original.pkg));
      const edit = await edited(bytes);
      const base = validate(original.pkg, dir, 'o');
      const results: [string, Map<string, Set<string>> | null][] = [
        ['saved', saved.ok ? validate(saved.pkg, dir, 's') : null],
        ['edited', edit ? validate(edit, dir, 'e') : null],
      ];
      const baseErrors = [...base.values()].reduce((n, s) => n + s.size, 0);
      const summary: string[] = [`original ${baseErrors} error(s)`];
      for (const [label, result] of results) {
        if (!result) {
          summary.push(`${label}: unreadable`);
          failed = true;
          continue;
        }
        const allBase = new Set([...base.values()].flatMap((s) => [...s]));
        const fresh = [...result.values()].flatMap((s) => [...s]).filter((m) => !allBase.has(m));
        summary.push(`${label}: ${fresh.length} new`);
        if (fresh.length) {
          failed = true;
          for (const m of fresh.slice(0, 5)) console.log(`  ${deck.name} ${label}: ${m}`);
        }
      }
      console.log(`${deck.name}: ${summary.join(', ')}`);
    }
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
  process.exit(failed ? 1 : 0);
}

await main();
