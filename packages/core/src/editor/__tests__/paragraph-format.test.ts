// The Paragraph dialog's command (`setParagraphFormat`): spacing before and after, indents, line
// spacing, and alignment. Each value is written in its schema unit, in schema order, lays
// out, undoes as one step, and survives save and reopen.

import { describe, expect, spyOn, test } from 'bun:test';
import { buildPptx, para, run, shape } from '../../../../../scripts/fixtures/builder.ts';
import {
  attr,
  child,
  elements,
  find,
  readPptx,
  xmlPartOf,
  type Element,
} from '../../store/index.ts';
import {
  paragraphCommandOf,
  paragraphFieldError,
  seedParagraphFields,
  withSpaceUnit,
} from '../paragraph-dialog-fields.ts';
import { resetDevWarnings } from '../dev-warnings.ts';
import {
  CANVAS_CONTEXT_MENUS,
  CHROME_MENUS,
  chromeRowState,
  commandForSlotValue,
  createPptxEditor,
  editorCommandFor,
  runChromeRow,
  formattingBarChromeGroups,
  isToolbarBarSlot,
  runToolbarCommand,
  toolbarCommandState,
  type PptxEditorInstance,
} from '../index.ts';

const IN = 914_400;
const NONE = {
  alignment: false,
  spaceBefore: false,
  spaceAfter: false,
  lineSpacing: false,
  indentLeft: false,
  indentFirstLine: false,
};
const BODY = '<p:ph type="body" idx="1"/>';

function deck(text: string, ph?: string): Uint8Array {
  return buildPptx({
    slides: [
      {
        shapes: shape({ id: 2, name: 'T', x: IN, y: IN, cx: 6 * IN, cy: 4 * IN, text, ph }),
      },
    ],
  });
}

async function open(bytes: Uint8Array): Promise<PptxEditorInstance> {
  const e = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await e.load(bytes)).ok).toBe(true);
  return e;
}

const root = (e: PptxEditorInstance): Element =>
  xmlPartOf(e.session!.store.package, e.session!.info.slides[0].partName)!.root;
const shapeId = (e: PptxEditorInstance): string =>
  elements(find(root(e), 'p:spTree')!).find((c) => attr(find(c, 'p:cNvPr'), 'name') === 'T')!.id;
const paragraphs = (e: PptxEditorInstance): Element[] =>
  elements(find(root(e), 'p:txBody')!).filter((c) => c.local === 'p');
const pPr = (e: PptxEditorInstance, i: number): Element | undefined =>
  child(paragraphs(e)[i], 'a:pPr');
const childNames = (el: Element | undefined): string[] => elements(el).map((c) => c.local);
const lines = (e: PptxEditorInstance) => {
  const item = e.session!.layout(0)!.items.find((x) => x.kind === 'shape' && x.text);
  return item?.kind === 'shape' ? item.text!.lines : [];
};
const select = (e: PptxEditorInstance, from: number, to = from) =>
  e.setSelection({
    kind: 'text',
    nodeId: shapeId(e),
    anchor: { paragraph: from, offset: 0 },
    focus: { paragraph: to, offset: 1 },
  });

describe('setParagraphFormat', () => {
  test('writes space before and after as a:spcPts and a:spcPct, in schema order', async () => {
    const e = await open(
      deck(
        para(run('one')) +
          para(
            run('two'),
            '<a:pPr><a:lnSpc><a:spcPct val="150000"/></a:lnSpc><a:buFont typeface="Arial"/><a:buChar char="•"/></a:pPr>'
          )
      )
    );
    select(e, 1);
    const r = e.exec({ type: 'setParagraphFormat', spaceBeforePt: 12, spaceAfterLines: 0.5 });
    expect(r.ok).toBe(true);
    const p = pPr(e, 1)!;
    expect(childNames(p)).toEqual(['lnSpc', 'spcBef', 'spcAft', 'buFont', 'buChar']);
    expect(attr(find(p, 'a:spcBef')!.children[0] as Element, 'val')).toBe('1200');
    expect(find(p, 'a:spcAft')!.children.map((c) => (c as Element).local)).toEqual(['spcPct']);
    expect(attr(find(find(p, 'a:spcAft')!, 'a:spcPct'), 'val')).toBe('50000');
    // The first paragraph is outside the selection and keeps no pPr.
    expect(pPr(e, 0)).toBeUndefined();
  });

  test('space before moves the paragraph down by that many points', async () => {
    const e = await open(deck(para(run('one')) + para(run('two'))));
    const before = lines(e)[1]!.top;
    select(e, 1);
    expect(e.exec({ type: 'setParagraphFormat', spaceBeforePt: 24 }).ok).toBe(true);
    expect(lines(e)[1]!.top - before).toBe(24 * 12_700);
    expect(e.snapshot().format?.paragraph?.spaceBeforePt).toBe(24);
  });

  test('one command is one undo step', async () => {
    const e = await open(deck(para(run('one')) + para(run('two'))));
    select(e, 0, 1);
    expect(
      e.exec({
        type: 'setParagraphFormat',
        alignment: 'center',
        indentLeftEmu: IN / 2,
        indentFirstLineEmu: -IN / 4,
        spaceBeforePt: 6,
        spaceAfterPt: 6,
        lineSpacing: { rule: 'exact', value: 30 },
      }).ok
    ).toBe(true);
    for (const i of [0, 1]) {
      const p = pPr(e, i)!;
      expect(attr(p, 'algn')).toBe('ctr');
      expect(attr(p, 'marL')).toBe(String(IN / 2));
      expect(attr(p, 'indent')).toBe(String(-IN / 4));
      expect(attr(find(find(p, 'a:lnSpc')!, 'a:spcPts'), 'val')).toBe('3000');
    }
    e.exec({ type: 'undo' });
    expect(pPr(e, 0)).toBeUndefined();
    expect(pPr(e, 1)).toBeUndefined();
    expect(e.snapshot().canUndo).toBe(false);
  });

  test('null removes the paragraph value, so the inherited value shows again', async () => {
    const e = await open(
      deck(
        para(
          run('item'),
          '<a:pPr marL="0" indent="0"><a:spcBef><a:spcPts val="0"/></a:spcBef></a:pPr>'
        ),
        BODY
      )
    );
    select(e, 0);
    expect(e.snapshot().format?.paragraph?.indentLeftEmu).toBe(0);
    expect(
      e.exec({
        type: 'setParagraphFormat',
        indentLeftEmu: null,
        indentFirstLineEmu: null,
        spaceBeforePt: null,
      }).ok
    ).toBe(true);
    const p = pPr(e, 0)!;
    expect(attr(p, 'marL')).toBeUndefined();
    expect(find(p, 'a:spcBef')).toBeUndefined();
    // The master body style of the fixture's builder states a level-1 margin and spacing.
    const state = e.snapshot().format!.paragraph!;
    expect(state.indentLeftEmu).toBeGreaterThan(0);
    expect(state.indentFirstLineEmu).toBeLessThan(0);
  });

  test('the snapshot shows the resolved paragraph', async () => {
    const e = await open(
      deck(
        para(
          run('x'),
          '<a:pPr algn="r" marL="457200" indent="228600"><a:lnSpc><a:spcPct val="120000"/></a:lnSpc><a:spcAft><a:spcPct val="25000"/></a:spcAft></a:pPr>'
        )
      )
    );
    select(e, 0);
    expect(e.snapshot().format?.paragraph).toEqual({
      alignment: 'right',
      indentLeftEmu: 457_200,
      indentFirstLineEmu: 228_600,
      spaceBeforePt: 0,
      spaceBeforeLines: null,
      spaceAfterPt: null,
      spaceAfterLines: 0.25,
      lineSpacing: { rule: 'multiple', value: 1.2 },
      disagrees: NONE,
    });
  });

  test('refuses values the schema does not allow, and an empty command', async () => {
    const e = await open(deck(para(run('one'))));
    select(e, 0);
    const bad = [
      {},
      { alignment: 'dist' },
      { indentLeftEmu: -1 },
      { indentLeftEmu: 60_000_000 },
      { indentFirstLineEmu: Number.NaN },
      { spaceBeforePt: -1 },
      { spaceBeforePt: 1585 },
      { spaceAfterLines: 200 },
      { spaceAfterPt: 1, spaceAfterLines: 1 },
      // A space given in both units, even with one of them undefined, is refused.
      { spaceBeforePt: 6, spaceBeforeLines: null },
      { spaceBeforePt: Number.NaN },
      { lineSpacing: { rule: 'multiple', value: 0 } },
      { lineSpacing: { rule: 'multiple', value: 10 } },
      { lineSpacing: { rule: 'exact', value: 2000 } },
      { lineSpacing: { rule: 'atLeast', value: 12 } },
      { lineSpacing: { rule: 'exactly', value: 12 } },
      // A value that rounds to zero would collapse the lines.
      { lineSpacing: { rule: 'exact', value: 0.004 } },
      { lineSpacing: { rule: 'multiple', value: 0.000001 } },
      { lineSpacing: { rule: 'multiple', value: 0.005 } },
    ];
    for (const fields of bad) {
      const r = e.exec({ type: 'setParagraphFormat', ...fields } as never);
      expect({ fields, ok: r.ok }).toEqual({ fields, ok: false });
    }
    expect(e.snapshot().canUndo).toBe(false);
  });

  test('needs a shape with text', async () => {
    const e = await open(deck(para(run('one'))));
    expect(e.can({ type: 'setParagraphFormat', alignment: 'left' })).toEqual({
      ok: false,
      reason: 'disabled.noSelection',
    });
  });

  test('the values survive save and reopen', async () => {
    const e = await open(deck(para(run('one')) + para(run('two'))));
    select(e, 0, 1);
    e.exec({
      type: 'setParagraphFormat',
      spaceBeforePt: 10.5,
      spaceAfterLines: 1.5,
      lineSpacing: { rule: 'multiple', value: 0.9 },
      indentLeftEmu: 342_900,
      indentFirstLineEmu: -342_900,
    });
    const saved = e.save()!;
    expect(readPptx(saved).ok).toBe(true);
    const again = await open(saved);
    again.setSelection({
      kind: 'text',
      nodeId: shapeId(again),
      anchor: { paragraph: 1, offset: 0 },
      focus: { paragraph: 1, offset: 0 },
    });
    expect(again.snapshot().format?.paragraph).toEqual({
      alignment: 'left',
      indentLeftEmu: 342_900,
      indentFirstLineEmu: -342_900,
      spaceBeforePt: 10.5,
      spaceBeforeLines: null,
      spaceAfterPt: null,
      spaceAfterLines: 1.5,
      lineSpacing: { rule: 'multiple', value: 0.9 },
      disagrees: NONE,
    });
  });
});

describe('paragraph.dialog slot', () => {
  test('is enabled for text and applies the dialog value as one command', async () => {
    const e = await open(deck(para(run('one'))));
    expect(toolbarCommandState(e, 'paragraph.dialog').enabled).toBe(false);
    select(e, 0);
    expect(toolbarCommandState(e, 'paragraph.dialog').enabled).toBe(true);
    expect(runToolbarCommand(e, 'paragraph.dialog', { spaceAfterPt: 18 }).ok).toBe(true);
    expect(attr(find(find(pPr(e, 0)!, 'a:spcAft')!, 'a:spcPts'), 'val')).toBe('1800');
    // A value-less run has no fixed command.
    expect(runToolbarCommand(e, 'paragraph.dialog').ok).toBe(false);
  });

  test("a value's own type key never picks another command", async () => {
    const e = await open(deck(para(run('one')) + para(run('two'))));
    select(e, 0);
    expect(commandForSlotValue('paragraph.dialog', { type: 'delete', spaceAfterPt: 6 })).toEqual({
      type: 'setParagraphFormat',
      spaceAfterPt: 6,
    });
    runToolbarCommand(e, 'paragraph.dialog', { type: 'delete', spaceAfterPt: 6 } as never);
    expect(paragraphs(e).length).toBe(2);
  });

  test('Format > Paragraph… and the text context menu list it; a row opens the dialog', async () => {
    const format = CHROME_MENUS.find((m) => m.id === 'format')!.entries;
    expect(format.some((r) => r.kind === 'item' && r.slot === 'paragraph.dialog')).toBe(true);
    expect(format.some((r) => r.kind === 'item' && r.slot === 'text.dialog')).toBe(true);
    expect(
      CANVAS_CONTEXT_MENUS.text.some((r) => r.kind === 'item' && r.slot === 'text.dialog')
    ).toBe(true);
    expect(
      CANVAS_CONTEXT_MENUS.text.some((r) => r.kind === 'item' && r.slot === 'paragraph.dialog')
    ).toBe(true);
    const e = await open(deck(para(run('one'))));
    select(e, 0);
    const opened: string[] = [];
    const host = {
      openDialog: (key: string) => void opened.push(key),
      startSlideshow: () => {},
      openFind: () => {},
      openLink: () => {},
      pickImage: () => {},
    };
    runChromeRow(e, host, 'paragraph.dialog');
    expect(opened).toEqual(['paragraph']);
    // `execute(fields)` writes the fields from core; `execute()` opens the dialog.
    const cmd = editorCommandFor(e, 'paragraph.dialog', host, 'disabled.noRoot');
    expect(cmd.state.isEnabled).toBe(true);
    expect(cmd.execute({ spaceBeforePt: 12 })).toBe(true);
    expect(attr(find(find(pPr(e, 0)!, 'a:spcBef')!, 'a:spcPts'), 'val')).toBe('1200');
    expect(opened).toEqual(['paragraph']);
    expect(cmd.execute()).toBe(true);
    expect(opened).toEqual(['paragraph', 'paragraph']);
    // A dialog row with a value (Format > Bullets and numbering…) opens its dialog and takes the
    // slot's state.
    expect(
      format.some((r) => r.kind === 'item' && r.slot === 'list.bulletStyle' && r.value === 'dialog')
    ).toBe(true);
    runChromeRow(e, host, 'list.bulletStyle', 'dialog');
    expect(opened.at(-1)).toBe('bulletsAndNumbering');
    expect(chromeRowState(e, 'list.bulletStyle', 'dialog').isEnabled).toBe(true);
    // The row asks the host about its own dialog.
    const asked: (string | undefined)[] = [];
    const offRow = { ...host, unavailable: (_: string, d?: string) => (asked.push(d), null) };
    chromeRowState(e, 'list.bulletStyle', 'dialog', offRow);
    expect(asked).toEqual(['bulletsAndNumbering']);
    // A row turned off by the host shows the host's reason.
    const off = { ...host, unavailable: () => 'dialog.unavailable' };
    expect(chromeRowState(e, 'paragraph.dialog', undefined, off).disabledReason).toBe(
      'dialog.unavailable'
    );
  });

  test('is a contextual slot: menus open it, the formatting bar does not draw it', () => {
    const groups = formattingBarChromeGroups();
    expect(groups.find((g) => g.id === 'paragraph')!.controls.map((c) => c.id)).not.toContain(
      'dialog'
    );
    expect(isToolbarBarSlot('paragraph.dialog')).toBe(false);
    const spacing = groups
      .find((g) => g.id === 'list')!
      .controls.find((c) => c.id === 'lineSpacing')!;
    expect(spacing.options!.at(-1)).toMatchObject({ dialog: 'paragraph' });
  });
});

describe('mixed selections', () => {
  test('a field the paragraphs disagree about reads null, and one value sets them all', async () => {
    const e = await open(
      deck(
        para(run('one'), '<a:pPr algn="ctr"><a:spcBef><a:spcPts val="600"/></a:spcBef></a:pPr>') +
          para(run('two'))
      )
    );
    select(e, 0, 1);
    const read = e.snapshot().format!.paragraph!;
    expect(read.spaceBeforePt).toBeNull();
    expect(read.alignment).toBeNull();
    expect(read.disagrees).toMatchObject({
      spaceBefore: true,
      alignment: true,
      lineSpacing: false,
    });
    expect(read.lineSpacing).not.toBeNull();
    expect(e.exec({ type: 'setParagraphFormat', spaceBeforePt: 6 }).ok).toBe(true);
    expect(e.snapshot().format!.paragraph!.spaceBeforePt).toBe(6);
    expect(e.snapshot().format!.paragraph!.disagrees.spaceBefore).toBe(false);
  });

  test('Distributed reads as distributed and writes algn="dist"', async () => {
    const e = await open(deck(para(run('one'), '<a:pPr algn="dist"/>') + para(run('two'))));
    select(e, 0);
    expect(e.snapshot().format!.paragraph!.alignment).toBe('distributed');
    select(e, 1);
    expect(e.exec({ type: 'setParagraphFormat', alignment: 'distributed' }).ok).toBe(true);
    expect(attr(pPr(e, 1), 'algn')).toBe('dist');
  });
});

describe('collaboration-friendly writes', () => {
  test('a new value of the same kind changes val in place', async () => {
    const e = await open(
      deck(para(run('one'), '<a:pPr><a:spcBef><a:spcPts val="600"/></a:spcBef></a:pPr>'))
    );
    select(e, 0);
    const spcPts = find(find(pPr(e, 0)!, 'a:spcBef')!, 'a:spcPts')!;
    expect(e.exec({ type: 'setParagraphFormat', spaceBeforePt: 12 }).ok).toBe(true);
    const after = find(find(pPr(e, 0)!, 'a:spcBef')!, 'a:spcPts')!;
    expect(after.id).toBe(spcPts.id);
    expect(attr(after, 'val')).toBe('1200');
    // A change of kind replaces the child.
    expect(e.exec({ type: 'setParagraphFormat', spaceBeforeLines: 1 }).ok).toBe(true);
    expect(childNames(find(pPr(e, 0)!, 'a:spcBef'))).toEqual(['spcPct']);
  });

  test('a command that only removes values adds no a:pPr', async () => {
    const e = await open(deck(para(run('one'))));
    select(e, 0);
    expect(e.exec({ type: 'setParagraphFormat', spaceBeforePt: null }).ok).toBe(true);
    expect(pPr(e, 0)).toBeUndefined();
  });
});

describe('reads across shapes', () => {
  test('several selected shapes: the read covers every shape the command writes', async () => {
    const e = await open(
      buildPptx({
        slides: [
          {
            shapes:
              shape({ id: 2, name: 'A', x: IN, y: IN, cx: 3 * IN, cy: IN, text: para(run('a')) }) +
              shape({
                id: 3,
                name: 'B',
                x: IN,
                y: 3 * IN,
                cx: 3 * IN,
                cy: IN,
                text: para(
                  run('b'),
                  '<a:pPr algn="ctr"><a:spcBef><a:spcPts val="1800"/></a:spcBef></a:pPr>'
                ),
              }),
          },
        ],
      })
    );
    const ids = elements(find(root(e), 'p:spTree')!)
      .filter((c) => c.local === 'sp')
      .map((c) => c.id);
    e.setSelection({ kind: 'shapes', nodeIds: ids });
    const read = e.snapshot().format!.paragraph!;
    expect(read.disagrees).toMatchObject({ alignment: true, spaceBefore: true });
    expect(read.alignment).toBeNull();
  });
});

describe('the dialog slot run as a command', () => {
  test('says it opens a dialog, not "not available"', async () => {
    const e = await open(deck(para(run('one'))));
    select(e, 0);
    expect(runToolbarCommand(e, 'paragraph.dialog')).toEqual({
      ok: false,
      reason: 'disabled.dialogSlot',
    });
  });
});

describe('dialog fields', () => {
  test('nothing read seeds every field blank, and OK then writes only what was typed', () => {
    const seed = seedParagraphFields(null);
    expect(Object.values(seed).filter((v) => v === '').length).toBe(8);
    expect(paragraphCommandOf(seed, { ...seed, spaceAfter: '6' })).toEqual({ spaceAfterPt: 6 });
  });

  test('a file value outside the dialog limits never blocks OK on another field', () => {
    const seed = seedParagraphFields({
      alignment: 'left',
      spaceBeforePt: 0,
      spaceBeforeLines: null,
      spaceAfterPt: 0,
      spaceAfterLines: null,
      lineSpacing: { rule: 'multiple', value: 12 },
      indentLeftEmu: 0,
      indentFirstLineEmu: 0,
      disagrees: NONE,
    });
    const current = { ...seed, alignment: 'center' as const };
    expect(paragraphFieldError(seed, current)).toBeNull();
    expect(paragraphCommandOf(seed, current)).toEqual({ alignment: 'center' });
    // A changed At is checked, with 0.01 as the smallest value.
    expect(
      paragraphFieldError(seed, { ...seed, lineRule: 'exact', lineAt: '0.001' })
    ).toMatchObject({ field: 'lineAt', params: { min: 0.01 } });
  });
});

describe('review round 3', () => {
  test('the Line spacing options… row run as a value says it opens a dialog', async () => {
    const e = await open(deck(para(run('one'))));
    select(e, 0);
    expect(runToolbarCommand(e, 'list.lineSpacing', 'dialog' as never)).toEqual({
      ok: false,
      reason: 'disabled.dialogSlot',
    });
  });

  test('can() keeps a reset (null) apart from a non-number', async () => {
    const e = await open(deck(para(run('one'))));
    select(e, 0);
    expect(e.can({ type: 'setParagraphFormat', indentLeftEmu: null }).ok).toBe(true);
    expect(e.can({ type: 'setParagraphFormat', indentLeftEmu: Number.NaN }).ok).toBe(false);
    expect(e.can({ type: 'setParagraphFormat', spaceBeforePt: Number.POSITIVE_INFINITY }).ok).toBe(
      false
    );
  });

  test('a shape without text first in the selection does not hide the text shapes', async () => {
    const e = await open(
      buildPptx({
        slides: [
          {
            shapes:
              shape({ id: 2, name: 'Empty', x: IN, y: IN, cx: IN, cy: IN }) +
              shape({
                id: 3,
                name: 'T',
                x: 3 * IN,
                y: IN,
                cx: 3 * IN,
                cy: IN,
                text: para(run('t')),
              }),
          },
        ],
      })
    );
    const ids = elements(find(root(e), 'p:spTree')!)
      .filter((c) => c.local === 'sp')
      .map((c) => c.id);
    e.setSelection({ kind: 'shapes', nodeIds: ids });
    expect(e.snapshot().format?.paragraph?.alignment).toBe('left');
  });

  test('the toolbar line spacing changes val in place', async () => {
    const e = await open(
      deck(para(run('one'), '<a:pPr><a:lnSpc><a:spcPct val="100000"/></a:lnSpc></a:pPr>'))
    );
    select(e, 0);
    const before = find(find(pPr(e, 0)!, 'a:lnSpc')!, 'a:spcPct')!;
    expect(e.exec({ type: 'setLineSpacing', multiple: 1.5 }).ok).toBe(true);
    const after = find(find(pPr(e, 0)!, 'a:lnSpc')!, 'a:spcPct')!;
    expect([after.id, attr(after, 'val')]).toEqual([before.id, '150000']);
  });

  test('a unit change keeps a typed number and resets the opened one', () => {
    const seed = seedParagraphFields(null);
    const opened = { ...seed, spaceBefore: '12' };
    expect(withSpaceUnit(opened, opened, 'spaceBefore', 'lines').spaceBefore).toBe('0');
    const typed = { ...opened, spaceBefore: '0.5' };
    expect(withSpaceUnit(typed, opened, 'spaceBefore', 'lines')).toMatchObject({
      spaceBefore: '0.5',
      spaceBeforeUnit: 'lines',
    });
    // A blank stays blank and needs no check.
    const blank = withSpaceUnit(seed, seed, 'spaceBefore', 'lines');
    expect(blank.spaceBefore).toBe('');
    expect(paragraphFieldError(seed, blank)).toBeNull();
  });
});

describe('review round 4', () => {
  test('a menu row value run through its slot names the row in the warning', async () => {
    const e = await open(deck(para(run('one'))));
    select(e, 0);
    resetDevWarnings();
    const warn = spyOn(console, 'warn').mockImplementation(() => {});
    runToolbarCommand(e, 'list.lineSpacing', 'dialog' as never);
    expect(String(warn.mock.calls[0]?.[0])).toContain('`list.lineSpacing`\'s row "dialog"');
    warn.mockRestore();
  });

  test('the toolbar line spacing has the dialog range', async () => {
    const e = await open(deck(para(run('one'))));
    select(e, 0);
    for (const multiple of [10, 0.005, 0])
      expect(e.exec({ type: 'setLineSpacing', multiple }).ok).toBe(false);
    expect(e.exec({ type: 'setLineSpacing', multiple: 9.99 }).ok).toBe(true);
  });

  test('can() keeps a non-number apart from a string that looks like its key', async () => {
    const e = await open(deck(para(run('one'))));
    expect(e.can({ type: 'setNotes', index: 0, text: Number.NaN as never }).ok).toBe(false);
    expect(e.can({ type: 'setNotes', index: 0, text: '#NaN' }).ok).toBe(true);
  });

  test('back to the opened unit, the opened value returns', () => {
    const seed = { ...seedParagraphFields(null), spaceBefore: '12' };
    const lines = withSpaceUnit(seed, seed, 'spaceBefore', 'lines');
    expect(lines.spaceBefore).toBe('0');
    expect(withSpaceUnit(lines, seed, 'spaceBefore', 'pt').spaceBefore).toBe('12');
  });
});
