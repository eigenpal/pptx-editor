/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/

// The two-peer convergence fixtures of the collaboration coverage gate, one per editing command
// and one per kind of text input. Each prepares peer A on the coverage deck
// (`convergence-harness.ts`) and builds the edit. The edit must change the deck: a fixture where
// the command is a no-op fails. A command without a fixture here needs a reason in
// `COLLABORATION_UNCOVERED` and a probe in `NO_WRITE_PROBES`. CONTRIBUTING.md, "Add a
// collaboration fixture", gives the steps.

import { expect } from 'bun:test';
import type {
  EditorCommandType,
  EditorSelection,
  PptxEditorInstance,
  TextInputKind,
} from '@pptx-editor.dev/core/editor';
import { png, type ConvergenceFixture, type TextInputFixture } from './convergence-harness.ts';

const IN = 914_400;

/** The node id of an element of the current slide, by `p:cNvPr/@name`. */
function idOf(e: PptxEditorInstance, name: string): string {
  const found = e.snapshot().shapes?.find((s) => s.name === name);
  if (!found) throw new Error(`no element named ${name}`);
  return found.nodeId;
}

const TABLE = 'Table 5';
const PICTURE = 'Picture';

/** Select elements of the current slide by name. */
const select =
  (...names: string[]) =>
  (e: PptxEditorInstance): void =>
    e.setSelection({ kind: 'shapes', nodeIds: names.map((n) => idOf(e, n)) });

/** Select characters `[start, end)` of the first paragraph of shape `name`. */
const selectText =
  (name: string, start: number, end: number) =>
  (e: PptxEditorInstance): void =>
    e.setSelection({
      kind: 'text',
      nodeId: idOf(e, name),
      anchor: { paragraph: 0, offset: start },
      focus: { paragraph: 0, offset: end },
    });

/** A caret in a table cell, at `offset` in its first paragraph. */
function cellCaret(
  e: PptxEditorInstance,
  row: number,
  column: number,
  offset = 0
): EditorSelection {
  return {
    kind: 'text',
    nodeId: idOf(e, TABLE),
    cell: { row, column },
    anchor: { paragraph: 0, offset },
    focus: { paragraph: 0, offset },
  };
}

/** A range of table cells. */
function cellRange(
  e: PptxEditorInstance,
  from: [number, number],
  to: [number, number]
): EditorSelection {
  return {
    kind: 'cells',
    nodeId: idOf(e, TABLE),
    anchor: { row: from[0], column: from[1] },
    focus: { row: to[0], column: to[1] },
  };
}

const atCell =
  (row: number, column: number) =>
  (e: PptxEditorInstance): void =>
    e.setSelection(cellCaret(e, row, column));

/** Run a command during `arrange`; it must succeed. */
function run(e: PptxEditorInstance, command: Parameters<PptxEditorInstance['exec']>[0]): void {
  expect(e.exec(command)).toEqual({ ok: true });
}

/** Build a fixture for a command that acts on shape `A`. */
const onA = (command: ConvergenceFixture['command']): ConvergenceFixture => ({
  arrange: select('A'),
  command,
});

const sectionId = (e: PptxEditorInstance, i: number): string => e.snapshot().sections![i]!.id;

/** The fixtures, by command type. The gate checks that each builds a command of its key. */
export const CONVERGENCE_FIXTURES: Partial<Record<EditorCommandType, ConvergenceFixture>> = {
  // History: undo and redo replicate the step they reverse.
  undo: {
    arrange: (e) => {
      select('A')(e);
      run(e, { type: 'delete' });
    },
    command: () => ({ type: 'undo' }),
  },
  redo: {
    arrange: (e) => {
      select('A')(e);
      run(e, { type: 'delete' });
      run(e, { type: 'undo' });
    },
    command: () => ({ type: 'redo' }),
  },

  // Character and paragraph formatting.
  toggleBold: { arrange: selectText('A', 0, 5), command: () => ({ type: 'toggleBold' }) },
  toggleItalic: onA(() => ({ type: 'toggleItalic' })),
  toggleUnderline: onA(() => ({ type: 'toggleUnderline' })),
  toggleStrike: onA(() => ({ type: 'toggleStrike' })),
  setFontSize: onA(() => ({ type: 'setFontSize', points: 32 })),
  stepFontSize: onA(() => ({ type: 'stepFontSize', direction: 1 })),
  setFontFamily: onA(() => ({ type: 'setFontFamily', family: 'Georgia' })),
  setTextColor: onA(() => ({ type: 'setTextColor', color: { theme: 'accent2' } })),
  setHighlight: {
    arrange: selectText('A', 0, 5),
    command: () => ({ type: 'setHighlight', color: '#FFFF00' }),
  },
  setAlign: onA(() => ({ type: 'setAlign', align: 'ctr' })),
  toggleBullets: onA(() => ({ type: 'toggleBullets' })),
  toggleNumbering: onA(() => ({ type: 'toggleNumbering' })),
  indent: {
    arrange: (e) => {
      select('A')(e);
      run(e, { type: 'toggleBullets' });
    },
    command: () => ({ type: 'indent', direction: 1 }),
  },
  removeBullet: {
    arrange: (e) => {
      select('A')(e);
      run(e, { type: 'toggleBullets' });
    },
    command: () => ({ type: 'removeBullet' }),
  },
  setLineSpacing: onA(() => ({ type: 'setLineSpacing', multiple: 1.5 })),
  setLink: {
    arrange: selectText('A', 0, 5),
    command: () => ({ type: 'setLink', url: 'https://example.com/' }),
  },
  clearFormatting: {
    arrange: (e) => {
      select('A')(e);
      run(e, { type: 'toggleBold' });
    },
    command: () => ({ type: 'clearFormatting' }),
  },
  setVerticalAlign: onA(() => ({ type: 'setVerticalAlign', anchor: 'b' })),
  setAutofit: onA(() => ({ type: 'setAutofit', mode: 'shrink' })),
  setParagraphFormat: onA(() => ({
    type: 'setParagraphFormat',
    spaceBeforePt: 12,
    indentLeftEmu: IN / 4,
  })),
  toggleSuperscript: onA(() => ({ type: 'toggleSuperscript' })),
  toggleSubscript: onA(() => ({ type: 'toggleSubscript' })),
  changeCase: onA(() => ({ type: 'changeCase', case: 'upper' })),
  setCharacterSpacing: onA(() => ({ type: 'setCharacterSpacing', points: 3 })),
  setBulletFormat: onA(() => ({
    type: 'setBulletFormat',
    format: { kind: 'bullet', char: '-' },
  })),
  setTextDirection: onA(() => ({ type: 'setTextDirection', direction: 'vertical' })),
  setColumns: onA(() => ({ type: 'setColumns', count: 2, gapEmu: IN / 4 })),
  setFontFormat: onA(() => ({
    type: 'setFontFormat',
    italic: true,
    caps: 'small',
    characterSpacingPt: 2,
  })),
  pasteFormat: {
    arrange: (e) => {
      select('A')(e);
      run(e, { type: 'toggleBold' });
      run(e, { type: 'copyFormat' });
      select('B')(e);
    },
    command: () => ({ type: 'pasteFormat' }),
  },

  // Insert.
  insertTextBox: { command: () => ({ type: 'insertTextBox' }) },
  insertShape: { command: () => ({ type: 'insertShape', preset: 'ellipse' }) },
  insertLine: { command: () => ({ type: 'insertLine', kind: 'arrow' }) },
  insertImage: {
    command: () => ({
      type: 'insertImage',
      bytes: png(4, 4, 90),
      mime: 'image/png',
      width: 4,
      height: 4,
    }),
  },

  // Arrange.
  order: onA(() => ({ type: 'order', move: 'front' })),
  alignObjects: {
    arrange: select('A', 'C'),
    command: () => ({ type: 'alignObjects', align: 'top' }),
  },
  // A at y 0.5 in, B and C at y 2 in: an even vertical spread moves B.
  distributeObjects: {
    arrange: select('A', 'B', 'C'),
    command: () => ({ type: 'distributeObjects', axis: 'vertical' }),
  },
  flipObjects: onA(() => ({ type: 'flipObjects', axis: 'horizontal' })),
  rotateObjects: onA(() => ({ type: 'rotateObjects', degrees: 90 })),
  duplicate: onA(() => ({ type: 'duplicate' })),
  delete: onA(() => ({ type: 'delete' })),
  setTransform: {
    command: (e) => ({
      type: 'setTransform',
      nodeId: idOf(e, 'B'),
      x: 2 * IN,
      y: 3 * IN,
      cx: 3 * IN,
      cy: IN,
      rot: 0,
    }),
  },
  setTransforms: {
    command: (e) => ({
      type: 'setTransforms',
      items: [
        { nodeId: idOf(e, 'A'), x: 2 * IN, y: IN, cx: 3 * IN, cy: IN, rot: 0 },
        { nodeId: idOf(e, 'B'), x: 2 * IN, y: 2.5 * IN, cx: 3 * IN, cy: IN, rot: 15 },
      ],
    }),
  },
  nudge: onA(() => ({ type: 'nudge', dx: IN / 10, dy: IN / 10 })),
  group: { arrange: select('A', 'B'), command: () => ({ type: 'group' }) },
  ungroup: {
    arrange: (e) => {
      select('A', 'B')(e);
      run(e, { type: 'group' });
    },
    command: () => ({ type: 'ungroup' }),
  },
  paste: {
    arrange: select('A'),
    command: (e) => {
      const payload = e.copy();
      if (!payload) throw new Error('nothing to copy');
      return { type: 'paste', payload };
    },
  },
  // Formatted text pasted with no text selected: a new text box.
  pasteText: {
    command: () => ({
      type: 'pasteText',
      paragraphs: [
        {
          ns: 'http://schemas.openxmlformats.org/drawingml/2006/main',
          local: 'p',
          prefix: 'a',
          children: [
            {
              ns: 'http://schemas.openxmlformats.org/drawingml/2006/main',
              local: 'r',
              prefix: 'a',
              children: [
                {
                  ns: 'http://schemas.openxmlformats.org/drawingml/2006/main',
                  local: 't',
                  prefix: 'a',
                  children: [{ text: 'Pasted' }],
                },
              ],
            },
          ],
        },
      ],
    }),
  },
  setLockAspectRatio: onA(() => ({ type: 'setLockAspectRatio', lock: true })),
  setShapesVisible: {
    command: (e) => ({ type: 'setShapesVisible', nodeIds: [idOf(e, 'B')], visible: false }),
  },
  renameShape: { command: (e) => ({ type: 'renameShape', nodeId: idOf(e, 'C'), name: 'Renamed' }) },

  // Shape format.
  setShapeFill: onA(() => ({ type: 'setShapeFill', color: '#70AD47' })),
  setShapeGradient: onA(() => ({
    type: 'setShapeGradient',
    gradient: {
      kind: 'linear',
      angle: 90,
      rotWithShape: true,
      stops: [
        { pos: 0, color: '#FFFFFF' },
        { pos: 1, color: '#4472C4' },
      ],
    },
  })),
  setShapePattern: onA(() => ({
    type: 'setShapePattern',
    pattern: { preset: 'dkDnDiag', foreground: '#C00000', background: '#FFFFFF' },
  })),
  setShapePicture: onA(() => ({
    type: 'setShapePicture',
    bytes: png(4, 4, 120),
    mime: 'image/png',
    tile: true,
  })),
  setLineColor: onA(() => ({ type: 'setLineColor', color: { theme: 'accent2', shift: -25 } })),
  setLineWidth: onA(() => ({ type: 'setLineWidth', points: 4 })),
  setLineDash: onA(() => ({ type: 'setLineDash', dash: 'dash' })),
  setLineEnds: {
    arrange: (e) => run(e, { type: 'insertLine', kind: 'line' }),
    command: () => ({ type: 'setLineEnds', head: 'oval', tail: 'triangle' }),
  },
  setShapeEffect: onA(() => ({
    type: 'setShapeEffect',
    effect: { kind: 'shadow', preset: 'outerBottomRight' },
  })),
  setShapeStyle: onA(() => ({
    type: 'setShapeStyle',
    style: { row: 'intenseEffect', color: 'accent2' },
  })),

  // A new rounded rectangle (selected after the insert): its corner handle moves.
  setAdjustments: {
    arrange: (e) => void e.exec({ type: 'insertShape', preset: 'roundRect' }),
    command: () => ({ type: 'setAdjustments', values: { adj: 40_000 } }),
  },

  // Slides, notes, and sections.
  addSlide: { command: () => ({ type: 'addSlide' }) },
  duplicateSlide: { command: () => ({ type: 'duplicateSlide', index: 1 }) },
  deleteSlide: { command: () => ({ type: 'deleteSlide', index: 2 }) },
  moveSlide: { command: () => ({ type: 'moveSlide', from: 0, to: 2 }) },
  duplicateSlides: { command: () => ({ type: 'duplicateSlides', indices: [0, 1] }) },
  // Slides copied from the deck itself, as a slide-list copy writes them.
  pasteSlides: {
    command: (e) => {
      const file = e.save();
      if (!file) throw new Error('nothing to save');
      return { type: 'pasteSlides', file, slides: [0, 1] };
    },
  },
  deleteSlides: { command: () => ({ type: 'deleteSlides', indices: [1, 2] }) },
  moveSlides: { command: () => ({ type: 'moveSlides', indices: [0, 1], to: 1 }) },
  setSlideNumbers: { command: () => ({ type: 'setSlideNumbers', show: true }) },
  setFooter: { command: () => ({ type: 'setFooter', show: true, text: 'Quarterly results' }) },
  setDateTime: {
    command: () => ({ type: 'setDateTime', show: true, text: '11/7/2026' }),
  },
  setSlideBackground: {
    command: () => ({
      type: 'setSlideBackground',
      background: { kind: 'solid', color: '#F2F2F2' },
    }),
  },
  // 4:3 from the 16:9 coverage deck: every box on the slides, layouts, and masters scales.
  setSlideSize: {
    command: () => ({ type: 'setSlideSize', width: 9_144_000, height: 6_858_000 }),
  },
  applyLayout: {
    command: (e) => {
      // Slide 3 is on the Blank layout; put it on Title and Content.
      const layout = e.snapshot().layouts?.find((l) => l.type === 'obj');
      if (!layout) throw new Error('no Title and Content layout');
      return { type: 'applyLayout', layoutId: layout.id, indices: [2] };
    },
  },
  // Reset takes the title placeholder `D` back to its layout position and adds the layout's
  // missing content placeholder.
  resetSlides: {
    arrange: (e) => {
      e.goToSlide(1);
      select('D')(e);
      run(e, { type: 'nudge', dx: IN, dy: IN });
    },
    command: () => ({ type: 'resetSlides', indices: [1] }),
  },
  setSlidesHidden: { command: () => ({ type: 'setSlidesHidden', indices: [1], hidden: true }) },
  setNotes: { command: () => ({ type: 'setNotes', index: 1, text: 'New notes' }) },
  formatNotes: {
    command: () => ({ type: 'formatNotes', index: 0, start: 0, end: 7, format: 'bold' }),
  },
  addSection: { command: () => ({ type: 'addSection', name: 'Part two', index: 1 }) },
  renameSection: {
    arrange: (e) => run(e, { type: 'addSection', name: 'Part two', index: 1 }),
    command: (e) => ({ type: 'renameSection', sectionId: sectionId(e, 1), name: 'Renamed' }),
  },
  removeSection: {
    arrange: (e) => run(e, { type: 'addSection', name: 'Part two', index: 1 }),
    command: (e) => ({ type: 'removeSection', sectionId: sectionId(e, 1) }),
  },
  removeAllSections: {
    arrange: (e) => run(e, { type: 'addSection', name: 'Part two', index: 1 }),
    command: () => ({ type: 'removeAllSections' }),
  },
  moveSection: {
    arrange: (e) => run(e, { type: 'addSection', name: 'Part two', index: 1 }),
    command: (e) => ({ type: 'moveSection', sectionId: sectionId(e, 1), direction: -1 }),
  },

  // Find and replace.
  replace: {
    command: (e) => {
      const match = e.find.matches('world')[0];
      if (!match) throw new Error('no match');
      return { type: 'replace', match, replacement: 'there' };
    },
  },
  replaceAll: { command: () => ({ type: 'replaceAll', query: 'shape', replacement: 'box' }) },

  // Pictures.
  cropImage: {
    arrange: select(PICTURE),
    command: () => ({
      type: 'cropImage',
      crop: { kind: 'rect', left: 10, top: 5, right: 10, bottom: 5 },
    }),
  },
  cropToShape: {
    arrange: select(PICTURE),
    command: () => ({ type: 'cropToShape', preset: 'ellipse' }),
  },
  replaceImage: {
    arrange: select(PICTURE),
    command: () => ({ type: 'replaceImage', bytes: png(6, 6, 40), mime: 'image/png' }),
  },
  resetImage: {
    arrange: (e) => {
      select(PICTURE)(e);
      run(e, { type: 'setImageStyle', style: 'simpleFrameBlack' });
    },
    command: () => ({ type: 'resetImage' }),
  },
  setImageAdjustments: {
    arrange: select(PICTURE),
    command: () => ({
      type: 'setImageAdjustments',
      adjustments: { brightness: 0.2, contrast: -0.1, recolor: 'grayscale' },
    }),
  },
  setImageStyle: {
    arrange: select(PICTURE),
    command: () => ({ type: 'setImageStyle', style: 'softEdgeOval' }),
  },
  setImageProperties: {
    arrange: select(PICTURE),
    command: () => ({
      type: 'setImageProperties',
      widthEmu: 3 * IN,
      rotationDegrees: 10,
    }),
  },
  setAltText: {
    arrange: select(PICTURE),
    command: () => ({ type: 'setAltText', description: 'A gray square', title: 'Square' }),
  },

  // Tables.
  insertTable: { command: () => ({ type: 'insertTable', rows: 2, cols: 2 }) },
  insertRow: { arrange: atCell(0, 1), command: () => ({ type: 'insertRow', where: 'below' }) },
  insertColumn: { arrange: atCell(1, 1), command: () => ({ type: 'insertColumn', where: 'left' }) },
  deleteRow: { arrange: atCell(1, 0), command: () => ({ type: 'deleteRow' }) },
  deleteColumn: { arrange: atCell(0, 2), command: () => ({ type: 'deleteColumn' }) },
  deleteTable: { arrange: atCell(0, 0), command: () => ({ type: 'deleteTable' }) },
  mergeCells: {
    arrange: (e) => e.setSelection(cellRange(e, [0, 0], [1, 1])),
    command: () => ({ type: 'mergeCells' }),
  },
  splitCell: {
    arrange: (e) => {
      e.setSelection(cellRange(e, [0, 0], [0, 1]));
      run(e, { type: 'mergeCells' });
      e.setSelection(cellCaret(e, 0, 0));
    },
    command: () => ({ type: 'splitCell', rows: 1, cols: 2 }),
  },
  setCellFill: {
    arrange: atCell(1, 2),
    command: () => ({ type: 'setCellFill', color: '#FFC000' }),
  },
  setTableBorders: {
    arrange: (e) => e.setSelection(cellRange(e, [0, 0], [1, 2])),
    command: () => ({
      type: 'setTableBorders',
      scope: 'outside',
      spec: { color: '#000000', width: 2, dash: 'sysDash' },
    }),
  },
  setTableStyle: {
    arrange: atCell(0, 0),
    command: (e) => {
      const style = e.tableStyles()[0];
      if (!style) throw new Error('no table style');
      return { type: 'setTableStyle', styleId: style.id };
    },
  },
  toggleTableStyleOption: {
    arrange: atCell(0, 0),
    command: () => ({ type: 'toggleTableStyleOption', option: 'bandCol' }),
  },
  setCellMargins: {
    arrange: atCell(0, 0),
    command: () => ({ type: 'setCellMargins', margins: { left: IN / 8, top: IN / 16 } }),
  },
  pasteCells: {
    arrange: atCell(0, 1),
    command: () => ({
      type: 'pasteCells',
      texts: [
        ['x', 'y'],
        ['two\nlines', 'w'],
      ],
    }),
  },
  setRowHeight: { arrange: atCell(1, 0), command: () => ({ type: 'setRowHeight', height: IN }) },
  setColumnWidth: {
    arrange: atCell(0, 1),
    command: () => ({ type: 'setColumnWidth', width: 1.5 * IN }),
  },
  distributeRows: {
    arrange: (e) => {
      e.setSelection(cellCaret(e, 0, 0));
      run(e, { type: 'setRowHeight', height: IN });
      e.setSelection(cellRange(e, [0, 0], [1, 0]));
    },
    command: () => ({ type: 'distributeRows' }),
  },
  distributeColumns: {
    arrange: (e) => {
      e.setSelection(cellCaret(e, 0, 0));
      run(e, { type: 'setColumnWidth', width: 2 * IN });
      e.setSelection(cellRange(e, [0, 0], [0, 2]));
    },
    command: () => ({ type: 'distributeColumns' }),
  },
  setTableGrid: {
    arrange: atCell(0, 0),
    command: () => ({
      type: 'setTableGrid',
      columns: [IN, 1.5 * IN, IN / 2],
      rows: [370_840, 2 * 370_840],
    }),
  },
};

/** A caret, or a selection when `end` is given, in the first paragraph of shape `name`. */
const caret = (name: string, start: number, end = start) => selectText(name, start, end);

/**
 * The text input fixtures, by kind: typing, deletion, Enter, and Shift+Enter, through
 * `PptxEditorInstance.input`. Shape `A` holds "Hello world".
 */
export const TEXT_INPUT_FIXTURES: Record<TextInputKind, TextInputFixture> = {
  text: { arrange: caret('A', 0, 5), data: 'Goodbye' },
  deleteBackward: { arrange: caret('A', 5) },
  deleteForward: { arrange: caret('A', 0) },
  deleteWordBackward: { arrange: caret('A', 11) },
  deleteWordForward: { arrange: caret('A', 0) },
  deleteLineBackward: { arrange: caret('A', 5) },
  deleteLineForward: { arrange: caret('A', 6) },
  // Enter in a table cell: the cell gets a second paragraph.
  paragraph: { arrange: (e) => e.setSelection(cellCaret(e, 0, 0, 1)) },
  lineBreak: { arrange: caret('A', 5) },
};

/**
 * Probes for the uncovered commands, whose reason is that they write nothing to the deck: the
 * gate runs each and checks that neither replica changes. Every entry of
 * `COLLABORATION_UNCOVERED` needs one.
 */
export const NO_WRITE_PROBES: Partial<Record<EditorCommandType, ConvergenceFixture>> = {
  toggleSmartGuides: { command: () => ({ type: 'toggleSmartGuides' }) },
  setTableBorderPen: {
    command: () => ({
      type: 'setTableBorderPen',
      pen: { target: 'outside', drawn: true, spec: { color: '#FF0000', width: 2 } },
    }),
  },
  setPresentationView: { command: () => ({ type: 'setPresentationView', view: 'slideSorter' }) },
  setSectionsCollapsed: {
    arrange: (e) => run(e, { type: 'addSection', name: 'Part two', index: 1 }),
    command: () => ({ type: 'setSectionsCollapsed', collapsed: true }),
  },
  toggleFormatPainter: onA(() => ({ type: 'toggleFormatPainter' })),
  copyFormat: onA(() => ({ type: 'copyFormat' })),
  toggleCrop: { arrange: select(PICTURE), command: () => ({ type: 'toggleCrop' }) },
  selectTableRegion: {
    arrange: atCell(0, 0),
    command: () => ({ type: 'selectTableRegion', region: 'row' }),
  },
};
