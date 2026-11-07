// Every `EditorCommand` type, as a value, for gates that must see each command: the
// collaboration coverage gate (`packages/pro/src/__tests__/collaboration-coverage.test.ts`)
// requires a two-peer convergence fixture or a written reason for each one.
//
// `EditorCommand` is a type, so the list cannot be read from it at run time. The table below is
// checked against the type in both directions (`satisfies`): a command added to the union
// without a row here, or a row for a command that no longer exists, fails `bun run typecheck`.
// So a new command reaches the gates without a hand-kept list that can drift.
//
// When `bun run typecheck` reports "Property '<name>' is missing" on `COMMAND_TABLE`, add
// `<name>: true` to the group the command belongs to. Then follow CONTRIBUTING.md, "Add a
// collaboration fixture".

import type { EditorCommandType } from '../contracts/index.ts';
import type { TextInputKind } from './text-input.ts';

const COMMAND_TABLE = {
  // History
  undo: true,
  redo: true,
  // Character and paragraph formatting
  toggleBold: true,
  toggleItalic: true,
  toggleUnderline: true,
  toggleStrike: true,
  setFontSize: true,
  stepFontSize: true,
  setFontFamily: true,
  setTextColor: true,
  setHighlight: true,
  setAlign: true,
  toggleBullets: true,
  toggleNumbering: true,
  indent: true,
  removeBullet: true,
  setLineSpacing: true,
  setLink: true,
  clearFormatting: true,
  setVerticalAlign: true,
  setAutofit: true,
  setTextWarp: true,
  setParagraphFormat: true,
  toggleSuperscript: true,
  toggleSubscript: true,
  changeCase: true,
  setCharacterSpacing: true,
  setBulletFormat: true,
  setTextDirection: true,
  setColumns: true,
  setFontFormat: true,
  toggleFormatPainter: true,
  copyFormat: true,
  pasteFormat: true,
  // Insert
  insertTextBox: true,
  insertShape: true,
  insertLine: true,
  insertImage: true,
  // Arrange
  order: true,
  alignObjects: true,
  distributeObjects: true,
  flipObjects: true,
  rotateObjects: true,
  duplicate: true,
  delete: true,
  setTransform: true,
  setTransforms: true,
  nudge: true,
  group: true,
  ungroup: true,
  paste: true,
  pasteText: true,
  pasteSlides: true,
  setLockAspectRatio: true,
  setShapesVisible: true,
  renameShape: true,
  // Shape format
  setShapeFill: true,
  setShapeGradient: true,
  setShapePattern: true,
  setShapePicture: true,
  setLineColor: true,
  setLineWidth: true,
  setLineDash: true,
  setLineEnds: true,
  setShapeEffect: true,
  setShapeStyle: true,
  setAdjustments: true,
  // Slides, notes, and sections
  addSlide: true,
  duplicateSlide: true,
  deleteSlide: true,
  moveSlide: true,
  duplicateSlides: true,
  deleteSlides: true,
  moveSlides: true,
  setSlideNumbers: true,
  setTransition: true,
  addAnimation: true,
  removeAnimation: true,
  moveAnimation: true,
  setAnimationTiming: true,
  setFooter: true,
  setDateTime: true,
  setSlideBackground: true,
  setSlideSize: true,
  applyLayout: true,
  resetSlides: true,
  setSlidesHidden: true,
  setNotes: true,
  formatNotes: true,
  addSection: true,
  renameSection: true,
  removeSection: true,
  removeAllSections: true,
  moveSection: true,
  // Find and replace
  replace: true,
  replaceAll: true,
  // View settings
  toggleSmartGuides: true,
  showAnimationOrder: true,
  setTableBorderPen: true,
  setPresentationView: true,
  setSectionsCollapsed: true,
  // Pictures
  cropImage: true,
  cropToShape: true,
  toggleCrop: true,
  replaceImage: true,
  resetImage: true,
  setImageAdjustments: true,
  setImageStyle: true,
  setImageProperties: true,
  setAltText: true,
  // Charts
  setChartData: true,
  // Diagrams
  addDiagramShape: true,
  deleteDiagramShape: true,
  // Tables
  insertTable: true,
  insertRow: true,
  insertColumn: true,
  deleteRow: true,
  deleteColumn: true,
  deleteTable: true,
  selectTableRegion: true,
  mergeCells: true,
  splitCell: true,
  setCellFill: true,
  setTableBorders: true,
  setTableStyle: true,
  toggleTableStyleOption: true,
  setCellMargins: true,
  setRowHeight: true,
  setColumnWidth: true,
  distributeRows: true,
  distributeColumns: true,
  pasteCells: true,
  setTableGrid: true,
} as const satisfies Record<EditorCommandType, true>;

/** Every `EditorCommand['type']`, checked against the union by the compiler. @internal */
export const EDITOR_COMMAND_TYPES: readonly EditorCommandType[] = Object.freeze(
  Object.keys(COMMAND_TABLE) as EditorCommandType[]
);

/**
 * Every `TextInputKind`, the other write path of the editor (`PptxEditorInstance.input`: typing,
 * deletion, Enter, Shift+Enter). Checked against the union the same way. @internal
 */
export const TEXT_INPUT_KINDS: readonly TextInputKind[] = Object.freeze(
  Object.keys({
    text: true,
    deleteBackward: true,
    deleteForward: true,
    deleteWordBackward: true,
    deleteWordForward: true,
    deleteLineBackward: true,
    deleteLineForward: true,
    paragraph: true,
    lineBreak: true,
  } as const satisfies Record<TextInputKind, true>) as TextInputKind[]
);

/** The commands that insert an object, for `snapshot.insertions`. @internal */
export const INSERT_COMMAND_TYPES: ReadonlySet<EditorCommandType> = new Set([
  'insertShape',
  'insertLine',
  'insertTextBox',
  'insertImage',
  'insertTable',
]);
