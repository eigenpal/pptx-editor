/** @jsxImportSource react */
// The named toolbar parts (`Toolbar.Bold`, `Toolbar.FontFamily`, …), after DOCX Editor's
// `definePart` (docx-editor/packages/react/src/editor/toolbar/parts.tsx). Each part is a
// `Toolbar.Slot` with its slot pinned, and carries the slot as a static (`pptxSlot`), so the
// toolbar recognizes it among its children and puts it in the slot's place. The table of
// names lives in core (`TOOLBAR_NAMED_PARTS`), so the Vue twin uses the same names.

import type { ReactNode } from 'react';
import { TOOLBAR_NAMED_PARTS, type ChromeSlotId } from '@pptx-editor.dev/core/editor';
import { ToolbarSlot } from './Toolbar.tsx';
import {
  namedPartSlotProps,
  type ToolbarColorSplitComponent,
  type ToolbarPartComponent,
  type ToolbarSlotPartComponent,
} from './toolbar-parts.tsx';
import { FONT_FAMILY_PARTS, type FontFamilyNamespace } from './FontFamilyControl.tsx';

function definePart(name: string, slot: ChromeSlotId): ToolbarPartComponent {
  const Part = (props: Record<string, unknown>): ReactNode => (
    <ToolbarSlot {...namedPartSlotProps(slot, props, `Toolbar.${name}`)} />
  );
  return Object.assign(Part, { pptxSlot: slot }) as unknown as ToolbarPartComponent;
}

/**
 * Every named toolbar part, attached to `PptxEditor.Toolbar` as statics: button parts take
 * `ToolbarPartProps`, picker parts `ToolbarSlotPartProps`, as in DOCX Editor. @public
 */
export interface ToolbarNamedParts {
  /** `history.undo`. */
  readonly Undo: ToolbarPartComponent;
  /** `history.redo`. */
  readonly Redo: ToolbarPartComponent;
  /** `font.family`, with its `Trigger`, `Content`, and `Item` parts. */
  readonly FontFamily: FontFamilyNamespace;
  /** `font.size`. */
  readonly FontSize: ToolbarSlotPartComponent;
  /** `text.bold`. */
  readonly Bold: ToolbarPartComponent;
  /** `text.italic`. */
  readonly Italic: ToolbarPartComponent;
  /** `text.underline`. */
  readonly Underline: ToolbarPartComponent;
  /** `text.strike`. */
  readonly Strike: ToolbarPartComponent;
  /** `text.color`. */
  readonly FontColor: ToolbarColorSplitComponent;
  /** `text.highlight`. */
  readonly Highlight: ToolbarColorSplitComponent;
  /** `text.link`. */
  readonly Link: ToolbarPartComponent;
  /** `format.clear`. */
  readonly ClearFormatting: ToolbarPartComponent;
  /** `text.characterSpacing`. */
  readonly CharacterSpacing: ToolbarSlotPartComponent;
  /** `text.changeCase`. */
  readonly ChangeCase: ToolbarSlotPartComponent;
  /** `text.dialog`: opens the Font dialog. */
  readonly Font: ToolbarPartComponent;
  /** `script.super`. */
  readonly Superscript: ToolbarPartComponent;
  /** `script.sub`. */
  readonly Subscript: ToolbarPartComponent;
  /** `paragraph.align`. */
  readonly Alignment: ToolbarSlotPartComponent;
  /** `paragraph.valign`. */
  readonly VerticalAlignment: ToolbarSlotPartComponent;
  /** `list.bullet`. */
  readonly BulletList: ToolbarPartComponent;
  /** `list.numbered`. */
  readonly NumberedList: ToolbarPartComponent;
  /** `list.bulletStyle`: the Bullets gallery. */
  readonly BulletStyle: ToolbarSlotPartComponent;
  /** `list.numberStyle`: the Numbering gallery. */
  readonly NumberStyle: ToolbarSlotPartComponent;
  /** `list.outdent`. */
  readonly Outdent: ToolbarPartComponent;
  /** `list.indent`. */
  readonly Indent: ToolbarPartComponent;
  /** `list.lineSpacing`. */
  readonly LineSpacing: ToolbarSlotPartComponent;
  /** `paragraph.dialog`: opens the Paragraph dialog. */
  readonly Paragraph: ToolbarPartComponent;
  /** `paragraph.textDirection`. */
  readonly TextDirection: ToolbarSlotPartComponent;
  /** `paragraph.autofit`. */
  readonly Autofit: ToolbarSlotPartComponent;
  /** `paragraph.columns`. */
  readonly Columns: ToolbarSlotPartComponent;
  /** `insert.textBox`. */
  readonly TextBox: ToolbarPartComponent;
  /** `insert.line`. */
  readonly Line: ToolbarSlotPartComponent;
  /** `insert.shape`. */
  readonly Shapes: ToolbarSlotPartComponent;
  /** `insert.image`. */
  readonly ImageInsert: ToolbarPartComponent;
  /** `shape.fill`. */
  readonly ShapeFill: ToolbarColorSplitComponent;
  /** `shape.lineColor`. */
  readonly LineColor: ToolbarColorSplitComponent;
  /** `shape.lineWidth`. */
  readonly LineWidth: ToolbarSlotPartComponent;
  /** `shape.lineStart`. */
  readonly LineStart: ToolbarSlotPartComponent;
  /** `shape.lineEnd`. */
  readonly LineEnd: ToolbarSlotPartComponent;
  /** `shape.lineDash`. */
  readonly LineDash: ToolbarSlotPartComponent;
  /** `shape.styles`. */
  readonly ShapeStyles: ToolbarSlotPartComponent;
  /** `shape.effects`. */
  readonly ShapeEffects: ToolbarSlotPartComponent;
  /** `shape.textEffects`. */
  readonly TextEffects: ToolbarSlotPartComponent;
  /** `shape.properties`. */
  readonly FormatShape: ToolbarPartComponent;
  /** `format.painter`. */
  readonly FormatPainter: ToolbarPartComponent;
  /** `arrange.order`. */
  readonly Order: ToolbarSlotPartComponent;
  /** `arrange.align`. */
  readonly AlignObjects: ToolbarSlotPartComponent;
  /** `arrange.group`. */
  readonly GroupObjects: ToolbarPartComponent;
  /** `arrange.ungroup`. */
  readonly UngroupObjects: ToolbarPartComponent;
  /** `arrange.rotate`. */
  readonly Rotate: ToolbarSlotPartComponent;
  /** `arrange.selectionPane`. */
  readonly SelectionPane: ToolbarPartComponent;
  /** `more.duplicate`. */
  readonly Duplicate: ToolbarPartComponent;
  /** `more.delete`. */
  readonly Delete: ToolbarPartComponent;
  /** `find.find`. */
  readonly Find: ToolbarPartComponent;
  /** `find.replace`. */
  readonly Replace: ToolbarPartComponent;
  /** `view.smartGuides`. */
  readonly SmartGuides: ToolbarPartComponent;
  /** `file.pageSetup`. */
  readonly PageSetup: ToolbarPartComponent;
  /** `edit.cut`. */
  readonly Cut: ToolbarPartComponent;
  /** `edit.copy`. */
  readonly Copy: ToolbarPartComponent;
  /** `edit.paste`. */
  readonly Paste: ToolbarPartComponent;
  /** `edit.selectAll`. */
  readonly SelectAll: ToolbarPartComponent;
  /** `slide.background`. */
  readonly FormatBackground: ToolbarPartComponent;
  /** `slide.add`: New slide, with its layout gallery. */
  readonly NewSlide: ToolbarPartComponent;
  /** `slide.layout`: the layout gallery of the current slide. */
  readonly SlideLayout: ToolbarPartComponent;
  /** `slide.reset`. */
  readonly ResetSlide: ToolbarPartComponent;
  /** `slide.hide`. */
  readonly HideSlide: ToolbarPartComponent;
  /** `view.presentationView`: the Normal or Slide sorter view. */
  readonly PresentationView: ToolbarPartComponent;
  /** `section.add`. */
  readonly AddSection: ToolbarPartComponent;
  /** `section.rename`. */
  readonly RenameSection: ToolbarPartComponent;
  /** `section.remove`. */
  readonly RemoveSection: ToolbarPartComponent;
  /** `section.removeAll`. */
  readonly RemoveAllSections: ToolbarPartComponent;
  /** `section.collapseAll`. */
  readonly CollapseAllSections: ToolbarPartComponent;
  /** `section.expandAll`. */
  readonly ExpandAllSections: ToolbarPartComponent;
  /** `view.playFromStart`. */
  readonly PlayFromStart: ToolbarPartComponent;
  /** `view.playFromCurrent`. */
  readonly PlayFromCurrent: ToolbarPartComponent;
  /** `image.corrections`. */
  readonly ImageCorrections: ToolbarPartComponent;
  /** `image.color`. */
  readonly ImageColor: ToolbarPartComponent;
  /** `image.transparency`. */
  readonly ImageTransparency: ToolbarPartComponent;
  /** `image.change`. */
  readonly ImageChange: ToolbarPartComponent;
  /** `image.reset`. */
  readonly ImageReset: ToolbarPartComponent;
  /** `image.styles`. */
  readonly ImageStyles: ToolbarPartComponent;
  /** `image.altText`. */
  readonly ImageAltText: ToolbarPartComponent;
  /** `image.crop`. */
  readonly ImageCrop: ToolbarPartComponent;
  /** `image.cropToShape`. */
  readonly ImageCropToShape: ToolbarPartComponent;
  /** `image.cropFit`. */
  readonly ImageCropFit: ToolbarPartComponent;
  /** `image.properties`. */
  readonly ImageProperties: ToolbarPartComponent;
  /** `chart.editData`. */
  readonly ChartEditData: ToolbarPartComponent;
  /** `table.insert`. */
  readonly TableInsert: ToolbarPartComponent;
  /** `table.style`. */
  readonly TableStyle: ToolbarSlotPartComponent;
  /** `table.headerRow`. */
  readonly TableHeaderRow: ToolbarPartComponent;
  /** `table.totalRow`. */
  readonly TableTotalRow: ToolbarPartComponent;
  /** `table.bandedRows`. */
  readonly TableBandedRows: ToolbarPartComponent;
  /** `table.firstColumn`. */
  readonly TableFirstColumn: ToolbarPartComponent;
  /** `table.lastColumn`. */
  readonly TableLastColumn: ToolbarPartComponent;
  /** `table.bandedColumns`. */
  readonly TableBandedColumns: ToolbarPartComponent;
  /** `table.cellFill`. */
  readonly TableCellFill: ToolbarColorSplitComponent;
  /** `table.borderTarget`. */
  readonly TableBorderTarget: ToolbarSlotPartComponent;
  /** `table.borderColor`. */
  readonly TableBorderColor: ToolbarColorSplitComponent;
  /** `table.borderStyle`. */
  readonly TableBorderStyle: ToolbarSlotPartComponent;
  /** `table.borderWidth`. */
  readonly TableBorderWidth: ToolbarSlotPartComponent;
  /** `table.select`. */
  readonly TableSelect: ToolbarSlotPartComponent;
  /** `table.delete`. */
  readonly TableDelete: ToolbarSlotPartComponent;
  /** `table.insertRowAbove`. */
  readonly TableInsertRowAbove: ToolbarPartComponent;
  /** `table.insertRowBelow`. */
  readonly TableInsertRowBelow: ToolbarPartComponent;
  /** `table.insertColumnLeft`. */
  readonly TableInsertColumnLeft: ToolbarPartComponent;
  /** `table.insertColumnRight`. */
  readonly TableInsertColumnRight: ToolbarPartComponent;
  /** `table.mergeCells`. */
  readonly TableMergeCells: ToolbarPartComponent;
  /** `table.splitCell`. */
  readonly TableSplitCell: ToolbarSlotPartComponent;
  /** `table.distributeRows`. */
  readonly TableDistributeRows: ToolbarPartComponent;
  /** `table.distributeColumns`. */
  readonly TableDistributeColumns: ToolbarPartComponent;
  /** `table.cellMargins`. */
  readonly TableCellMargins: ToolbarSlotPartComponent;
}

/** The named parts, built from core's table. `FontFamily` carries its compound parts. */
export const TOOLBAR_PARTS = Object.fromEntries(
  Object.entries(TOOLBAR_NAMED_PARTS).map(([name, slot]) => {
    const part = definePart(name, slot);
    return [name, name === 'FontFamily' ? Object.assign(part, FONT_FAMILY_PARTS) : part];
  })
) as unknown as ToolbarNamedParts;

// Every name in core's table has a typed part here: a name added there and not here fails the
// build, instead of reaching hosts as an untyped static.
type MissingPart = Exclude<keyof typeof TOOLBAR_NAMED_PARTS, keyof ToolbarNamedParts>;
const everyPartTyped: [MissingPart] extends [never] ? true : MissingPart = true;
void everyPartTyped;
