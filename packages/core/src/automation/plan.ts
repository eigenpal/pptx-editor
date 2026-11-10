// One operation, planned: a query answers from the working package, and a command adds its
// store ops to the batch. The host runs this for each operation in order.

import { readShapeFormatProperties, patchShapeFormatProperties } from './commands/format-detail.ts';
import { edit } from './targets.ts';
import type { AutomationOperation } from './operations.ts';
import type { AutomationRef, AutomationValue } from './protocol.ts';
import * as adjustments from './plan-adjustments.ts';
import * as backgrounds from './plan-backgrounds.ts';
import * as bindings from './plan-bindings.ts';
import type { AutomationBindings } from './plan-bindings.ts';
import * as images from './plan-images.ts';
import * as importing from './plan-import.ts';
import * as custom from './plan-custom-properties.ts';
import * as customXml from './plan-custom-xml.ts';
import * as reads from './plan-reads.ts';
import * as props from './plan-properties.ts';
import * as shapes from './plan-shapes.ts';
import * as slides from './plan-slides.ts';
import * as tableFormat from './plan-table-format.ts';
import * as tables from './plan-tables.ts';
import * as text from './plan-text.ts';
import * as theme from './plan-theme.ts';
import { AUTOMATION_COMMANDS } from './operations.ts';
import {
  AutomationFailure,
  assertUniqueShape,
  badArgument,
  shapeAt,
  slideOf,
  targetOf,
} from './targets.ts';
import type { TextPlanContext } from './plan-text.ts';

/** A reader's selection, as the browser host's port reports and sets it. */
export interface AutomationSelectionPort {
  /** The selected slides (`p:sldId/@id`), the selected shapes, and the selected text. */
  read(): {
    readonly slides: readonly string[];
    readonly shapes: { readonly sldId: string; readonly shapeIds: readonly string[] } | null;
    readonly text: {
      readonly sldId: string;
      readonly shapeId: string;
      readonly start: number;
      readonly length: number;
    } | null;
  };
  selectSlides(sldIds: readonly string[]): void;
  selectShapes(sldId: string, shapeIds: readonly string[]): void;
  selectText(sldId: string, shapeId: string, start: number, length: number): void;
}

/** What a batch plans against. */
export interface BatchPlanContext extends TextPlanContext {
  /** The id `readPresentation` answers. */
  readonly presentationId: string;
  readonly selection?: AutomationSelectionPort;
  /** The working copy of the host's bindings, which the host keeps when the batch commits. */
  readonly bindings: AutomationBindings;
  /** Work to run after the commit: selection changes name committed objects. */
  readonly afterCommit: (() => void)[];
  /**
   * The selection that the selection writes of this batch leave, once they run after the
   * commit. A selection read later in the same batch answers it, so a read sees the writes
   * queued before it.
   */
  readonly selectionAfter?: { value?: ReturnType<AutomationSelectionPort['read']> };
}

function selectionPort(ctx: BatchPlanContext): AutomationSelectionPort {
  if (!ctx.selection)
    throw new AutomationFailure(
      'unsupported-capability',
      'this host has no selection',
      'selection'
    );
  return ctx.selection;
}

function readSelection(ctx: BatchPlanContext): AutomationValue {
  const port = selectionPort(ctx);
  const now = ctx.selectionAfter?.value ?? port.read();
  const known = (sldId: string): boolean => {
    try {
      slideOf(ctx.batch.pkg, sldId);
      return true;
    } catch {
      return false;
    }
  };
  const shapeHandle = (sldId: string, shapeId: string) =>
    ctx.handles.mint({ kind: 'shape', owner: { kind: 'slide', sldId }, shapeId });
  return {
    kind: 'selection',
    selection: {
      slides: now.slides.filter(known).map((sldId) => ctx.handles.mint({ kind: 'slide', sldId })),
      shapes:
        now.shapes && known(now.shapes.sldId)
          ? now.shapes.shapeIds.map((id) => shapeHandle(now.shapes!.sldId, id))
          : [],
      text:
        now.text && known(now.text.sldId)
          ? {
              shape: shapeHandle(now.text.sldId, now.text.shapeId),
              start: now.text.start,
              length: now.text.length,
            }
          : null,
    },
  };
}

function shapesOnOneSlide(ctx: BatchPlanContext, refs: readonly AutomationRef[]) {
  if (!Array.isArray(refs)) badArgument('shapes');
  const at = refs.map((ref) => shapeAt(ctx, ref));
  const owners = new Set(at.map((s) => (s.owner.kind === 'slide' ? s.owner.sldId : '')));
  if (owners.has('') || owners.size > 1)
    throw new AutomationFailure('invalid-argument', 'select shapes of one slide', 'shapes');
  return at;
}

/** The shape references a command writes through. */
function writtenShapes(op: AutomationOperation): readonly unknown[] {
  if (!AUTOMATION_COMMANDS.has(op.op)) return [];
  const fields = op as {
    shape?: unknown;
    shapes?: unknown;
    group?: unknown;
    owner?: unknown;
    container?: unknown;
  };
  // `addShape` names the new shape's kind in `shape`, not a reference.
  const refs: unknown[] = [
    op.op === 'addShape' ? undefined : fields.shape,
    fields.group,
    fields.owner,
    fields.container,
  ];
  const listed: unknown[] = Array.isArray(fields.shapes) ? fields.shapes : [];
  return refs.concat(listed).filter((ref) => typeof ref === 'object' && ref !== null);
}

/** Plan one operation. Throws `AutomationFailure` or `NullReference`. */
export function planOperation(ctx: BatchPlanContext, op: AutomationOperation): AutomationValue {
  for (const ref of writtenShapes(op)) assertUniqueShape(ctx, ref);
  switch (op.op) {
    case 'getPresentation':
      return { kind: 'handle', handle: ctx.handles.mint({ kind: 'presentation' }) };
    case 'readPresentation':
      return reads.readPresentation(ctx, ctx.presentationId);
    case 'setPageSetup':
      return props.setPageSetup(ctx, op.slideWidth, op.slideHeight);
    case 'getProperties':
      return props.getProperties(ctx);
    case 'setProperties':
      return props.setProperties(ctx, op.properties);
    case 'readCustomProperties':
      targetOf(ctx, op.owner, 'presentation');
      return custom.readCustomProperties(ctx);
    case 'setCustomProperty':
      targetOf(ctx, op.owner, 'presentation');
      return custom.setCustomProperty(ctx, op.key, op.value);
    case 'deleteCustomProperty':
      targetOf(ctx, op.owner, 'presentation');
      return custom.deleteCustomProperty(ctx, op.key);
    case 'deleteCustomProperties':
      targetOf(ctx, op.owner, 'presentation');
      return custom.deleteCustomProperties(ctx);
    case 'listCustomXmlParts':
      return customXml.listCustomXmlParts(ctx, op.owner, op.namespaceUri);
    case 'readCustomXml':
      return customXml.readCustomXml(ctx, op.owner, op.id);
    case 'addCustomXmlPart':
      return customXml.addCustomXmlPart(ctx, op.owner, op.xml, op.id);
    case 'setCustomXml':
      return customXml.setCustomXml(ctx, op.owner, op.id, op.xml);
    case 'deleteCustomXmlPart':
      return customXml.deleteCustomXmlPart(ctx, op.owner, op.id);
    case 'listSlides':
      return reads.listSlides(ctx);
    case 'getSlide':
      return reads.getSlide(ctx, op.key, op.orNull);
    case 'readSlide':
      return reads.readSlide(ctx, op.slide);
    case 'addSlide':
      return slides.addSlide(ctx, op.index, op.layout, op.master, op.layoutId, op.masterId);
    case 'deleteSlide':
      return slides.deleteSlide(ctx, op.slide);
    case 'moveSlide':
      return slides.moveSlide(ctx, op.slide, op.index);
    case 'applyLayout':
      return slides.applyLayout(ctx, op.slide, op.layout);
    case 'exportSlides':
      return slides.exportSlides(ctx, op.slides);
    case 'prepareImage':
      return images.prepareImage(ctx, op.target, op.width, op.height);
    case 'insertSlides':
      return importing.insertSlides(ctx, op.base64, op.options);
    case 'listHyperlinks':
      return slides.listHyperlinks(ctx, op.slide);
    case 'listTextHyperlinks':
      return slides.listTextHyperlinks(ctx, op.shape, op.span);
    case 'listBindings':
      return bindings.listBindings(ctx, op.owner);
    case 'addBinding':
      return bindings.addBinding(ctx, op.owner, op.shape, op.id);
    case 'addBindingFromSelection':
      return bindings.addBindingFromSelection(ctx, op.owner, op.id);
    case 'deleteBinding':
      return bindings.deleteBinding(ctx, op.owner, op.id);
    case 'getBindingShape':
      return bindings.getBindingShape(ctx, op.owner, op.id);
    case 'listMasters':
      return reads.listMasters(ctx);
    case 'getMaster':
      return reads.getMaster(ctx, op.key, op.orNull);
    case 'readMaster':
      return reads.readMaster(ctx, op.master);
    case 'listLayouts':
      return reads.listLayouts(ctx, op.master);
    case 'getLayout':
      return reads.getLayout(ctx, op.master, op.key, op.orNull);
    case 'readLayout':
      return reads.readLayout(ctx, op.layout);
    case 'listShapes':
      return reads.listShapes(ctx, op.container, op.orNull);
    case 'getShape':
      return reads.getShape(ctx, op.container, op.key, op.orNull);
    case 'readShape':
      return op.group && !reads.isGroup(ctx, op.shape, op.orNull)
        ? { kind: 'null' }
        : reads.readShape(ctx, op.shape);
    case 'addShape':
      return shapes.addShape(ctx, op.container, op.shape, op.box);
    case 'deleteShape':
      return shapes.deleteShape(ctx, op.shape);
    case 'setShape':
      return shapes.setShape(ctx, op.shape, op.props);
    case 'setZOrder':
      return shapes.setZOrder(ctx, op.shape, op.position);
    case 'group':
      return shapes.group(ctx, op.shapes);
    case 'ungroup':
      return shapes.ungroup(ctx, op.group);
    case 'parentOf':
      return shapes.parentOf(ctx, op.shape, op.kind, op.orNull);
    case 'getParentGroup':
      return shapes.parentGroupOf(ctx, op.shape, op.orNull);
    case 'setShapeHyperlink':
      return shapes.setShapeHyperlink(ctx, op.shape, op.address, op.screenTip);
    case 'getSlideLayout': {
      const read = reads.readSlide(ctx, op.slide);
      const layout = read.kind === 'slide' ? read.slide.layout : null;
      if (!layout) throw new AutomationFailure('unsupported-content', 'the slide has no layout');
      return { kind: 'handle', handle: layout };
    }
    case 'getSlideMaster': {
      const target = targetOf(ctx, op.owner, 'slide', 'slideLayout');
      const read =
        target.kind === 'slide' ? reads.readSlide(ctx, op.owner) : reads.readLayout(ctx, op.owner);
      const master =
        read.kind === 'slide'
          ? read.slide.master
          : read.kind === 'layout'
            ? read.layout.master
            : null;
      if (!master) throw new AutomationFailure('unsupported-content', 'no slide master');
      return { kind: 'handle', handle: master };
    }
    case 'readShapeFormat': {
      const at = shapeAt(ctx, op.shape);
      return {
        kind: 'shapeFormat',
        format: edit(() => readShapeFormatProperties(ctx.batch, at.part, at.el.id)),
      };
    }
    case 'patchShapeFormat': {
      const at = shapeAt(ctx, op.shape);
      edit(() => patchShapeFormatProperties(ctx.batch, at.part, at.el.id, op.patch));
      return { kind: 'applied' };
    }
    case 'readFill':
      return op.cell === undefined
        ? reads.readFill(ctx, op.shape)
        : tables.readTableCellFill(ctx, op.shape, op.cell);
    case 'setFill':
      return op.cell === undefined
        ? shapes.setFill(ctx, op.shape, op.fill)
        : tables.setTableCellFill(ctx, op.shape, op.cell, op.fill);
    case 'readLine':
      return reads.readLine(ctx, op.shape);
    case 'setLine':
      return shapes.setLine(ctx, op.shape, op.line);
    case 'readBackground':
      return backgrounds.readBackground(ctx, op.owner);
    case 'setBackground':
      return backgrounds.setBackground(ctx, op.owner, op.background);
    case 'readThemeColor':
      return theme.readThemeColor(ctx, op.owner, op.slot);
    case 'setThemeColor':
      return theme.setThemeColor(ctx, op.owner, op.slot, op.color);
    case 'readTextFrame':
      return text.readTextFrame(ctx, op.shape, op.orNull);
    case 'setTextFrame':
      return text.setTextFrame(ctx, op.shape, op.props);
    case 'readText':
      return text.readText(ctx, op.shape, op.span, op.cell, op.orNull);
    case 'setText':
      return text.setText(ctx, op.shape, op.span, op.text, op.cell);
    case 'readFont':
      return text.readFont(ctx, op.shape, op.span, op.cell, op.orNull);
    case 'readTextRuns':
      return text.readTextRuns(ctx, op.shape, op.cell);
    case 'setFont':
      return text.setFont(ctx, op.shape, op.span, op.font, op.cell);
    case 'readParagraph':
      return text.readParagraph(ctx, op.shape, op.span, op.cell, op.orNull);
    case 'setParagraph':
      return text.setParagraph(ctx, op.shape, op.span, op.paragraph, op.cell);
    case 'setHyperlink':
      return text.setHyperlink(ctx, op.shape, op.span, op.address, op.screenTip);
    case 'readTags':
      return props.readTags(ctx, op.owner);
    case 'setTag':
      return props.setTag(ctx, op.owner, op.key, op.value);
    case 'deleteTag':
      return props.deleteTag(ctx, op.owner, op.key);
    case 'readTable':
      return tables.readTableValue(ctx, op.shape);
    case 'readTableCell':
      return tables.readTableCell(ctx, op.shape, op.row, op.column);
    case 'setTableCell':
      return tables.setTableCell(ctx, op.shape, op.row, op.column, op.text);
    case 'insertTableRows':
      return tables.insertTableLines(ctx, op.shape, 'row', op.index, op.count);
    case 'insertTableColumns':
      return tables.insertTableLines(ctx, op.shape, 'column', op.index, op.count);
    case 'deleteTableRows':
      return tables.deleteTableLines(ctx, op.shape, 'row', op.index, op.count);
    case 'deleteTableColumns':
      return tables.deleteTableLines(ctx, op.shape, 'column', op.index, op.count);
    case 'mergeTableCells':
      return tables.mergeTableCells(ctx, op.shape, op.row, op.column, op.rowCount, op.columnCount);
    case 'setTableCellAnchor':
      return tables.setTableCellAnchor(
        ctx,
        op.shape,
        op.row,
        op.column,
        op.anchor,
        op.anchorCenter
      );
    case 'splitTableCell':
      return tables.splitTableCell(ctx, op.shape, op.row, op.column, op.rowCount, op.columnCount);
    case 'resizeTableCell':
      return tables.resizeTableCell(ctx, op.shape, op.row, op.column, op.rowCount, op.columnCount);
    case 'clearTable':
      return tables.clearTable(ctx, op.shape, op.text, op.format);
    case 'setTableSize':
      return tables.setTableSize(ctx, op.shape, op.target, op.index, op.size);
    case 'readTableCellFormat':
      return tableFormat.readTableCellFormat(ctx, op.shape, op.cell);
    case 'setTableCellBorder':
      return tableFormat.setTableCellBorder(ctx, op.shape, op.cell, op.side, op.line);
    case 'setTableCellMargins':
      return tableFormat.setTableCellMargins(ctx, op.shape, op.cell, op.margins);
    case 'setTableStyle':
      return tableFormat.setTableStyle(ctx, op.shape, op.style);
    case 'readAdjustments':
      return adjustments.readAdjustments(ctx, op.shape);
    case 'setAdjustment':
      return adjustments.setAdjustment(ctx, op.shape, op.index, op.value);
    case 'readSelection':
      return readSelection(ctx);
    case 'selectSlides': {
      const port = selectionPort(ctx);
      if (!Array.isArray(op.slides)) badArgument('slides');
      const ids = op.slides.map((ref) => targetOf(ctx, ref, 'slide').sldId);
      for (const id of ids) slideOf(ctx.batch.pkg, id);
      ctx.afterCommit.push(() => port.selectSlides(ids));
      if (ctx.selectionAfter)
        ctx.selectionAfter.value = { slides: [...new Set(ids)], shapes: null, text: null };
      return { kind: 'applied' };
    }
    case 'selectShapes': {
      const port = selectionPort(ctx);
      const at = shapesOnOneSlide(ctx, op.shapes);
      const sldId = at[0]?.owner.kind === 'slide' ? at[0].owner.sldId : null;
      if (sldId) {
        const shapeIds = at.map((s) => s.shapeId);
        ctx.afterCommit.push(() => port.selectShapes(sldId, shapeIds));
        if (ctx.selectionAfter)
          ctx.selectionAfter.value = { slides: [sldId], shapes: { sldId, shapeIds }, text: null };
      }
      return { kind: 'applied' };
    }
    case 'selectText': {
      const port = selectionPort(ctx);
      const at = shapeAt(ctx, op.shape);
      if (at.owner.kind !== 'slide')
        throw new AutomationFailure('invalid-argument', 'select text on a slide', 'shape');
      const whole = text.readText(ctx, op.shape);
      const total = whole.kind === 'text' ? whole.text.length : 0;
      const span = op.span ?? { start: 0, length: total };
      text.readText(ctx, op.shape, span);
      const sldId = at.owner.sldId;
      ctx.afterCommit.push(() => port.selectText(sldId, at.shapeId, span.start, span.length));
      if (ctx.selectionAfter)
        ctx.selectionAfter.value = {
          slides: [sldId],
          shapes: { sldId, shapeIds: [at.shapeId] },
          text: { sldId, shapeId: at.shapeId, start: span.start, length: span.length },
        };
      return { kind: 'applied' };
    }
    default:
      throw new AutomationFailure(
        'unknown-operation',
        'not an operation this host defines',
        String((op as { op?: unknown }).op)
      );
  }
}
