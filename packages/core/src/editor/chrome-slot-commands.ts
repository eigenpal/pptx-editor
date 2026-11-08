// The command each chrome slot dispatches: the fixed command of a command slot, and the command
// a value slot builds from a picked value. Split from `chrome-controls.ts`, which re-exports
// `commandForSlot` and `commandForSlotValue`, so the registry stays under the line cap.

import type {
  Editor,
  EditorCommand,
  FontFormatCommand,
  ParagraphFormatCommand,
  TextCase,
  LineDash,
  LineEndType,
  LineKind,
  ObjectAlign,
  OrderMove,
  TextAlign,
} from '../contracts/index.ts';
import type { ChromeSlotId } from './chrome-controls.ts';
import { shapeFormatSlotCommand } from './chrome-shapes.ts';
import { SLIDE_SLOT_COMMANDS } from './chrome-slide-groups.ts';
import { TABLE_SLOT_COMMANDS, tableCommandForSlotValue } from './chrome-table.ts';
import { isColorChoice } from '../automation/commands/style-commands.ts';
import { isChartType } from '../automation/commands/chart-type.ts';
import { imageSlotCommand } from './chrome-picture.ts';
import { DIAGRAM_SLOT_COMMANDS } from './chrome-diagram.ts';
import { LINE_DASHES } from './line-presets.ts';
import { tableBorderPen } from './chrome-table.ts';
import { TEXT_CASES } from '../automation/commands/font-format.ts';
import { BULLET_GALLERY, NUMBER_GALLERY, galleryFormat } from './bullet-format.ts';

/** The line kinds, line ends, and dashes the shape controls offer. @internal */
export const LINE_KINDS: readonly LineKind[] = ['line', 'arrow', 'elbow', 'curved'];
/** @internal */
export const LINE_ENDS: readonly LineEndType[] = [
  'none',
  'triangle',
  'arrow',
  'stealth',
  'oval',
  'diamond',
];
export { LINE_DASHES, LINE_WIDTHS } from './line-presets.ts';

const SLOT_COMMANDS: Partial<Record<ChromeSlotId, EditorCommand>> = {
  'history.undo': { type: 'undo' },
  'history.redo': { type: 'redo' },
  'format.painter': { type: 'toggleFormatPainter' },
  'text.bold': { type: 'toggleBold' },
  'text.italic': { type: 'toggleItalic' },
  'text.underline': { type: 'toggleUnderline' },
  'text.strike': { type: 'toggleStrike' },
  'format.clear': { type: 'clearFormatting' },
  'script.super': { type: 'toggleSuperscript' },
  'script.sub': { type: 'toggleSubscript' },
  'list.bullet': { type: 'toggleBullets' },
  'list.numbered': { type: 'toggleNumbering' },
  'list.outdent': { type: 'indent', direction: -1 },
  'list.indent': { type: 'indent', direction: 1 },
  'insert.textBox': { type: 'insertTextBox' },
  'arrange.group': { type: 'group' },
  'arrange.ungroup': { type: 'ungroup' },
  'more.duplicate': { type: 'duplicate' },
  'more.delete': { type: 'delete' },
  'view.smartGuides': { type: 'toggleSmartGuides' },
  'image.crop': { type: 'toggleCrop' },
  ...SLIDE_SLOT_COMMANDS,
  ...TABLE_SLOT_COMMANDS,
  ...DIAGRAM_SLOT_COMMANDS,
};

/** The fixed command behind a command slot, or null. @public */
export function commandForSlot(slot: ChromeSlotId): EditorCommand | null {
  return SLOT_COMMANDS[slot] ?? null;
}

/**
 * The command a value slot dispatches for a picked value, or null when the value is invalid.
 * Most slots take the value of one of their options. `font.size` takes a number of points,
 * or `{ step: 1 }` / `{ step: -1 }` for Increase and Decrease Font Size (the size
 * list). The table border slots draw with the border pen of `editor` (`tableBorderPen`), or
 * with the default pen when no editor is given. @public
 */
export function commandForSlotValue(
  slot: ChromeSlotId,
  value: unknown,
  editor?: Editor | null
): EditorCommand | null {
  switch (slot) {
    case 'font.family':
      return typeof value === 'string' && value ? { type: 'setFontFamily', family: value } : null;
    case 'font.size': {
      if (typeof value === 'number') return { type: 'setFontSize', points: value };
      // A number is points. The stepper's buttons send `{ step: 1 }` or `{ step: -1 }`
      // (Increase and Decrease Font Size); a bare number never steps.
      const step = typeof value === 'object' && value !== null ? Reflect.get(value, 'step') : 0;
      return step === 1 || step === -1 ? { type: 'stepFontSize', direction: step } : null;
    }
    case 'text.color':
      return isColorChoice(value) || value === null ? { type: 'setTextColor', color: value } : null;
    case 'text.highlight':
      return isColorChoice(value) || value === null ? { type: 'setHighlight', color: value } : null;
    case 'paragraph.align':
      return { type: 'setAlign', align: value as TextAlign };
    case 'list.lineSpacing':
      return { type: 'setLineSpacing', multiple: Number(value) };
    case 'text.characterSpacing': {
      const points = typeof value === 'string' && value.trim() ? Number(value) : value;
      return typeof points === 'number' && Number.isFinite(points)
        ? { type: 'setCharacterSpacing', points }
        : null;
    }
    case 'text.changeCase':
      return TEXT_CASES.includes(value as TextCase)
        ? { type: 'changeCase', case: value as TextCase }
        : null;
    case 'text.dialog':
      // The type goes last, so a value's own `type` key can never pick another command.
      return typeof value === 'object' && value !== null
        ? { ...(value as FontFormatCommand), type: 'setFontFormat' }
        : null;
    case 'list.bulletStyle':
    case 'list.numberStyle': {
      const format = galleryFormat(
        slot === 'list.bulletStyle' ? BULLET_GALLERY : NUMBER_GALLERY,
        value
      );
      return format === undefined ? null : { type: 'setBulletFormat', format };
    }
    case 'paragraph.autofit':
      return value === 'none' || value === 'shrink' || value === 'shape'
        ? { type: 'setAutofit', mode: value }
        : null;
    case 'paragraph.textDirection':
      return typeof value === 'string'
        ? ({ type: 'setTextDirection', direction: value } as EditorCommand)
        : null;
    case 'paragraph.columns': {
      // A number or a numeric string, as the menu rows hold; never `true` or an array.
      const count =
        typeof value === 'number' ? value : typeof value === 'string' ? Number(value) : NaN;
      return Number.isInteger(count) ? { type: 'setColumns', count } : null;
    }
    case 'paragraph.dialog':
      // The type goes last, so a value's own `type` key can never pick another command.
      return typeof value === 'object' && value !== null
        ? { ...(value as ParagraphFormatCommand), type: 'setParagraphFormat' }
        : null;
    case 'insert.shape':
      return typeof value === 'string' ? { type: 'insertShape', preset: value } : null;
    case 'insert.equation':
      return typeof value === 'string' ? { type: 'insertEquation', linear: value } : null;
    case 'insert.chart':
      return isChartType(value) ? { type: 'insertChart', chartType: value } : null;
    case 'shape.fill':
      return isColorChoice(value) || value === null ? { type: 'setShapeFill', color: value } : null;
    case 'shape.lineColor':
      return isColorChoice(value) || value === null ? { type: 'setLineColor', color: value } : null;
    case 'shape.lineWidth': {
      const points = Number(value);
      return points > 0 ? { type: 'setLineWidth', points } : null;
    }
    case 'text.link':
      return typeof value === 'string' || value === null ? { type: 'setLink', url: value } : null;
    case 'paragraph.valign':
      return value === 't' || value === 'ctr' || value === 'b'
        ? { type: 'setVerticalAlign', anchor: value }
        : null;
    case 'insert.line':
      return LINE_KINDS.includes(value as LineKind)
        ? { type: 'insertLine', kind: value as LineKind }
        : null;
    case 'shape.lineStart':
      return LINE_ENDS.includes(value as LineEndType)
        ? { type: 'setLineEnds', head: value as LineEndType }
        : null;
    case 'shape.lineEnd':
      return LINE_ENDS.includes(value as LineEndType)
        ? { type: 'setLineEnds', tail: value as LineEndType }
        : null;
    case 'shape.lineDash':
      return LINE_DASHES.includes(value as LineDash)
        ? { type: 'setLineDash', dash: value as LineDash }
        : null;
    case 'arrange.order':
      return { type: 'order', move: value as OrderMove };
    case 'arrange.align':
      return value === 'distributeHorizontal' || value === 'distributeVertical'
        ? {
            type: 'distributeObjects',
            axis: value === 'distributeHorizontal' ? 'horizontal' : 'vertical',
          }
        : { type: 'alignObjects', align: value as ObjectAlign };
    case 'view.presentationView':
      return value === 'normal' || value === 'slideSorter'
        ? { type: 'setPresentationView', view: value }
        : null;
    case 'slide.add':
      return typeof value === 'string' && value ? { type: 'addSlide', layoutId: value } : null;
    case 'slide.layout':
      return typeof value === 'string' && value ? { type: 'applyLayout', layoutId: value } : null;
    default: {
      // `null` is an invalid value of a Shape Format slot; `undefined` is a slot it does not own.
      const shaped = shapeFormatSlotCommand(slot, value);
      if (shaped !== undefined) return shaped;
      return (
        imageSlotCommand(slot, value) ??
        tableCommandForSlotValue(slot, value, tableBorderPen(editor)) ??
        commandForSlot(slot)
      );
    }
  }
}
