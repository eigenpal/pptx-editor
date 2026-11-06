/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Table format: Office.js `Borders`, `Border`, and `Margins` over one cell's `a:tcPr` (ECMA-376
// Part 1, 21.1.3.17), and `TableStyleSettings` over `a:tblPr` (21.1.3.15). Weights and margins
// are in points; transparency is 0 (opaque) to 1 (clear).

import type {
  AutomationCellAddress,
  AutomationCellBorderSide,
  AutomationLineRead,
  AutomationLineWrite,
  AutomationTableCellFormatRead,
  AutomationTableStyleWrite,
} from '@pptx-editor.dev/core/automation';
import { diagnosticOf, fail, failArgument } from '../runtime/errors.ts';
import { isDefaultLoad, type ResolvedLoadOptions } from '../runtime/load-options.ts';
import { ObjectPath } from '../runtime/object-path.ts';
import type { RequestContext } from '../runtime/request-context.ts';
import {
  DASH_STYLE,
  tableStyleName,
  tableStyleOf,
  type ShapeLineDashStyle,
  type TableStyle,
} from './enums.ts';
import { fraction } from './fill-line.ts';
import { finite, flag, ModelObject, text } from './model-object.ts';

/** The cell a format object writes through: the table shape and a grid position. */
interface CellTarget {
  readonly shape: ObjectPath;
  readonly cell: AutomationCellAddress;
}

/** An object over one cell's `a:tcPr`. */
abstract class CellFormatPart extends ModelObject {
  /** @internal */
  protected readonly target: CellTarget;

  /** @internal */
  constructor(context: RequestContext, path: ObjectPath, target: CellTarget) {
    super(context, path);
    this.target = target;
  }

  /** @internal Queue the cell format read that fills `names`. */
  protected loadCellFormat(
    names: readonly string[],
    project: (format: AutomationTableCellFormatRead) => Readonly<Record<string, unknown>>
  ): void {
    const { shape, cell } = this.target;
    this.loadScalars(
      names,
      'tableCellFormat',
      (ref) => ({ op: 'readTableCellFormat', shape: ref(shape), cell }),
      (v) => project(v.format)
    );
  }
}

/** The values of one border, or undefined for each when the cell shows no line there. */
function borderValues(line: AutomationLineRead): Readonly<Record<string, unknown>> {
  if (!line.visible)
    return { color: undefined, dashStyle: undefined, transparency: undefined, weight: undefined };
  return {
    color: line.color ?? undefined,
    dashStyle: DASH_STYLE.fromFile(line.dashStyle) ?? 'Solid',
    transparency: line.transparency ?? 0,
    weight: line.weight ?? undefined,
  };
}

const BORDER_NAMES = ['color', 'dashStyle', 'transparency', 'weight'] as const;

/**
 * One border of a table cell: an edge or a diagonal. It reads the line the cell shows, its own
 * or the table style's; a cell with no line there reads `undefined` for each property. A write
 * keeps the properties it does not set, and an outside edge also changes the cell across it.
 *
 * @public
 */
export class Border extends CellFormatPart {
  readonly #side: AutomationCellBorderSide;

  /** @internal */
  constructor(
    context: RequestContext,
    path: ObjectPath,
    target: CellTarget,
    side: AutomationCellBorderSide
  ) {
    super(context, path, target);
    this.#side = side;
  }

  /** The line color as `#RRGGBB`. A write also takes a preset color name. */
  get color(): string | undefined {
    return this.loadedProperty('color');
  }
  set color(value: string | undefined) {
    this.#write('color', { color: text(value, this.at('color')) });
  }
  /** The dash style. */
  get dashStyle(): ShapeLineDashStyle | `${ShapeLineDashStyle}` | undefined {
    return this.loadedProperty('dashStyle');
  }
  set dashStyle(value: ShapeLineDashStyle | `${ShapeLineDashStyle}` | undefined) {
    this.#write('dashStyle', { dashStyle: DASH_STYLE.toFile(value, this.at('dashStyle')) });
  }
  /** The transparency, from 0 (opaque) to 1 (clear). */
  get transparency(): number | undefined {
    return this.loadedProperty('transparency');
  }
  set transparency(value: number | undefined) {
    this.#write('transparency', { transparency: fraction(value, this.at('transparency')) });
  }
  /** The line weight, in points. */
  get weight(): number | undefined {
    return this.loadedProperty('weight');
  }
  set weight(value: number | undefined) {
    const weight = finite(value, this.at('weight'));
    if (weight < 0) failArgument(this.at('weight'), 'a weight of 0 or more points');
    this.#write('weight', { weight });
  }

  #write(name: string, line: AutomationLineWrite): void {
    const { shape, cell } = this.target;
    const side = this.#side;
    this.command(name, (ref) => ({
      op: 'setTableCellBorder',
      shape: ref(shape),
      cell,
      side,
      line,
    }));
  }

  /** @internal */
  protected override onLoad(request: ResolvedLoadOptions): void {
    const side = this.#side;
    this.loadCellFormat(this.selection(request, BORDER_NAMES), (f) =>
      borderValues(f.borders[side])
    );
  }
}

const SIDES = ['left', 'right', 'top', 'bottom', 'diagonalDown', 'diagonalUp'] as const;

/**
 * The borders of a table cell (`a:tcPr/a:lnL`, `a:lnR`, `a:lnT`, `a:lnB`, `a:lnTlToBr`,
 * `a:lnBlToTr`).
 *
 * @public
 */
export class Borders extends CellFormatPart {
  readonly #sides = new Map<AutomationCellBorderSide, Border>();

  /** The bottom edge. */
  get bottom(): Border {
    return this.#side('bottom');
  }
  /** The diagonal from the top-left to the bottom-right corner. */
  get diagonalDown(): Border {
    return this.#side('diagonalDown');
  }
  /** The diagonal from the bottom-left to the top-right corner. */
  get diagonalUp(): Border {
    return this.#side('diagonalUp');
  }
  /** The left edge. */
  get left(): Border {
    return this.#side('left');
  }
  /** The right edge. */
  get right(): Border {
    return this.#side('right');
  }
  /** The top edge. */
  get top(): Border {
    return this.#side('top');
  }

  #side(side: AutomationCellBorderSide): Border {
    let border = this.#sides.get(side);
    if (!border) {
      border = new Border(
        this.context,
        ObjectPath.derived(`${this.path.label}.${side}`, this.path),
        this.target,
        side
      );
      this.#sides.set(side, border);
    }
    return border;
  }

  /** @internal */
  protected override onLoad(request: ResolvedLoadOptions): void {
    const navigation = Object.fromEntries(SIDES.map((side) => [side, () => this.#side(side)]));
    // Every name is a navigation property: an empty selection loads all six borders.
    const names = isDefaultLoad(request) ? [...SIDES] : request.select;
    this.selection({ ...request, select: names }, [], navigation);
  }
}

const MARGIN_NAMES = ['bottom', 'left', 'right', 'top'] as const;

/**
 * The margins of a table cell, in points (`a:tcPr/@marL`, `@marR`, `@marT`, `@marB`). Writing
 * `undefined` removes the margin, back to its default (7.2 points left and right, 3.6 points
 * top and bottom).
 *
 * @public
 */
export class Margins extends CellFormatPart {
  /** The bottom margin, in points. */
  get bottom(): number | undefined {
    return this.loadedProperty('bottom');
  }
  set bottom(value: number | undefined) {
    this.#write('bottom', value);
  }
  /** The left margin, in points. */
  get left(): number | undefined {
    return this.loadedProperty('left');
  }
  set left(value: number | undefined) {
    this.#write('left', value);
  }
  /** The right margin, in points. */
  get right(): number | undefined {
    return this.loadedProperty('right');
  }
  set right(value: number | undefined) {
    this.#write('right', value);
  }
  /** The top margin, in points. */
  get top(): number | undefined {
    return this.loadedProperty('top');
  }
  set top(value: number | undefined) {
    this.#write('top', value);
  }

  #write(name: (typeof MARGIN_NAMES)[number], value: number | undefined): void {
    const margin = value === undefined ? null : finite(value, this.at(name));
    if (margin !== null && margin < 0) failArgument(this.at(name), 'a margin of 0 or more points');
    const { shape, cell } = this.target;
    this.command(name, (ref) => ({
      op: 'setTableCellMargins',
      shape: ref(shape),
      cell,
      margins: { [name]: margin },
    }));
  }

  /** @internal */
  protected override onLoad(request: ResolvedLoadOptions): void {
    this.loadCellFormat(this.selection(request, MARGIN_NAMES), (f) => ({ ...f.margins }));
  }
}

/** The style option flags, by their Office.js names. */
const STYLE_FLAGS = [
  ['areColumnsBanded', 'bandedColumns'],
  ['areRowsBanded', 'bandedRows'],
  ['isFirstColumnHighlighted', 'firstColumn'],
  ['isFirstRowHighlighted', 'firstRow'],
  ['isLastColumnHighlighted', 'lastColumn'],
  ['isLastRowHighlighted', 'lastRow'],
] as const;

/** A style that has no `TableStyle` value; reading `style` refuses. */
const UNKNOWN_STYLE = Symbol('unknown table style');

/**
 * The style options and the style of a table (`a:tblPr`). A style is found by its name in the
 * presentation's `ppt/tableStyles.xml`: `MediumStyle2Accent1` is "Medium Style 2 - Accent 1".
 * The presentation must define the style, except `MediumStyle2Accent1`, which is built in, and
 * `NoStyleNoGrid`, which removes the style choice.
 *
 * @public
 */
export class TableStyleSettings extends ModelObject {
  readonly #shape: ObjectPath;

  /** @internal */
  constructor(context: RequestContext, path: ObjectPath, shape: ObjectPath) {
    super(context, path);
    this.#shape = shape;
  }

  /** Whether the columns show banded formatting (`a:tblPr/@bandCol`). */
  get areColumnsBanded(): boolean {
    return this.loadedProperty('areColumnsBanded');
  }
  set areColumnsBanded(value: boolean) {
    this.#flag('areColumnsBanded', 'bandedColumns', value);
  }
  /** Whether the rows show banded formatting (`a:tblPr/@bandRow`). */
  get areRowsBanded(): boolean {
    return this.loadedProperty('areRowsBanded');
  }
  set areRowsBanded(value: boolean) {
    this.#flag('areRowsBanded', 'bandedRows', value);
  }
  /** Whether the first column has its own format (`a:tblPr/@firstCol`). */
  get isFirstColumnHighlighted(): boolean {
    return this.loadedProperty('isFirstColumnHighlighted');
  }
  set isFirstColumnHighlighted(value: boolean) {
    this.#flag('isFirstColumnHighlighted', 'firstColumn', value);
  }
  /** Whether the first row has its own format (`a:tblPr/@firstRow`). */
  get isFirstRowHighlighted(): boolean {
    return this.loadedProperty('isFirstRowHighlighted');
  }
  set isFirstRowHighlighted(value: boolean) {
    this.#flag('isFirstRowHighlighted', 'firstRow', value);
  }
  /** Whether the last column has its own format (`a:tblPr/@lastCol`). */
  get isLastColumnHighlighted(): boolean {
    return this.loadedProperty('isLastColumnHighlighted');
  }
  set isLastColumnHighlighted(value: boolean) {
    this.#flag('isLastColumnHighlighted', 'lastColumn', value);
  }
  /** Whether the last row has its own format (`a:tblPr/@lastRow`). */
  get isLastRowHighlighted(): boolean {
    return this.loadedProperty('isLastRowHighlighted');
  }
  set isLastRowHighlighted(value: boolean) {
    this.#flag('isLastRowHighlighted', 'lastRow', value);
  }
  /**
   * The table style. A table with no style reads `NoStyleNoGrid`. A style whose name has no
   * `TableStyle` value refuses the read with `NotImplemented`.
   */
  get style(): TableStyle | `${TableStyle}` {
    const value = this.loadedProperty<`${TableStyle}` | typeof UNKNOWN_STYLE>('style');
    if (value === UNKNOWN_STYLE)
      fail({
        code: 'NotImplemented',
        target: this.at('style'),
        diagnostic: diagnosticOf('table-style-unknown')!,
      });
    return value;
  }
  set style(value: TableStyle | `${TableStyle}`) {
    this.#write('style', { styleName: tableStyleName(value, this.at('style')) });
  }

  #flag(
    name: string,
    key: Exclude<keyof AutomationTableStyleWrite, 'styleName'>,
    value: boolean
  ): void {
    this.#write(name, { [key]: flag(value, this.at(name)) });
  }

  #write(name: string, style: AutomationTableStyleWrite): void {
    const shape = this.#shape;
    this.command(name, (ref) => ({ op: 'setTableStyle', shape: ref(shape), style }));
  }

  /** @internal */
  protected override onLoad(request: ResolvedLoadOptions): void {
    const names = this.selection(request, [...STYLE_FLAGS.map(([n]) => n), 'style']);
    const shape = this.#shape;
    this.loadScalars(
      names,
      'table',
      (ref) => ({ op: 'readTable', shape: ref(shape) }),
      ({ table: { style } }) => {
        const out: Record<string, unknown> = {};
        for (const [name, key] of STYLE_FLAGS) out[name] = style[key];
        out.style =
          style.styleId === null && style.styleName === null
            ? 'NoStyleNoGrid'
            : ((style.styleName === null ? null : tableStyleOf(style.styleName)) ?? UNKNOWN_STYLE);
        return out;
      }
    );
  }
}
