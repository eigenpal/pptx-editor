/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Office.js `TableCell` over one `a:tc` (ECMA-376 Part 1, 21.1.3.16). A covered grid position
// addresses the anchor cell of the merged block that covers it. The cell's text body takes the
// same text, font, and paragraph operations as a shape's, with the cell as their target.

import type {
  AutomationCellAddress,
  AutomationFontWrite,
  AutomationOperation,
  AutomationTextRunRead,
} from '@pptx-editor.dev/core/automation';
import { failArgument, nearest, oneOf } from '../runtime/errors.ts';
import type { ResolvedLoadOptions } from '../runtime/load-options.ts';
import { ObjectPath } from '../runtime/object-path.ts';
import type { RequestContext } from '../runtime/request-context.ts';
import {
  HORIZONTAL_ALIGNMENT,
  UNDERLINE,
  VERTICAL_ALIGNMENT,
  type ParagraphHorizontalAlignment,
  type ShapeFontUnderlineStyle,
  type TextVerticalAlignment,
} from './enums.ts';
import { ShapeFill } from './fill-line.ts';
import { finite, ModelObject, text, type RefOf } from './model-object.ts';
import { Borders, Margins } from './table-format.ts';
import { ParagraphFormat, ShapeFont } from './text.ts';

/**
 * Character formatting for one run of `TableCell.textRuns`. In a write, an absent property
 * keeps the cell's format. A read gives every property, except `color` when the text has no
 * solid fill.
 *
 * @public
 */
export interface FontProperties {
  allCaps?: boolean | undefined;
  bold?: boolean | undefined;
  /** `#RRGGBB` or a preset color name. */
  color?: string | undefined;
  doubleStrikethrough?: boolean | undefined;
  italic?: boolean | undefined;
  name?: string | undefined;
  /** Points. */
  size?: number | undefined;
  smallCaps?: boolean | undefined;
  strikethrough?: boolean | undefined;
  subscript?: boolean | undefined;
  superscript?: boolean | undefined;
  underline?: ShapeFontUnderlineStyle | `${ShapeFontUnderlineStyle}` | undefined;
}

/**
 * One run of text with its own format, for `TableCell.textRuns`.
 *
 * @public
 */
export interface TextRun {
  font?: FontProperties | undefined;
  text: string;
}

const FONT_KEYS = new Set([
  'allCaps',
  'bold',
  'color',
  'doubleStrikethrough',
  'italic',
  'name',
  'size',
  'smallCaps',
  'strikethrough',
  'subscript',
  'superscript',
  'underline',
]);

/** The checked font write of one run, or undefined when the run keeps the cell's format. */
function runFont(
  font: FontProperties | undefined,
  target: string
): AutomationFontWrite | undefined {
  if (font === undefined) return undefined;
  if (typeof font !== 'object' || font === null)
    failArgument(target, 'an object of font properties');
  const out: Record<string, unknown> = {};
  for (const [key, value] of Object.entries(font)) {
    if (!FONT_KEYS.has(key))
      failArgument(
        `${target}.${key}`,
        `a font property: ${oneOf([...FONT_KEYS])}`,
        nearest(key, [...FONT_KEYS])
      );
    if (value === undefined) continue;
    out[key] = key === 'underline' ? UNDERLINE.toFile(value, `${target}.underline`) : value;
  }
  return Object.keys(out).length === 0 ? undefined : (out as AutomationFontWrite);
}

/** A positive whole number, or `InvalidArgument`. */
function count(value: unknown, target: string): number {
  if (typeof value !== 'number' || !Number.isInteger(value) || value < 1)
    failArgument(target, 'a whole number of 1 or more');
  return value;
}

/** The span of a cell a collection read already knows. */
export interface CellSpan {
  readonly rowIndex: number;
  readonly columnIndex: number;
  readonly rowCount: number;
  readonly columnCount: number;
}

/** Fills the span of a cell from a collection read. Module-private: no consumer can call it. */
export const CELL_SPANS = new WeakMap<TableCell, (span: CellSpan) => void>();

/**
 * One cell of a table.
 *
 * @public
 */
export class TableCell extends ModelObject {
  readonly #cell: AutomationCellAddress;
  #font: ShapeFont | undefined;
  #fill: ShapeFill | undefined;
  #paragraph: ParagraphFormat | undefined;
  #borders: Borders | undefined;
  #margins: Margins | undefined;

  /** @internal */
  constructor(context: RequestContext, path: ObjectPath, row: number, column: number) {
    super(context, path, { nullable: true });
    this.#cell = Object.freeze({ row, column });
    CELL_SPANS.set(this, (span) => {
      for (const [name, value] of Object.entries(span)) this.setLoadedProperty(name, value);
    });
  }

  /** The text of the cell. */
  get text(): string {
    return this.loadedProperty('text');
  }
  set text(value: string) {
    const body = text(value, this.at('text'));
    this.#command('text', { op: 'setTableCell', ...this.#cell, text: body });
  }
  /** The zero-based row of the cell; the first row of its merged area. */
  get rowIndex(): number {
    return this.loadedProperty('rowIndex');
  }
  /** The zero-based column of the cell; the first column of its merged area. */
  get columnIndex(): number {
    return this.loadedProperty('columnIndex');
  }
  /** How many rows the cell spans. */
  get rowCount(): number {
    return this.loadedProperty('rowCount');
  }
  /** How many columns the cell spans. */
  get columnCount(): number {
    return this.loadedProperty('columnCount');
  }
  /** The character format of the cell's text. */
  get font(): ShapeFont {
    return (this.#font ??= new ShapeFont(
      this.context,
      ObjectPath.derived(`${this.path.label}.font`, this.path),
      this.path,
      () => undefined,
      this.#cell
    ));
  }
  /** The borders of the cell (`a:tcPr` lines). */
  get borders(): Borders {
    return (this.#borders ??= new Borders(
      this.context,
      ObjectPath.derived(`${this.path.label}.borders`, this.path),
      { shape: this.path, cell: this.#cell }
    ));
  }
  /** The margins of the cell (`a:tcPr/@marL` and the like), in points. */
  get margins(): Margins {
    return (this.#margins ??= new Margins(
      this.context,
      ObjectPath.derived(`${this.path.label}.margins`, this.path),
      { shape: this.path, cell: this.#cell }
    ));
  }
  /** The cell's fill (`a:tcPr`). */
  get fill(): ShapeFill {
    return (this.#fill ??= new ShapeFill(
      this.context,
      ObjectPath.derived(`${this.path.label}.fill`, this.path),
      { kind: 'cell', shape: this.path, cell: this.#cell }
    ));
  }
  /** The horizontal alignment of the cell's paragraphs (`a:pPr/@algn`). */
  get horizontalAlignment():
    ParagraphHorizontalAlignment | `${ParagraphHorizontalAlignment}` | null {
    return this.loadedProperty('horizontalAlignment');
  }
  set horizontalAlignment(
    value: ParagraphHorizontalAlignment | `${ParagraphHorizontalAlignment}` | null
  ) {
    this.#paragraphs().horizontalAlignment = value;
  }
  /** The indent level of the cell's paragraphs, from 1; null when they differ. */
  get indentLevel(): number | null {
    return this.loadedProperty('indentLevel');
  }
  set indentLevel(value: number | null) {
    this.#paragraphs().indentLevel = finite(value, this.at('indentLevel'));
  }
  /**
   * The vertical alignment of the cell's text (`a:tcPr/@anchor`); the `…Centered` values also
   * set `@anchorCtr` (ECMA-376 Part 1, 21.1.3.17).
   */
  get verticalAlignment(): TextVerticalAlignment | `${TextVerticalAlignment}` {
    return this.loadedProperty('verticalAlignment');
  }
  set verticalAlignment(value: TextVerticalAlignment | `${TextVerticalAlignment}`) {
    const file = VERTICAL_ALIGNMENT.toFile(value, `${this.path.label}.verticalAlignment`);
    const anchorCenter = file.endsWith('Centered');
    const anchor = anchorCenter ? file.slice(0, -'Centered'.length) : file;
    this.#command('verticalAlignment', {
      op: 'setTableCellAnchor',
      ...this.#cell,
      anchor,
      anchorCenter,
    });
  }
  /**
   * The cell's text as runs, each with its own character format. Neighbouring text with the
   * same format reads as one run, and the run texts join to `text`. A write replaces the text
   * with the runs in one sync.
   */
  get textRuns(): TextRun[] {
    return this.loadedProperty('textRuns');
  }
  set textRuns(value: TextRun[]) {
    const target = `${this.path.label}.textRuns`;
    if (!Array.isArray(value)) failArgument(target, 'an array of text runs');
    let start = 0;
    const formats: { span: { start: number; length: number }; font: AutomationFontWrite }[] = [];
    const whole = value
      .map((run, i) => {
        if (typeof run !== 'object' || run === null || typeof run.text !== 'string')
          failArgument(`${target}[${i}]`, 'a text run with a text string');
        const font = runFont(run.font, `${target}[${i}].font`);
        // The engine writes one paragraph break for `\r\n`; count it as one, so the spans
        // that follow still match the written text.
        const piece = run.text.replace(/\r\n|\n/g, '\r');
        if (font && piece.length > 0) formats.push({ span: { start, length: piece.length }, font });
        start += piece.length;
        return piece;
      })
      .join('');
    const cell = this.#cell;
    this.command('textRuns', (ref) => ({
      op: 'setText',
      shape: ref(this.path),
      text: whole,
      cell,
    }));
    for (const { span, font } of formats)
      this.command('textRuns', (ref) => ({
        op: 'setFont',
        shape: ref(this.path),
        span,
        font,
        cell,
      }));
  }

  /** Make the cell span `rowCount` rows and `columnCount` columns from its first grid position. */
  resize(rowCount: number, columnCount: number): void {
    const target = `${this.path.label}.resize`;
    this.#command('resize', {
      op: 'resizeTableCell',
      ...this.#cell,
      rowCount: count(rowCount, target),
      columnCount: count(columnCount, target),
    });
  }

  /** Split the cell into `rowCount` rows and `columnCount` columns of cells. */
  split(rowCount: number, columnCount: number): void {
    const target = `${this.path.label}.split`;
    this.#command('split', {
      op: 'splitTableCell',
      ...this.#cell,
      rowCount: count(rowCount, target),
      columnCount: count(columnCount, target),
    });
  }

  #paragraphs(): ParagraphFormat {
    return (this.#paragraph ??= new ParagraphFormat(
      this.context,
      ObjectPath.derived(`${this.path.label}.paragraphFormat`, this.path),
      this.path,
      () => undefined,
      this.#cell
    ));
  }

  /** Queue a command on the table shape: `op` without its `shape`. */
  #command(name: string, op: DistributiveOmit<AutomationOperation, 'shape'>): void {
    this.command(name, (ref: RefOf) => ({ ...op, shape: ref(this.path) }) as AutomationOperation);
  }

  /** @internal */
  protected override onLoad(request: ResolvedLoadOptions): void {
    const names = this.selection(
      request,
      [
        'text',
        'rowIndex',
        'columnIndex',
        'rowCount',
        'columnCount',
        'verticalAlignment',
        'horizontalAlignment',
        'indentLevel',
        'textRuns',
      ],
      {
        font: () => this.font,
        fill: () => this.fill,
        borders: () => this.borders,
        margins: () => this.margins,
      }
    );
    const cell = this.#cell;
    const fromCell = names.filter(
      (n) => n !== 'horizontalAlignment' && n !== 'indentLevel' && n !== 'textRuns'
    );
    this.loadScalars(
      fromCell,
      'tableCell',
      (ref) => ({ op: 'readTableCell', shape: ref(this.path), ...cell }),
      (v) => ({
        ...v.cell,
        verticalAlignment: VERTICAL_ALIGNMENT.fromFile(
          v.cell.anchor + (v.cell.anchorCenter ? 'Centered' : '')
        ),
      })
    );
    this.loadScalars(
      names.filter((n) => n === 'textRuns'),
      'textRuns',
      (ref) => ({ op: 'readTextRuns', shape: ref(this.path), cell }),
      (v) => ({ textRuns: v.runs.map(runOf) })
    );
    const fromParagraphs = names.filter((n) => n === 'horizontalAlignment' || n === 'indentLevel');
    this.loadScalars(
      fromParagraphs,
      'paragraph',
      (ref) => ({ op: 'readParagraph', shape: ref(this.path), cell }),
      (v) => ({
        horizontalAlignment: HORIZONTAL_ALIGNMENT.fromFile(v.paragraph.alignment),
        indentLevel: v.paragraph.indentLevel,
      })
    );
  }
}

/** A run the host read, as the Office.js `TextRun`: a font property that is unknown is absent. */
function runOf(run: AutomationTextRunRead): TextRun {
  const { underline, ...rest } = run.font;
  const font: FontProperties = {};
  for (const [key, value] of Object.entries(rest))
    if (value !== null) (font as Record<string, unknown>)[key] = value;
  const style = underline === null ? null : UNDERLINE.fromFile(underline);
  if (style) font.underline = style;
  return { text: run.text, font };
}

/** `Omit` over each member of a union. */
type DistributiveOmit<T, K extends PropertyKey> = T extends unknown ? Omit<T, K> : never;
