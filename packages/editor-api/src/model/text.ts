/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Text: Office.js `TextFrame`, `TextRange`, `ShapeFont`, `ParagraphFormat`, and `BulletFormat`.
//
// A text range is a shape and a span: UTF-16 offsets over the paragraphs joined with `\r`, where
// a line break (`a:br`) reads as `\v`. A text write re-fits the text body by its autofit rules
// (ECMA-376 Part 1, 21.1.2.1) when the host can measure text.
//
// A range whose span is known only after a read (the reader's selected text, a substring to the
// end of the text) is over a pending path that read resolves; a call on it in the same sync
// waits for that read, which the sync sends in an earlier round.

import type {
  AutomationCellAddress,
  AutomationHandle,
  AutomationFontWrite,
  AutomationOperation,
  AutomationParagraphWrite,
  AutomationTextFrameWrite,
  AutomationTextSpan,
} from '@pptx-editor.dev/core/automation';
import { fail, failArgument } from '../runtime/errors.ts';
import { hydrated } from '../runtime/hydrate.ts';
import type { ResolvedLoadOptions } from '../runtime/load-options.ts';
import { ObjectPath } from '../runtime/object-path.ts';
import type { RequestContext } from '../runtime/request-context.ts';
import {
  AUTO_SIZE,
  BULLET_STYLE,
  BULLET_TYPE,
  HORIZONTAL_ALIGNMENT,
  UNDERLINE,
  VERTICAL_ALIGNMENT,
  type BulletStyle,
  type BulletType,
  type ParagraphHorizontalAlignment,
  type ShapeAutoSize,
  type ShapeFontUnderlineStyle,
  type TextVerticalAlignment,
} from './enums.ts';
import { Hyperlink, HyperlinkScopedCollection, type HyperlinkAddOptions } from './hyperlinks.ts';
import { finite, flag, ModelObject, text, type RefOf } from './model-object.ts';
import { Shape } from './shapes.ts';

const FRAME_NAMES = [
  'autoSizeSetting',
  'bottomMargin',
  'leftMargin',
  'rightMargin',
  'topMargin',
  'verticalAlignment',
  'wordWrap',
  'hasText',
] as const;

/**
 * The text frame of a shape (`p:txBody` and its `a:bodyPr`).
 *
 * @public
 */
export class TextFrame extends ModelObject {
  readonly #shape: ObjectPath;
  #range: TextRange | undefined;

  /** @internal */
  constructor(context: RequestContext, path: ObjectPath, shape: ObjectPath, nullable = false) {
    super(context, path, { nullable });
    this.#shape = shape;
  }

  /** How the text and the shape fit each other. */
  get autoSizeSetting(): ShapeAutoSize | `${ShapeAutoSize}` {
    return this.loadedProperty('autoSizeSetting');
  }
  set autoSizeSetting(value: ShapeAutoSize | `${ShapeAutoSize}`) {
    const autoSize = AUTO_SIZE.toFile(value, 'autoSizeSetting') as 'none' | 'normal' | 'shape';
    this.#write('autoSizeSetting', { autoSize });
  }
  /** The bottom inset, in points. */
  get bottomMargin(): number {
    return this.loadedProperty('bottomMargin');
  }
  set bottomMargin(value: number) {
    this.#write('bottomMargin', { bottomMargin: finite(value, this.at('bottomMargin')) });
  }
  /** The left inset, in points. */
  get leftMargin(): number {
    return this.loadedProperty('leftMargin');
  }
  set leftMargin(value: number) {
    this.#write('leftMargin', { leftMargin: finite(value, this.at('leftMargin')) });
  }
  /** The right inset, in points. */
  get rightMargin(): number {
    return this.loadedProperty('rightMargin');
  }
  set rightMargin(value: number) {
    this.#write('rightMargin', { rightMargin: finite(value, this.at('rightMargin')) });
  }
  /** The top inset, in points. */
  get topMargin(): number {
    return this.loadedProperty('topMargin');
  }
  set topMargin(value: number) {
    this.#write('topMargin', { topMargin: finite(value, this.at('topMargin')) });
  }
  /** The vertical anchor of the text. */
  get verticalAlignment(): TextVerticalAlignment | `${TextVerticalAlignment}` {
    return this.loadedProperty('verticalAlignment');
  }
  set verticalAlignment(value: TextVerticalAlignment | `${TextVerticalAlignment}`) {
    this.#write('verticalAlignment', {
      verticalAlignment: VERTICAL_ALIGNMENT.toFile(value, 'verticalAlignment'),
    });
  }
  /** Whether lines wrap at the frame width. */
  get wordWrap(): boolean {
    return this.loadedProperty('wordWrap');
  }
  set wordWrap(value: boolean) {
    this.#write('wordWrap', { wordWrap: flag(value, this.at('wordWrap')) });
  }
  /** Whether the frame holds any text. */
  get hasText(): boolean {
    return this.loadedProperty('hasText');
  }

  /** All the text of the frame. */
  get textRange(): TextRange {
    return (this.#range ??= new TextRange(
      this.context,
      ObjectPath.derived(`${this.path.label}.textRange`, this.path),
      this.#shape,
      undefined
    ));
  }

  /** Delete all the text. */
  deleteText(): void {
    const shape = this.#shape;
    this.command('deleteText', (ref) => ({ op: 'setText', shape: ref(shape), text: '' }));
  }

  /** The shape that holds the frame. */
  getParentShape(): Shape {
    return new Shape(this.context, this.#shape);
  }

  #write(name: string, props: AutomationTextFrameWrite): void {
    const shape = this.#shape;
    this.command(name, (ref) => ({ op: 'setTextFrame', shape: ref(shape), props }));
  }

  /** @internal */
  protected override onLoad(request: ResolvedLoadOptions): void {
    const names = this.selection(request, FRAME_NAMES, { textRange: () => this.textRange });
    const shape = this.#shape;
    this.loadScalars(
      names,
      'textFrame',
      (ref) => ({ op: 'readTextFrame', shape: ref(shape) }),
      (v) => ({
        ...v.textFrame,
        autoSizeSetting: AUTO_SIZE.fromFile(v.textFrame.autoSize) ?? 'AutoSizeMixed',
        verticalAlignment: VERTICAL_ALIGNMENT.fromFile(v.textFrame.verticalAlignment) ?? 'Top',
      })
    );
  }
}

/** The table cell a text operation addresses, spread into it; empty for a shape's own text. */
type CellArg = { readonly cell?: AutomationCellAddress };

/** A span, or the whole text when absent. */
type Span = AutomationTextSpan | undefined;

/** The read of a paragraph format, planned against the batch. */
type ParagraphRead = (ref: RefOf) => AutomationOperation;

/**
 * A run of text in a text frame.
 *
 * @public
 */
export class TextRange extends ModelObject {
  readonly #shape: ObjectPath;
  #span: Span;
  #font: ShapeFont | undefined;
  #paragraph: ParagraphFormat | undefined;
  #hyperlinks: HyperlinkScopedCollection | undefined;

  /** @internal */
  constructor(
    context: RequestContext,
    path: ObjectPath,
    shape: ObjectPath,
    span: Span,
    nullable = false
  ) {
    super(context, path, { nullable });
    this.#shape = shape;
    this.#span = span;
  }

  /** @internal The reader's selected text, over a pending path the selection read resolves. */
  static selected(owner: ModelObject, name: string, nullable: boolean): TextRange {
    const path = ObjectPath.pending(`presentation.${name}`);
    const range = new TextRange(owner.context, path, path, undefined, nullable);
    range.enqueue(
      'read',
      '',
      () => ({ op: 'readSelection' }),
      (value) => {
        const selected = hydrated(value, 'selection', path.label).selection.text;
        if (!selected) {
          if (nullable) return path.resolveNull();
          fail({ code: 'ItemNotFound', target: path.label });
        }
        range.#span = { start: selected.start, length: selected.length };
        path.resolveTo(selected.shape);
      },
      path,
      false
    );
    return range;
  }

  /**
   * @internal A nullable range over a pending path whose shape and span a later read gives:
   * the text range of a hyperlink that a lookup finds. `settle` resolves it, or marks it null.
   */
  static later(
    context: RequestContext,
    path: ObjectPath
  ): { range: TextRange; settle: (to: { shape: AutomationHandle; span: Span } | null) => void } {
    const range = new TextRange(context, path, path, undefined, true);
    return {
      range,
      settle(to) {
        if (!to) return path.resolveNull();
        range.#span = to.span;
        path.resolveTo(to.shape);
      },
    };
  }

  /** The text. Paragraphs are separated by `\r`, and a line break reads as `\v`. */
  get text(): string {
    return this.loadedProperty('text');
  }
  set text(value: string) {
    const body = text(value, this.at('text'));
    this.command('text', (ref) => ({
      op: 'setText',
      shape: ref(this.#shape),
      ...this.#spanArg(),
      text: body,
    }));
  }
  /** The zero-based start of the range in the text of the frame. */
  get start(): number {
    return this.loadedProperty('start');
  }
  set start(value: number) {
    const start = this.#offset(value, 'start');
    const length = this.#span?.length ?? this.#loadedLength();
    this.#span = { start, length };
    this.setLoadedProperty('start', start);
  }
  /** The length of the range. */
  get length(): number {
    return this.loadedProperty('length');
  }
  set length(value: number) {
    const length = this.#offset(value, 'length');
    this.#span = { start: this.#span?.start ?? 0, length };
    this.setLoadedProperty('length', length);
  }

  /** The character formatting of the range. */
  get font(): ShapeFont {
    return (this.#font ??= new ShapeFont(
      this.context,
      ObjectPath.derived(`${this.path.label}.font`, this.path),
      this.#shape,
      () => this.#span
    ));
  }
  /** The paragraph formatting of the paragraphs the range touches. */
  get paragraphFormat(): ParagraphFormat {
    return (this.#paragraph ??= new ParagraphFormat(
      this.context,
      ObjectPath.derived(`${this.path.label}.paragraphFormat`, this.path),
      this.#shape,
      () => this.#span
    ));
  }

  /** The text links that overlap the range, in text order. */
  get hyperlinks(): HyperlinkScopedCollection {
    return (this.#hyperlinks ??= new HyperlinkScopedCollection(
      this.context,
      ObjectPath.derived(`${this.path.label}.hyperlinks`, this.path),
      this.#shape,
      () => this.#span
    ));
  }

  /** The text frame that holds the range. */
  getParentTextFrame(): TextFrame {
    return new TextFrame(
      this.context,
      ObjectPath.derived(`${this.path.label}.getParentTextFrame()`, this.#shape),
      this.#shape
    );
  }

  /**
   * A range inside this one, from `start` (relative to this range) for `length` characters, or
   * to the end of this range when `length` is absent.
   */
  getSubstring(start: number, length?: number): TextRange {
    const from = this.#offset(start, 'start');
    const count = length === undefined ? undefined : this.#offset(length, 'length');
    const base = this.#span;
    if (base && count !== undefined) {
      if (from + count > base.length)
        failArgument(
          `${this.path.label}.getSubstring`,
          'a start and a length inside the text range'
        );
      return this.#sub({ start: base.start + from, length: count });
    }
    if (base) {
      if (from > base.length)
        failArgument(`${this.path.label}.getSubstring`, 'a start inside the text range');
      return this.#sub({ start: base.start + from, length: base.length - from });
    }
    if (count !== undefined) return this.#sub({ start: from, length: count });
    // To the end of the whole text: the length is known after a read of the text.
    const path = ObjectPath.pending(`${this.path.label}.getSubstring()`);
    const range = new TextRange(this.context, path, path, undefined);
    const shape = this.#shape;
    this.enqueue(
      'read',
      'getSubstring()',
      (ref) => ({ op: 'readText', shape: ref(shape) }),
      (value) => {
        const all = hydrated(value, 'text', path.label).text;
        if (from > all.length) failArgument(path.label, 'a start inside the text');
        range.#span = { start: from, length: all.length - from };
        path.resolveTo(shape.handle());
      },
      path,
      false
    );
    return range;
  }

  /** Select the range in the editor. A server runtime refuses with `NotSupported`. */
  setSelected(): void {
    this.command('setSelected', (ref) => ({
      op: 'selectText',
      shape: ref(this.#shape),
      ...this.#spanArg(),
    }));
  }

  /** Link the range to an address. Answers the hyperlink. */
  setHyperlink(options?: HyperlinkAddOptions): Hyperlink {
    const address = text(options?.address ?? '', this.at('address'));
    const screenTip = options?.screenTip;
    if (screenTip !== undefined) text(screenTip, this.at('screenTip'));
    this.command('setHyperlink', (ref) => ({
      op: 'setHyperlink',
      shape: ref(this.#shape),
      ...this.#spanArg(),
      address,
      ...(screenTip === undefined ? {} : { screenTip }),
    }));
    return Hyperlink.created(this.context, this.#shape, () => this.#span, address, screenTip ?? '');
  }

  #sub(span: AutomationTextSpan): TextRange {
    return new TextRange(
      this.context,
      ObjectPath.derived(`${this.path.label}.getSubstring()`, this.#shape),
      this.#shape,
      span
    );
  }

  #spanArg(): { span?: AutomationTextSpan } {
    return this.#span ? { span: this.#span } : {};
  }

  #offset(value: unknown, name: string): number {
    if (typeof value !== 'number' || !Number.isInteger(value) || value < 0)
      failArgument(`${this.path.label}.${name}`, 'a whole number of 0 or more');
    return value;
  }

  #loadedLength(): number {
    if (!this.hasLoadedProperty('length'))
      fail({ code: 'PropertyNotLoaded', target: `${this.path.label}.length` });
    return this.loadedProperty<number>('length');
  }

  /** @internal */
  protected override onLoad(request: ResolvedLoadOptions): void {
    const names = this.selection(request, ['text', 'start', 'length'], {
      font: () => this.font,
      paragraphFormat: () => this.paragraphFormat,
      hyperlinks: () => this.hyperlinks,
    });
    this.loadScalars(
      names,
      'text',
      (ref) => ({ op: 'readText', shape: ref(this.#shape), ...this.#spanArg() }),
      (v) => ({
        text: v.text,
        start: this.#span?.start ?? 0,
        length: this.#span?.length ?? v.text.length,
      })
    );
  }
}

const FONT_NAMES = [
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
] as const;

/**
 * The character formatting of a text range. A property reads `null` when the range mixes values.
 *
 * @public
 */
export class ShapeFont extends ModelObject {
  readonly #shape: ObjectPath;
  readonly #span: () => Span;
  readonly #cell: CellArg;

  /** @internal */
  constructor(
    context: RequestContext,
    path: ObjectPath,
    shape: ObjectPath,
    span: () => Span,
    cell?: AutomationCellAddress
  ) {
    super(context, path);
    this.#shape = shape;
    this.#span = span;
    this.#cell = cell ? { cell } : {};
  }

  get allCaps(): boolean | null {
    return this.loadedProperty('allCaps');
  }
  set allCaps(value: boolean | null) {
    this.#write('allCaps', { allCaps: flag(value, this.at('allCaps')) });
  }
  get bold(): boolean | null {
    return this.loadedProperty('bold');
  }
  set bold(value: boolean | null) {
    this.#write('bold', { bold: flag(value, this.at('bold')) });
  }
  /** The color as `#RRGGBB`. */
  get color(): string | null {
    return this.loadedProperty('color');
  }
  set color(value: string | null) {
    this.#write('color', { color: text(value, this.at('color')) });
  }
  get doubleStrikethrough(): boolean | null {
    return this.loadedProperty('doubleStrikethrough');
  }
  set doubleStrikethrough(value: boolean | null) {
    this.#write('doubleStrikethrough', {
      doubleStrikethrough: flag(value, this.at('doubleStrikethrough')),
    });
  }
  get italic(): boolean | null {
    return this.loadedProperty('italic');
  }
  set italic(value: boolean | null) {
    this.#write('italic', { italic: flag(value, this.at('italic')) });
  }
  /** The font family name. */
  get name(): string | null {
    return this.loadedProperty('name');
  }
  set name(value: string | null) {
    this.#write('name', { name: text(value, this.at('name')) });
  }
  /** The size, in points. */
  get size(): number | null {
    return this.loadedProperty('size');
  }
  set size(value: number | null) {
    this.#write('size', { size: finite(value, this.at('size')) });
  }
  get smallCaps(): boolean | null {
    return this.loadedProperty('smallCaps');
  }
  set smallCaps(value: boolean | null) {
    this.#write('smallCaps', { smallCaps: flag(value, this.at('smallCaps')) });
  }
  get strikethrough(): boolean | null {
    return this.loadedProperty('strikethrough');
  }
  set strikethrough(value: boolean | null) {
    this.#write('strikethrough', { strikethrough: flag(value, this.at('strikethrough')) });
  }
  get subscript(): boolean | null {
    return this.loadedProperty('subscript');
  }
  set subscript(value: boolean | null) {
    this.#write('subscript', { subscript: flag(value, this.at('subscript')) });
  }
  get superscript(): boolean | null {
    return this.loadedProperty('superscript');
  }
  set superscript(value: boolean | null) {
    this.#write('superscript', { superscript: flag(value, this.at('superscript')) });
  }
  get underline(): ShapeFontUnderlineStyle | `${ShapeFontUnderlineStyle}` | null {
    return this.loadedProperty('underline');
  }
  set underline(value: ShapeFontUnderlineStyle | `${ShapeFontUnderlineStyle}` | null) {
    this.#write('underline', { underline: UNDERLINE.toFile(value, 'underline') });
  }

  #write(name: string, font: AutomationFontWrite): void {
    const shape = this.#shape;
    this.command(name, (ref: RefOf) => {
      const span = this.#span();
      return { op: 'setFont', shape: ref(shape), ...(span ? { span } : {}), ...this.#cell, font };
    });
  }

  /** @internal */
  protected override onLoad(request: ResolvedLoadOptions): void {
    const names = this.selection(request, FONT_NAMES);
    const shape = this.#shape;
    this.loadScalars(
      names,
      'font',
      (ref) => {
        const span = this.#span();
        return { op: 'readFont', shape: ref(shape), ...(span ? { span } : {}), ...this.#cell };
      },
      (v) => ({ ...v.font, underline: UNDERLINE.fromFile(v.font.underline) })
    );
  }
}

/**
 * The paragraph formatting of the paragraphs a text range touches. A property reads `null` when
 * they differ.
 *
 * @public
 */
export class ParagraphFormat extends ModelObject {
  readonly #shape: ObjectPath;
  readonly #span: () => Span;
  readonly #cell: CellArg;
  #bullets: BulletFormat | undefined;

  /** @internal */
  constructor(
    context: RequestContext,
    path: ObjectPath,
    shape: ObjectPath,
    span: () => Span,
    cell?: AutomationCellAddress
  ) {
    super(context, path);
    this.#shape = shape;
    this.#span = span;
    this.#cell = cell ? { cell } : {};
  }

  /** The horizontal alignment (`a:pPr/@algn`). */
  get horizontalAlignment():
    ParagraphHorizontalAlignment | `${ParagraphHorizontalAlignment}` | null {
    return this.loadedProperty('horizontalAlignment');
  }
  set horizontalAlignment(
    value: ParagraphHorizontalAlignment | `${ParagraphHorizontalAlignment}` | null
  ) {
    this.#write('horizontalAlignment', {
      alignment: HORIZONTAL_ALIGNMENT.toFile(value, 'horizontalAlignment'),
    });
  }
  /**
   * The outline level from 1 to 9, as for `TableCell.indentLevel`: one more than `a:pPr/@lvl`,
   * so a paragraph with no `lvl` reads 1.
   */
  get indentLevel(): number {
    return this.loadedProperty('indentLevel');
  }
  set indentLevel(value: number) {
    this.#write('indentLevel', { indentLevel: finite(value, this.at('indentLevel')) });
  }
  /** The bullets of the paragraphs. */
  get bulletFormat(): BulletFormat {
    return (this.#bullets ??= new BulletFormat(
      this.context,
      ObjectPath.derived(`${this.path.label}.bulletFormat`, this.path),
      (name, paragraph) => this.#write(name, paragraph),
      () => this.#plan()
    ));
  }

  #plan(): ParagraphRead {
    const shape = this.#shape;
    return (ref) => {
      const span = this.#span();
      return { op: 'readParagraph', shape: ref(shape), ...(span ? { span } : {}), ...this.#cell };
    };
  }

  #write(name: string, paragraph: AutomationParagraphWrite): void {
    const shape = this.#shape;
    this.command(name, (ref) => {
      const span = this.#span();
      return {
        op: 'setParagraph',
        shape: ref(shape),
        ...(span ? { span } : {}),
        ...this.#cell,
        paragraph,
      };
    });
  }

  /** @internal */
  protected override onLoad(request: ResolvedLoadOptions): void {
    const names = this.selection(request, ['horizontalAlignment', 'indentLevel'], {
      bulletFormat: () => this.bulletFormat,
    });
    this.loadScalars(names, 'paragraph', this.#plan(), (v) => ({
      horizontalAlignment: HORIZONTAL_ALIGNMENT.fromFile(v.paragraph.alignment),
      indentLevel: v.paragraph.indentLevel,
    }));
  }
}

/**
 * The bullets of paragraphs (`a:buNone`, `a:buChar`, `a:buAutoNum`).
 *
 * @public
 */
export class BulletFormat extends ModelObject {
  readonly #write: (name: string, paragraph: AutomationParagraphWrite) => void;
  readonly #plan: () => ParagraphRead;

  /** @internal */
  constructor(
    context: RequestContext,
    path: ObjectPath,
    write: (name: string, paragraph: AutomationParagraphWrite) => void,
    plan: () => ParagraphRead
  ) {
    super(context, path);
    this.#write = write;
    this.#plan = plan;
  }

  /** The numbering scheme of numbered bullets, or `null` when the paragraphs differ. */
  get style(): BulletStyle | `${BulletStyle}` | null {
    return this.loadedProperty('style');
  }
  set style(value: BulletStyle | `${BulletStyle}` | null) {
    this.#write('style', { bulletStyle: BULLET_STYLE.toFile(value, 'style') });
  }
  /** The kind of bullet, or `null` when the paragraphs differ. */
  get type(): BulletType | `${BulletType}` | null {
    return this.loadedProperty('type');
  }
  set type(value: BulletType | `${BulletType}` | null) {
    const bulletType = BULLET_TYPE.toFile(value, 'type') as 'none' | 'char' | 'autonum';
    this.#write('type', { bulletType });
  }
  /** Whether the paragraphs show bullets, or `null` when they differ. */
  get visible(): boolean | null {
    return this.loadedProperty('visible');
  }
  set visible(value: boolean | null) {
    this.#write('visible', { bulletVisible: flag(value, this.at('visible')) });
  }

  /** @internal */
  protected override onLoad(request: ResolvedLoadOptions): void {
    const names = this.selection(request, ['style', 'type', 'visible']);
    this.loadScalars(names, 'paragraph', this.#plan(), (v) => ({
      style:
        v.paragraph.bulletStyle === null ? null : BULLET_STYLE.fromFile(v.paragraph.bulletStyle),
      type: v.paragraph.bulletType === null ? null : BULLET_TYPE.fromFile(v.paragraph.bulletType),
      visible: v.paragraph.bulletVisible,
    }));
  }
}
