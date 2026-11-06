/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Office.js `ShapeFill` and `ShapeLineFormat`: the fill and the outline of `p:spPr`
// (EG_FillProperties, ECMA-376 Part 1, 20.1.8; CT_LineProperties, 20.1.2.2.24). Colors are
// `#RRGGBB` strings or preset color names; transparency is 0 (opaque) to 1 (clear).

import type {
  AutomationCellAddress,
  AutomationFillWrite,
  AutomationLineWrite,
} from '@pptx-editor.dev/core/automation';
import { failArgument } from '../runtime/errors.ts';
import type { ResolvedLoadOptions } from '../runtime/load-options.ts';
import type { ObjectPath } from '../runtime/object-path.ts';
import type { RequestContext } from '../runtime/request-context.ts';
import {
  DASH_STYLE,
  FILL_TYPE,
  LINE_STYLE,
  type ShapeFillType,
  type ShapeLineDashStyle,
  type ShapeLineStyle,
} from './enums.ts';
import { finite, flag, ModelObject, text } from './model-object.ts';

/** What a fill belongs to: a shape, or one cell of a table shape (`a:tcPr`). */
export type FillTarget =
  | { readonly kind: 'shape'; readonly shape: ObjectPath }
  | { readonly kind: 'cell'; readonly shape: ObjectPath; readonly cell: AutomationCellAddress };

/** A fraction from 0 to 1, or `InvalidArgument`. */
export function fraction(value: unknown, target: string): number {
  const n = finite(value, target);
  if (n < 0 || n > 1) failArgument(target, 'a fraction from 0 (opaque) to 1 (clear)');
  return n;
}

/**
 * The fill of a shape.
 *
 * @public
 */
export class ShapeFill extends ModelObject {
  readonly #target: FillTarget;

  /** @internal */
  constructor(context: RequestContext, path: ObjectPath, target: FillTarget) {
    super(context, path);
    this.#target = target;
  }

  /** The solid fill color as `#RRGGBB`, or an empty string for any other fill. */
  get foregroundColor(): string {
    return this.loadedProperty('foregroundColor');
  }
  set foregroundColor(value: string) {
    this.#write('foregroundColor', {
      kind: 'solid',
      color: text(value, this.at('foregroundColor')),
    });
  }
  /** The transparency of a solid fill, from 0 (opaque) to 1 (clear). */
  get transparency(): number {
    return this.loadedProperty('transparency');
  }
  set transparency(value: number) {
    this.#write('transparency', {
      kind: 'transparency',
      transparency: fraction(value, 'transparency'),
    });
  }
  /** The kind of fill. */
  get type(): ShapeFillType | `${ShapeFillType}` {
    return this.loadedProperty('type');
  }

  /** Remove the fill (`a:noFill`). */
  clear(): void {
    this.#write('clear', { kind: 'clear' });
  }

  /** Fill with a solid color, `#RRGGBB` or a preset color name. */
  setSolidColor(color: string): void {
    this.#write('setSolidColor', { kind: 'solid', color: text(color, this.at('color')) });
  }

  /**
   * Fill with a picture, stretched (`a:blipFill`): PNG, JPEG, or GIF bytes in base64. A table
   * cell refuses a picture fill with `InvalidArgument`.
   */
  setImage(base64EncodedImage: string): void {
    const base64 = text(base64EncodedImage, `${this.path.label}.setImage`);
    this.#write('setImage', { kind: 'image', base64 });
  }

  #write(name: string, fill: AutomationFillWrite): void {
    const shape = this.#target.shape;
    const cell = this.#cell();
    this.command(name, (ref) => ({ op: 'setFill', shape: ref(shape), fill, ...cell }));
  }

  #cell(): { readonly cell?: AutomationCellAddress } {
    return this.#target.kind === 'cell' ? { cell: this.#target.cell } : {};
  }

  /** @internal */
  protected override onLoad(request: ResolvedLoadOptions): void {
    const names = this.selection(request, ['foregroundColor', 'transparency', 'type']);
    const shape = this.#target.shape;
    this.loadScalars(
      names,
      'fill',
      (ref) => ({ op: 'readFill', shape: ref(shape), ...this.#cell() }),
      (v) => ({
        foregroundColor: v.fill.foregroundColor ?? '',
        transparency: v.fill.transparency ?? 0,
        // A group fill (`a:grpFill`) takes the fill of the group; it reads as `NoFill` here.
        type: FILL_TYPE.fromFile(v.fill.type) ?? 'NoFill',
      })
    );
  }
}

const LINE_NAMES = ['color', 'dashStyle', 'style', 'transparency', 'visible', 'weight'] as const;

/**
 * The outline of a shape, or a line.
 *
 * @public
 */
export class ShapeLineFormat extends ModelObject {
  readonly #shape: ObjectPath;

  /** @internal */
  constructor(context: RequestContext, path: ObjectPath, shape: ObjectPath) {
    super(context, path);
    this.#shape = shape;
  }

  /** The line color as `#RRGGBB`, or an empty string when it has none. */
  get color(): string {
    return this.loadedProperty('color');
  }
  set color(value: string) {
    this.#write('color', { color: text(value, this.at('color')) });
  }
  /** The dash style. */
  get dashStyle(): ShapeLineDashStyle | `${ShapeLineDashStyle}` {
    return this.loadedProperty('dashStyle');
  }
  set dashStyle(value: ShapeLineDashStyle | `${ShapeLineDashStyle}`) {
    this.#write('dashStyle', { dashStyle: DASH_STYLE.toFile(value, 'dashStyle') });
  }
  /** The compound line style. */
  get style(): ShapeLineStyle | `${ShapeLineStyle}` {
    return this.loadedProperty('style');
  }
  set style(value: ShapeLineStyle | `${ShapeLineStyle}`) {
    this.#write('style', { style: LINE_STYLE.toFile(value, 'style') });
  }
  /** The transparency, from 0 (opaque) to 1 (clear). */
  get transparency(): number {
    return this.loadedProperty('transparency');
  }
  set transparency(value: number) {
    this.#write('transparency', { transparency: fraction(value, 'transparency') });
  }
  /** Whether the line shows. */
  get visible(): boolean {
    return this.loadedProperty('visible');
  }
  set visible(value: boolean) {
    this.#write('visible', { visible: flag(value, this.at('visible')) });
  }
  /** The line weight, in points. */
  get weight(): number {
    return this.loadedProperty('weight');
  }
  set weight(value: number) {
    this.#write('weight', { weight: finite(value, this.at('weight')) });
  }

  #write(name: string, line: AutomationLineWrite): void {
    const shape = this.#shape;
    this.command(name, (ref) => ({ op: 'setLine', shape: ref(shape), line }));
  }

  /** @internal */
  protected override onLoad(request: ResolvedLoadOptions): void {
    const names = this.selection(request, LINE_NAMES);
    const shape = this.#shape;
    this.loadScalars(
      names,
      'line',
      (ref) => ({ op: 'readLine', shape: ref(shape) }),
      (v) => ({
        color: v.line.color ?? '',
        dashStyle: DASH_STYLE.fromFile(v.line.dashStyle) ?? 'Solid',
        style: LINE_STYLE.fromFile(v.line.style) ?? 'Single',
        transparency: v.line.transparency ?? 0,
        visible: v.line.visible,
        weight: v.line.weight ?? 0,
      })
    );
  }
}
