/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Office.js `SlideBackground`, `SlideLayoutBackground`, `SlideMasterBackground`,
// `SlideBackgroundFill`, and its solid, gradient, pattern, and picture fills: the background of
// a slide, a layout, or a master (`p:cSld/p:bg`, ECMA-376 Part 1, 19.3.1.1). A read gives the
// background the part shows, also when it comes from the layout or master. A fill write gives
// the part its own `p:bgPr`.

import type {
  AutomationBackgroundRead,
  AutomationBackgroundWrite,
} from '@pptx-editor.dev/core/automation';
import { fail, failArgument, nearest, oneOf } from '../runtime/errors.ts';
import { hydrated } from '../runtime/hydrate.ts';
import type { ResolvedLoadOptions } from '../runtime/load-options.ts';
import { ObjectPath } from '../runtime/object-path.ts';
import type { RequestContext } from '../runtime/request-context.ts';
import {
  BACKGROUND_FILL_TYPE,
  BACKGROUND_GRADIENT_TYPE,
  BACKGROUND_PATTERN,
  type SlideBackgroundFillType,
  type SlideBackgroundGradientFillType,
  type SlideBackgroundPatternFillType,
} from './enums.ts';
import { fraction } from './fill-line.ts';
import { flag, ModelObject, text } from './model-object.ts';

/**
 * Options for `SlideBackgroundFill.setSolidFill`.
 *
 * @public
 */
export interface SlideBackgroundSolidFillOptions {
  /** `#RRGGBB` or a preset color name. The current solid color, or white, when absent. */
  color?: string | undefined;
  /** From 0 (opaque) to 1 (clear). The current transparency, or 0, when absent. */
  transparency?: number | undefined;
}

/**
 * Options for `SlideBackgroundFill.setGradientFill`.
 *
 * @public
 */
export interface SlideBackgroundGradientFillOptions {
  /** The shading. The current shading, or `Linear`, when absent. */
  type?: SlideBackgroundGradientFillType | `${SlideBackgroundGradientFillType}` | undefined;
}

/**
 * Options for `SlideBackgroundFill.setPatternFill`.
 *
 * @public
 */
export interface SlideBackgroundPatternFillOptions {
  /** The color behind the pattern. The current one, or white, when absent. */
  backgroundColor?: string | undefined;
  /** The color of the pattern. The current one, or black, when absent. */
  foregroundColor?: string | undefined;
  /** The pattern. The current one, or `Percent5`, when absent. */
  pattern?: SlideBackgroundPatternFillType | `${SlideBackgroundPatternFillType}` | undefined;
}

/**
 * Options for `SlideBackgroundFill.setPictureOrTextureFill`.
 *
 * @public
 */
export interface SlideBackgroundPictureOrTextureFillOptions {
  /** PNG, JPEG, or GIF bytes in base64. The current picture when absent. */
  imageBase64?: string | undefined;
  /** From 0 (opaque) to 1 (clear). */
  transparency: number;
}

/** The names an options object may carry, or `InvalidArgument` naming the first other one. */
function optionsOf<T extends object>(
  value: T | undefined,
  names: readonly string[],
  target: string
): Partial<Record<string, unknown>> {
  if (value === undefined) return {};
  if (typeof value !== 'object' || value === null || Array.isArray(value))
    failArgument(target, 'an options object');
  for (const key of Object.keys(value))
    if (!names.includes(key))
      failArgument(`${target}.${key}`, `an option name: ${oneOf(names)}`, nearest(key, names));
  return value as Partial<Record<string, unknown>>;
}

/** A gradient shading a write accepts. `ShadeFromTitle` is a read state only. */
function gradientArg(value: unknown, target: string): 'linear' | 'circle' | 'rect' | 'shape' {
  if (value === 'ShadeFromTitle') fail({ code: 'NotImplemented', target });
  return BACKGROUND_GRADIENT_TYPE.toFile(value, target) as 'linear' | 'circle' | 'rect' | 'shape';
}

/** The object that writes and reads one background, by the slide, layout, or master path. */
abstract class BackgroundPart extends ModelObject {
  /** @internal The slide, layout, or master; for a typed fill, its null-object path. */
  protected readonly owner: ObjectPath;

  /** @internal */
  constructor(context: RequestContext, path: ObjectPath, owner: ObjectPath, nullable = false) {
    super(context, path, { nullable });
    this.owner = owner;
  }

  /** @internal Queue one background write. */
  protected writeBackground(name: string, background: AutomationBackgroundWrite): void {
    const owner = this.owner;
    this.command(name, (ref) => ({ op: 'setBackground', owner: ref(owner), background }));
  }

  /** @internal Queue the background read that fills `names`. */
  protected loadBackground(
    names: readonly string[],
    project: (background: AutomationBackgroundRead) => Readonly<Record<string, unknown>>
  ): void {
    const owner = this.owner;
    this.loadScalars(
      names,
      'background',
      (ref) => ({ op: 'readBackground', owner: ref(owner) }),
      (v) => project(v.background)
    );
  }
}

/**
 * A solid background fill.
 *
 * @public
 */
export class SlideBackgroundSolidFill extends BackgroundPart {
  /** The color as `#RRGGBB`. */
  get color(): string {
    return this.loadedProperty('color');
  }
  set color(value: string) {
    this.writeBackground('color', { kind: 'solid', color: text(value, this.at('color')) });
  }
  /** The transparency, from 0 (opaque) to 1 (clear). */
  get transparency(): number {
    return this.loadedProperty('transparency');
  }
  set transparency(value: number) {
    const transparency = fraction(value, this.at('transparency'));
    this.writeBackground('transparency', { kind: 'solid', transparency });
  }

  /** @internal */
  protected override onLoad(request: ResolvedLoadOptions): void {
    this.loadBackground(this.selection(request, ['color', 'transparency']), (b) => ({
      color: b.fill.color ?? '',
      transparency: b.fill.transparency ?? 0,
    }));
  }
}

/**
 * A gradient background fill.
 *
 * @public
 */
export class SlideBackgroundGradientFill extends BackgroundPart {
  /** The shading: `a:lin`, or an `a:path` of `circle` (Radial), `rect`, or `shape` (Path). */
  get type(): SlideBackgroundGradientFillType | `${SlideBackgroundGradientFillType}` {
    return this.loadedProperty('type');
  }
  set type(value: SlideBackgroundGradientFillType | `${SlideBackgroundGradientFillType}`) {
    const gradientType = gradientArg(value, this.at('type'));
    this.writeBackground('type', { kind: 'gradient', gradientType });
  }

  /** @internal */
  protected override onLoad(request: ResolvedLoadOptions): void {
    this.loadBackground(this.selection(request, ['type']), (b) => ({
      type:
        b.fill.gradientType === 'shadeToTitle'
          ? 'ShadeFromTitle'
          : (BACKGROUND_GRADIENT_TYPE.fromFile(b.fill.gradientType) ?? 'Unsupported'),
    }));
  }
}

/**
 * A pattern background fill (`a:pattFill`).
 *
 * @public
 */
export class SlideBackgroundPatternFill extends BackgroundPart {
  /** The color behind the pattern, as `#RRGGBB`. */
  get backgroundColor(): string {
    return this.loadedProperty('backgroundColor');
  }
  set backgroundColor(value: string) {
    const backgroundColor = text(value, this.at('backgroundColor'));
    this.writeBackground('backgroundColor', { kind: 'pattern', backgroundColor });
  }
  /** The color of the pattern, as `#RRGGBB`. */
  get foregroundColor(): string {
    return this.loadedProperty('foregroundColor');
  }
  set foregroundColor(value: string) {
    const foregroundColor = text(value, this.at('foregroundColor'));
    this.writeBackground('foregroundColor', { kind: 'pattern', foregroundColor });
  }
  /** The pattern (`ST_PresetPatternVal`). */
  get pattern(): SlideBackgroundPatternFillType | `${SlideBackgroundPatternFillType}` {
    return this.loadedProperty('pattern');
  }
  set pattern(value: SlideBackgroundPatternFillType | `${SlideBackgroundPatternFillType}`) {
    const pattern = BACKGROUND_PATTERN.toFile(value, this.at('pattern'));
    this.writeBackground('pattern', { kind: 'pattern', pattern });
  }

  /** @internal */
  protected override onLoad(request: ResolvedLoadOptions): void {
    const names = ['backgroundColor', 'foregroundColor', 'pattern'];
    this.loadBackground(this.selection(request, names), (b) => ({
      backgroundColor: b.fill.backgroundColor ?? '',
      foregroundColor: b.fill.foregroundColor ?? '',
      pattern: BACKGROUND_PATTERN.fromFile(b.fill.pattern) ?? 'Unsupported',
    }));
  }
}

/**
 * A picture background fill (`a:blipFill`), stretched to the slide.
 *
 * @public
 */
export class SlideBackgroundPictureOrTextureFill extends BackgroundPart {
  /** The transparency of the picture (`a:alphaModFix`), from 0 (opaque) to 1 (clear). */
  get transparency(): number {
    return this.loadedProperty('transparency');
  }
  set transparency(value: number) {
    const transparency = fraction(value, this.at('transparency'));
    this.writeBackground('transparency', { kind: 'picture', transparency });
  }

  /** Show other picture bytes: PNG, JPEG, or GIF in base64. The transparency stays. */
  setImage(base64EncodedImage: string): void {
    const base64 = text(base64EncodedImage, this.at('setImage'));
    this.writeBackground('setImage', { kind: 'picture', base64 });
  }

  /** @internal */
  protected override onLoad(request: ResolvedLoadOptions): void {
    this.loadBackground(this.selection(request, ['transparency']), (b) => ({
      transparency: b.fill.transparency ?? 0,
    }));
  }
}

/**
 * The fill of a background.
 *
 * @public
 */
export class SlideBackgroundFill extends BackgroundPart {
  /** The kind of fill. `Unsupported` for no fill. */
  get type(): SlideBackgroundFillType | `${SlideBackgroundFillType}` {
    return this.loadedProperty('type');
  }

  /** The gradient fill, or a null object when the fill is of another kind. */
  getGradientFillOrNullObject(): SlideBackgroundGradientFill {
    return this.#typed(
      'getGradientFillOrNullObject',
      'gradient',
      (p) => new SlideBackgroundGradientFill(this.context, p, p, true)
    );
  }
  /** The pattern fill, or a null object when the fill is of another kind. */
  getPatternFillOrNullObject(): SlideBackgroundPatternFill {
    return this.#typed(
      'getPatternFillOrNullObject',
      'pattern',
      (p) => new SlideBackgroundPatternFill(this.context, p, p, true)
    );
  }
  /** The picture fill, or a null object when the fill is of another kind. */
  getPictureOrTextureFillOrNullObject(): SlideBackgroundPictureOrTextureFill {
    return this.#typed(
      'getPictureOrTextureFillOrNullObject',
      'picture',
      (p) => new SlideBackgroundPictureOrTextureFill(this.context, p, p, true)
    );
  }
  /** The solid fill, or a null object when the fill is of another kind. */
  getSolidFillOrNullObject(): SlideBackgroundSolidFill {
    return this.#typed(
      'getSolidFillOrNullObject',
      'solid',
      (p) => new SlideBackgroundSolidFill(this.context, p, p, true)
    );
  }

  /** Fill with a gradient. The stops of a current gradient stay. */
  setGradientFill(options?: SlideBackgroundGradientFillOptions): void {
    const target = this.at('setGradientFill');
    const o = optionsOf(options, ['type'], target);
    this.writeBackground('setGradientFill', {
      kind: 'gradient',
      ...(o.type === undefined ? {} : { gradientType: gradientArg(o.type, `${target}.type`) }),
    });
  }

  /** Fill with a pattern (`a:pattFill`). */
  setPatternFill(options?: SlideBackgroundPatternFillOptions): void {
    const target = this.at('setPatternFill');
    const o = optionsOf(options, ['backgroundColor', 'foregroundColor', 'pattern'], target);
    this.writeBackground('setPatternFill', {
      kind: 'pattern',
      ...(o.pattern === undefined
        ? {}
        : { pattern: BACKGROUND_PATTERN.toFile(o.pattern, `${target}.pattern`) }),
      ...(o.foregroundColor === undefined
        ? {}
        : { foregroundColor: text(o.foregroundColor, `${target}.foregroundColor`) }),
      ...(o.backgroundColor === undefined
        ? {}
        : { backgroundColor: text(o.backgroundColor, `${target}.backgroundColor`) }),
    });
  }

  /** Fill with a stretched picture. Without `imageBase64`, the current picture stays. */
  setPictureOrTextureFill(options?: SlideBackgroundPictureOrTextureFillOptions): void {
    const target = this.at('setPictureOrTextureFill');
    const o = optionsOf(options, ['imageBase64', 'transparency'], target);
    this.writeBackground('setPictureOrTextureFill', {
      kind: 'picture',
      ...(o.imageBase64 === undefined
        ? {}
        : { base64: text(o.imageBase64, `${target}.imageBase64`) }),
      ...(o.transparency === undefined
        ? {}
        : { transparency: fraction(o.transparency, `${target}.transparency`) }),
    });
  }

  /** Fill with one color. */
  setSolidFill(options?: SlideBackgroundSolidFillOptions): void {
    const target = this.at('setSolidFill');
    const o = optionsOf(options, ['color', 'transparency'], target);
    this.writeBackground('setSolidFill', {
      kind: 'solid',
      ...(o.color === undefined ? {} : { color: text(o.color, `${target}.color`) }),
      ...(o.transparency === undefined
        ? {}
        : { transparency: fraction(o.transparency, `${target}.transparency`) }),
    });
  }

  /**
   * A typed fill over a pending path: the read queued now resolves it to the owner when the fill
   * is of `kind`, and to null otherwise. A call on it waits for that read.
   */
  #typed<T>(name: string, kind: string, make: (path: ObjectPath) => T): T {
    const path = ObjectPath.pending(`${this.path.label}.${name}()`);
    const object = make(path);
    const owner = this.owner;
    const label = path.label;
    this.enqueue(
      'read',
      `${name}()`,
      (ref) => ({ op: 'readBackground', owner: ref(owner) }),
      (value) => {
        const read = value.kind === 'null' ? null : hydrated(value, 'background', label).background;
        if (read?.fill.type === kind && owner.isAddressable) path.resolveTo(owner.handle());
        else path.resolveNull();
      },
      path,
      false
    );
    return object;
  }

  /** @internal */
  protected override onLoad(request: ResolvedLoadOptions): void {
    this.loadBackground(this.selection(request, ['type']), (b) => ({
      type: BACKGROUND_FILL_TYPE.fromFile(b.fill.type) ?? 'Unsupported',
    }));
  }
}

/** The fill object of a background object. */
function fillOf(context: RequestContext, path: ObjectPath, owner: ObjectPath): SlideBackgroundFill {
  return new SlideBackgroundFill(context, ObjectPath.derived(`${path.label}.fill`, path), owner);
}

/**
 * The background of a slide.
 *
 * @public
 */
export class SlideBackground extends BackgroundPart {
  #fill: SlideBackgroundFill | undefined;

  /** Whether the shapes of the layout and master are hidden (`p:sld/@showMasterSp="0"`). */
  get areBackgroundGraphicsHidden(): boolean {
    return this.loadedProperty('areBackgroundGraphicsHidden');
  }
  set areBackgroundGraphicsHidden(value: boolean) {
    const hidden = flag(value, this.at('areBackgroundGraphicsHidden'));
    this.writeBackground('areBackgroundGraphicsHidden', { kind: 'graphicsHidden', hidden });
  }
  /**
   * Whether the slide shows the background of its layout or master (it has no own `p:bg`).
   * Setting `false` copies the shown background into the slide.
   */
  get isMasterBackgroundFollowed(): boolean {
    return this.loadedProperty('isMasterBackgroundFollowed');
  }
  set isMasterBackgroundFollowed(value: boolean) {
    const follow = flag(value, this.at('isMasterBackgroundFollowed'));
    this.writeBackground('isMasterBackgroundFollowed', { kind: follow ? 'follow' : 'own' });
  }
  /** The fill the slide shows. */
  get fill(): SlideBackgroundFill {
    return (this.#fill ??= fillOf(this.context, this.path, this.owner));
  }

  /** Remove the own background: the slide shows the background of its layout or master. */
  reset(): void {
    this.writeBackground('reset', { kind: 'follow' });
  }

  /** @internal */
  protected override onLoad(request: ResolvedLoadOptions): void {
    const names = ['areBackgroundGraphicsHidden', 'isMasterBackgroundFollowed'];
    this.loadBackground(this.selection(request, names, { fill: () => this.fill }), (b) => ({
      areBackgroundGraphicsHidden: b.graphicsHidden,
      isMasterBackgroundFollowed: b.followsMaster,
    }));
  }
}

/**
 * The background of a slide layout.
 *
 * @public
 */
export class SlideLayoutBackground extends BackgroundPart {
  #fill: SlideBackgroundFill | undefined;

  /** Whether the shapes of the master are hidden (`p:sldLayout/@showMasterSp="0"`). */
  get areBackgroundGraphicsHidden(): boolean {
    return this.loadedProperty('areBackgroundGraphicsHidden');
  }
  set areBackgroundGraphicsHidden(value: boolean) {
    const hidden = flag(value, this.at('areBackgroundGraphicsHidden'));
    this.writeBackground('areBackgroundGraphicsHidden', { kind: 'graphicsHidden', hidden });
  }
  /**
   * Whether the layout shows the background of its master (it has no own `p:bg`). Setting
   * `false` copies the master background into the layout.
   */
  get isMasterBackgroundFollowed(): boolean {
    return this.loadedProperty('isMasterBackgroundFollowed');
  }
  set isMasterBackgroundFollowed(value: boolean) {
    const follow = flag(value, this.at('isMasterBackgroundFollowed'));
    this.writeBackground('isMasterBackgroundFollowed', { kind: follow ? 'follow' : 'own' });
  }
  /** The fill the layout shows. */
  get fill(): SlideBackgroundFill {
    return (this.#fill ??= fillOf(this.context, this.path, this.owner));
  }

  /** Remove the own background: the layout shows the background of its master. */
  reset(): void {
    this.writeBackground('reset', { kind: 'follow' });
  }

  /** @internal */
  protected override onLoad(request: ResolvedLoadOptions): void {
    const names = ['areBackgroundGraphicsHidden', 'isMasterBackgroundFollowed'];
    this.loadBackground(this.selection(request, names, { fill: () => this.fill }), (b) => ({
      areBackgroundGraphicsHidden: b.graphicsHidden,
      isMasterBackgroundFollowed: b.followsMaster,
    }));
  }
}

/**
 * The background of a slide master.
 *
 * @public
 */
export class SlideMasterBackground extends BackgroundPart {
  #fill: SlideBackgroundFill | undefined;

  /** The fill of the master background. */
  get fill(): SlideBackgroundFill {
    return (this.#fill ??= fillOf(this.context, this.path, this.owner));
  }

  /** @internal */
  protected override onLoad(request: ResolvedLoadOptions): void {
    this.selection(request, [], { fill: () => this.fill });
  }
}
