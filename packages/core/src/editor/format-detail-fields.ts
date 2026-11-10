import { colorPickOf } from './color-fields.ts';
import { renderedColorOf } from './chrome-forms.ts';
import type { GradientDetail, ThemeSwatch } from '../contracts/index.ts';
import type { TranslationKey } from '@pptx-editor.dev/i18n';
import type { ShapeFormatDetail, ShapeFormatPatch, ColorChoice } from '../contracts/index.ts';

/** One numeric property of an existing fill or effect. @internal */
export interface FormatNumberProperty {
  readonly id: string;
  readonly labelKey: TranslationKey;
  readonly value: number;
  readonly unit: '%' | '°' | 'pt';
  readonly min?: number;
  readonly max?: number;
  patch(value: number): ShapeFormatPatch;
}
/** A color property retains theme choices. @internal */
export interface FormatColorProperty {
  readonly id: string;
  readonly labelKey: TranslationKey;
  readonly color?: ColorChoice;
  patch(color: ColorChoice): ShapeFormatPatch;
}
/** Numeric fields for partial edits. Values use the units displayed beside the fields. @internal */
export function formatNumberProperties(
  detail: ShapeFormatDetail,
  section: 'fill' | 'effects'
): readonly FormatNumberProperty[] {
  const fields: FormatNumberProperty[] = [];
  if (section === 'fill' && detail.gradient) {
    const g = detail.gradient;
    if (g.kind === 'linear')
      fields.push({
        id: 'gradient.angle',
        labelKey: 'gradient.angle',
        value: g.angle,
        unit: '°',
        min: 0,
        max: 360,
        patch: (angle) => ({ kind: 'gradient', angle }),
      });
    for (const [index, stop] of g.stops.entries()) {
      fields.push({
        id: `gradient-stop.${index}.position`,
        labelKey: 'panel.format.position',
        value: stop.position * 100,
        min: 0,
        max: 100,
        unit: '%',
        patch: (v) => ({ kind: 'gradient', stop: { id: stop.id, position: v / 100 } }),
      });
      fields.push({
        id: `gradient-stop.${index}.opacity`,
        labelKey: 'inspector.opacity',
        value: stop.opacity * 100,
        min: 0,
        max: 100,
        unit: '%',
        patch: (v) => ({ kind: 'gradient', stop: { id: stop.id, opacity: v / 100 } }),
      });
    }
  }
  if (section === 'effects') {
    for (const kind of ['shadow', 'reflection'] as const) {
      const effect = detail[kind];
      if (!effect) continue;
      for (const property of ['blur', 'distance'] as const)
        fields.push({
          id: `${kind}.${property}`,
          labelKey: `inspector.${property}`,
          value: effect[property],
          unit: 'pt',
          min: 0,
          max: 150,
          patch: (v) => ({ kind, [property]: v }),
        });
    }
    if (detail.shadow) {
      fields.push({
        id: 'shadow.angle',
        labelKey: 'gradient.angle',
        value: detail.shadow.angle,
        unit: '°',
        min: 0,
        max: 360,
        patch: (angle) => ({ kind: 'shadow', angle }),
      });
      fields.push({
        id: 'shadow.opacity',
        labelKey: 'inspector.opacity',
        value: detail.shadow.opacity * 100,
        unit: '%',
        min: 0,
        max: 100,
        patch: (v) => ({ kind: 'shadow', opacity: v / 100 }),
      });
    }
    if (detail.reflection)
      for (const property of [
        'startOpacity',
        'endOpacity',
        'startPosition',
        'endPosition',
      ] as const)
        fields.push({
          id: `reflection.${property}`,
          labelKey: `inspector.${property}`,
          value: detail.reflection[property] * 100,
          unit: '%',
          min: 0,
          max: 100,
          patch: (v) => ({ kind: 'reflection', [property]: v / 100 }),
        });
  }
  return fields;
}
/** Colors for stops and shadows. @internal */
export function formatColorProperties(
  detail: ShapeFormatDetail,
  section: 'fill' | 'effects'
): readonly FormatColorProperty[] {
  if (section === 'fill')
    return (
      detail.gradient?.stops.map((stop, index) => ({
        id: `gradient-stop.${index}.color`,
        labelKey: 'gradient.stopColor',
        color: stop.color,
        patch: (color: ColorChoice) => ({ kind: 'gradient', stop: { id: stop.id, color } }),
      })) ?? []
    );
  return detail.shadow
    ? [
        {
          id: 'shadow.color',
          labelKey: 'inspector.color',
          color: detail.shadow.color,
          patch: (color) => ({ kind: 'shadow', color }),
        },
      ]
    : [];
}

/** One source-bound group of properties. @internal */
export interface FormatPropertyGroup {
  readonly id: string;
  readonly labelKey: TranslationKey;
  readonly index?: number;
  readonly numbers: readonly FormatNumberProperty[];
  readonly colors: readonly FormatColorProperty[];
  readonly remove?: ShapeFormatPatch;
}
/** Group each stop's color and numeric properties under its stable source identity. @internal */
export function formatPropertyGroups(
  detail: ShapeFormatDetail,
  section: 'fill' | 'effects'
): readonly FormatPropertyGroup[] {
  const numbers = formatNumberProperties(detail, section);
  const colors = formatColorProperties(detail, section);
  if (section === 'fill')
    return detail.gradient
      ? [
          ...(detail.gradient.kind === 'linear'
            ? [
                {
                  id: 'gradient',
                  labelKey: 'inspector.gradient' as const,
                  numbers: numbers.filter((f) => f.id === 'gradient.angle'),
                  colors: [],
                },
              ]
            : []),
          ...detail.gradient.stops.map((stop, index) => ({
            id: stop.id,
            labelKey: 'inspector.gradientStop' as const,
            index: index + 1,
            numbers: numbers.filter((f) => f.id.startsWith(`gradient-stop.${index}.`)),
            colors: colors.filter((f) => f.id.startsWith(`gradient-stop.${index}.`)),
            ...(detail.gradient!.stops.length > 2
              ? { remove: { kind: 'gradient' as const, removeStop: stop.id } }
              : {}),
          })),
        ]
      : [];
  return (['shadow', 'reflection'] as const)
    .filter((kind) => detail[kind])
    .map((kind) => ({
      id: kind,
      labelKey: `toolbar.${kind}`,
      numbers: numbers.filter((f) => f.id.startsWith(`${kind}.`)),
      colors: colors.filter((f) => f.id.startsWith(`${kind}.`)),
    }));
}

/** A preview color without replacing an unsupported source color. @internal */
export function gradientPreviewStops(
  gradient: GradientDetail,
  rows: readonly (readonly ThemeSwatch[])[]
): readonly { id: string; position: number; color: string | null; opacity: number }[] {
  return gradient.stops.map((stop) => ({
    id: stop.id,
    position: stop.position,
    opacity: stop.opacity,
    color: stop.color ? renderedColorOf(colorPickOf(stop.color), rows, null, null) : null,
  }));
}
/** A linear ramp previews stop colors and alpha independently of the shape's fill direction. @internal */
export function gradientPreviewCss(
  stops: ReturnType<typeof gradientPreviewStops>
): string | undefined {
  if (stops.some((stop) => !stop.color || !/^#[0-9a-f]{6}$/i.test(stop.color))) return undefined;
  return `linear-gradient(90deg, ${[...stops]
    .sort((a, b) => a.position - b.position)
    .map((s) => {
      const color = s.color!;
      return `rgba(${parseInt(color.slice(1, 3), 16)},${parseInt(color.slice(3, 5), 16)},${parseInt(color.slice(5, 7), 16)},${s.opacity}) ${s.position * 100}%`;
    })
    .join(', ')})`;
}
