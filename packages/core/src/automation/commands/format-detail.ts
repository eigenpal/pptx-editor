// Partial DrawingML edits preserve every property that the patch does not name.
import {
  attr,
  child,
  elements,
  intAttr,
  NS,
  specOf,
  xmlPartOf,
  type Element,
  type NodeSpec,
} from '../../store/index.ts';
import { inheritedPlaceholders, parseTheme } from '../../cascade/index.ts';
import type {
  ColorChoice,
  GradientDetail,
  ShapeFormatDetail,
  ShapeFormatPatch,
} from '../../contracts/neutral.ts';
import { CommandRefusal, type OpBatch } from './op-batch.ts';
import { effectListOf, inheritedEffects, themePartFor, withRefColor } from './effect-commands.ts';
import { colorXml, isColorChoice } from './style-commands.ts';
import { specFromXml } from './shape-commands.ts';

type FormatNode = Element | NodeSpec;
// A batch can override an inherited fill, then edit a stop returned before that override.
// These aliases live only for that batch. A later sync must read the new source identities.
const inheritedStopAliases = new WeakMap<OpBatch, Map<string, string>>();
const stopAliasKey = (part: string, nodeId: string, stopId: string): string =>
  JSON.stringify([part, nodeId, stopId]);
const children = (n: FormatNode | undefined): FormatNode[] =>
  (n?.children ?? []).filter((c): c is Element | NodeSpec => 'local' in c && c.ns === NS.a);
const named = (n: FormatNode | undefined, local: string): FormatNode | undefined =>
  children(n).find((c) => c.local === local);
const value = (n: FormatNode | undefined, key: string): string | undefined =>
  n?.attrs?.find((a) => a.ns === '' && a.local === key)?.value;
const number = (n: FormatNode | undefined, key: string, fallback = 0): number => {
  const s = value(n, key);
  const v = s === undefined ? fallback : Number(s);
  return Number.isFinite(v) ? v : fallback;
};
const rotates = (n: FormatNode): boolean =>
  !['0', 'false'].includes(value(n, 'rotWithShape') ?? '1');
const COLORS = new Set(['srgbClr', 'schemeClr', 'prstClr', 'scrgbClr', 'sysClr', 'hslClr']);
const colorOf = (n: FormatNode): FormatNode | undefined =>
  children(n).find((c) => COLORS.has(c.local));

function colorChoice(n: FormatNode): ColorChoice | undefined {
  const c = colorOf(n);
  if (!c) return undefined;
  const val = value(c, 'val');
  const transforms = children(c).filter(
    (v) => !['alpha', 'alphaMod', 'alphaOff'].includes(v.local)
  );
  if (transforms.length) return undefined;
  const candidate =
    c.local === 'srgbClr' ? `#${val}` : c.local === 'schemeClr' ? { theme: val } : undefined;
  return isColorChoice(candidate) ? candidate : undefined;
}

function opacity(n: FormatNode): number {
  let alpha = 1;
  for (const c of children(colorOf(n))) {
    const v = number(c, 'val') / 100_000;
    if (c.local === 'alpha') alpha = v;
    if (c.local === 'alphaMod') alpha *= v;
    if (c.local === 'alphaOff') alpha += v;
  }
  return Math.max(0, Math.min(1, alpha));
}

function shape(batch: OpBatch, part: string, id: string): Element {
  const el = batch.element(part, id);
  if (!['sp', 'pic', 'cxnSp'].includes(el.kind ?? ''))
    throw new CommandRefusal('disabled.unsupportedElement');
  if (!child(el, 'p:spPr')) throw new CommandRefusal('errors.malformed');
  return el;
}

const FILLS = new Set(['noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill']);
interface FillSource {
  raw: Element;
  spec: NodeSpec;
  own: boolean;
}
function fillSource(batch: OpBatch, part: string, el: Element): FillSource | undefined {
  const chain = [el, ...inheritedPlaceholders(batch.pkg, part, el)];
  for (const e of chain) {
    const raw = elements(child(e, 'p:spPr')).find((c) => c.ns === NS.a && FILLS.has(c.local));
    if (raw) return { raw, spec: specOf(raw), own: e === el };
  }
  const style = chain.map((e) => child(e, 'p:style')).find(Boolean);
  const ref = child(style, 'a:fillRef');
  const idx = intAttr(ref, 'idx') ?? 0;
  const themePart = themePartFor(batch, part);
  const theme = themePart ? parseTheme(xmlPartOf(batch.pkg, themePart)) : undefined;
  const raw = idx >= 1001 ? theme?.backgroundFillStyles[idx - 1001] : theme?.fillStyles[idx - 1];
  if (!raw) return undefined;
  const color = elements(ref)[0];
  return { raw, spec: withRefColor(specOf(raw), color ? specOf(color) : undefined), own: false };
}

function effectsSource(batch: OpBatch, part: string, el: Element): FormatNode[] {
  const props = child(el, 'p:spPr');
  if (child(props, 'a:effectDag')) throw new CommandRefusal('errors.unsupported', 'effectDag');
  const own = child(props, 'a:effectLst');
  return own ? elements(own) : inheritedEffects(batch, part, el);
}

/** Read source values without flattening unsupported DrawingML properties. @internal */
export function readShapeFormatProperties(
  batch: OpBatch,
  part: string,
  id: string
): Omit<ShapeFormatDetail, 'target'> {
  const el = shape(batch, part, id);
  const source = fillSource(batch, part, el);
  const fill = source?.spec;
  let gradient: GradientDetail | null = null;
  if (source && fill?.local === 'gradFill') {
    const lin = named(fill, 'lin');
    const path = named(fill, 'path');
    const rawStops = elements(child(source.raw, 'a:gsLst')).filter(
      (c) => c.ns === NS.a && c.local === 'gs'
    );
    gradient = {
      kind: lin ? 'linear' : value(path, 'path') === 'circle' ? 'radial' : 'other',
      angle: number(lin, 'ang') / 60_000,
      rotWithShape: rotates(fill),
      stops: children(named(fill, 'gsLst'))
        .filter((c) => c.local === 'gs')
        .map((s, i) => ({
          id: rawStops[i]!.id,
          position: number(s, 'pos') / 100_000,
          ...(colorChoice(s) === undefined ? {} : { color: colorChoice(s)! }),
          opacity: opacity(s),
        })),
    };
  }
  const solid =
    source && fill?.local === 'solidFill' && colorOf(fill)
      ? {
          ...(colorChoice(fill) === undefined ? {} : { color: colorChoice(fill)! }),
          opacity: opacity(fill),
        }
      : null;
  let fx: FormatNode[];
  let effectsReason: string | undefined;
  try {
    fx = effectsSource(batch, part, el);
  } catch (error) {
    if (!(error instanceof CommandRefusal)) throw error;
    fx = [];
    effectsReason = error.reason;
  }
  const shadows = fx.filter(
    (e) => e.ns === NS.a && ['innerShdw', 'outerShdw', 'prstShdw'].includes(e.local)
  );
  const s = shadows.length === 1 && shadows[0]!.local !== 'prstShdw' ? shadows[0] : undefined;
  const r = fx.find((e) => e.ns === NS.a && e.local === 'reflection');
  return {
    solid,
    gradient,
    ...(effectsReason ? { effectsReason } : {}),
    shadow: s
      ? {
          kind: s.local === 'innerShdw' ? 'inner' : 'outer',
          blur: number(s, 'blurRad') / 12_700,
          distance: number(s, 'dist') / 12_700,
          angle: number(s, 'dir') / 60_000,
          opacity: opacity(s),
          ...(colorChoice(s) === undefined ? {} : { color: colorChoice(s)! }),
          rotWithShape: s.local === 'outerShdw' && rotates(s),
        }
      : null,
    reflection: r
      ? {
          blur: number(r, 'blurRad') / 12_700,
          distance: number(r, 'dist') / 12_700,
          startOpacity: number(r, 'stA', 100_000) / 100_000,
          endOpacity: number(r, 'endA') / 100_000,
          startPosition: number(r, 'stPos') / 100_000,
          endPosition: number(r, 'endPos', 100_000) / 100_000,
          rotWithShape: rotates(r),
        }
      : null,
  };
}

function set(batch: OpBatch, part: string, el: Element, key: string, v: number | boolean): void {
  const encoded = typeof v === 'boolean' ? (v ? '1' : '0') : String(Math.round(v));
  if (attr(el, key) !== encoded)
    batch.push({
      op: 'setAttribute',
      part,
      node: el.id,
      name: { ns: '', local: key },
      value: encoded,
    });
}

function checked(v: unknown, max: number): number {
  if (typeof v !== 'number' || !Number.isFinite(v) || v < 0 || v > max)
    throw new CommandRefusal('errors.unsupported', 'format value');
  return v;
}

function patchColor(
  batch: OpBatch,
  part: string,
  el: Element,
  color: ColorChoice | undefined,
  alpha: number | undefined
): void {
  const old = elements(el).find((c) => c.ns === NS.a && COLORS.has(c.local));
  if (!old) throw new CommandRefusal('errors.unsupported', 'color');
  let current = old;
  if (color !== undefined) {
    const spec = specFromXml(colorXml(color));
    // Keep the current opacity unless the caller changes it explicitly.
    const transforms = new Set([
      'tint',
      'shade',
      'comp',
      'inv',
      'gray',
      'hue',
      'hueOff',
      'hueMod',
      'sat',
      'satOff',
      'satMod',
      'lum',
      'lumOff',
      'lumMod',
      'red',
      'redOff',
      'redMod',
      'green',
      'greenOff',
      'greenMod',
      'blue',
      'blueOff',
      'blueMod',
      'gamma',
      'invGamma',
    ]);
    const keep = elements(old)
      .filter((c) => c.ns !== NS.a || !transforms.has(c.local))
      .map(specOf);
    const colorAttributes = new Set(['val', 'lastClr', 'r', 'g', 'b', 'hue', 'sat', 'lum']);
    const attrs = old.attrs.filter((a) => a.ns !== '' || !colorAttributes.has(a.local));
    batch.push({ op: 'removeNode', part, node: old.id });
    const id = batch.insert({
      op: 'insertNode',
      part,
      parent: el.id,
      index: el.children.indexOf(old),
      spec: {
        ...spec,
        attrs: [...(spec.attrs ?? []), ...attrs],
        children: [...(spec.children ?? []), ...keep],
      },
    });
    current = batch.element(part, id);
  }
  if (alpha !== undefined) {
    checked(alpha, 1);
    const transforms = elements(current).filter(
      (c) => c.ns === NS.a && ['alpha', 'alphaMod', 'alphaOff'].includes(c.local)
    );
    // A single alpha is updated in place, so concurrent numeric edits share its identity.
    if (transforms.length === 1 && transforms[0]!.local === 'alpha')
      set(batch, part, transforms[0]!, 'val', alpha * 100_000);
    else {
      for (const c of transforms) batch.push({ op: 'removeNode', part, node: c.id });
      batch.insert({
        op: 'insertNode',
        part,
        parent: current.id,
        index: batch.element(part, current.id).children.length,
        spec: {
          ns: NS.a,
          local: 'alpha',
          prefix: 'a',
          attrs: [{ ns: '', local: 'val', value: String(Math.round(alpha * 100_000)) }],
        },
      });
    }
  }
}

/** Copy an inherited fill into the shape's own properties, after its geometry. */
function ownFill(batch: OpBatch, part: string, el: Element, spec: NodeSpec): Element {
  const props = child(el, 'p:spPr')!;
  let index = 0;
  props.children.forEach((c, i) => {
    if (c.type === 'element' && ['xfrm', 'prstGeom', 'custGeom'].includes(c.local)) index = i + 1;
  });
  return batch.element(
    part,
    batch.insert({ op: 'insertNode', part, parent: props.id, index, spec })
  );
}

/** Patch an existing solid fill, gradient, shadow, or reflection through store operations. @internal */
export function patchShapeFormatProperties(
  batch: OpBatch,
  part: string,
  id: string,
  patch: ShapeFormatPatch
): void {
  const el = shape(batch, part, id);
  if (!patch || !['solid', 'gradient', 'shadow', 'reflection'].includes(patch.kind))
    throw new CommandRefusal('errors.unsupported', 'format patch');
  if (patch.kind === 'solid') {
    if (patch.opacity === undefined) throw new CommandRefusal('disabled.noEffectChange');
    checked(patch.opacity, 1);
    const source = fillSource(batch, part, el);
    if (!source || source.raw.local !== 'solidFill' || !colorOf(source.spec))
      throw new CommandRefusal('errors.unsupported', 'solid fill');
    patchColor(
      batch,
      part,
      source.own ? source.raw : ownFill(batch, part, el, source.spec),
      undefined,
      patch.opacity
    );
    return;
  }
  const hasChange =
    patch.kind === 'gradient'
      ? patch.angle !== undefined ||
        patch.rotWithShape !== undefined ||
        patch.addStop !== undefined ||
        patch.removeStop !== undefined ||
        (patch.stop !== undefined &&
          [patch.stop.position, patch.stop.color, patch.stop.opacity].some((v) => v !== undefined))
      : patch.kind === 'shadow'
        ? [
            patch.blur,
            patch.distance,
            patch.angle,
            patch.color,
            patch.opacity,
            patch.rotWithShape,
          ].some((v) => v !== undefined)
        : [
            patch.blur,
            patch.distance,
            patch.startOpacity,
            patch.endOpacity,
            patch.startPosition,
            patch.endPosition,
            patch.rotWithShape,
          ].some((v) => v !== undefined);
  if (!hasChange) throw new CommandRefusal('disabled.noEffectChange');
  if (patch.rotWithShape !== undefined && typeof patch.rotWithShape !== 'boolean')
    throw new CommandRefusal('errors.unsupported', 'rotation');
  if (patch.kind === 'gradient') {
    const source = fillSource(batch, part, el);
    if (!source || source.raw.local !== 'gradFill')
      throw new CommandRefusal('errors.unsupported', 'gradient');
    const sourceStops = elements(child(source.raw, 'a:gsLst')).filter(
      (c) => c.ns === NS.a && c.local === 'gs'
    );
    const currentStopId = (stopId: string): string =>
      inheritedStopAliases.get(batch)?.get(stopAliasKey(part, id, stopId)) ?? stopId;
    const removeIndex =
      patch.removeStop === undefined
        ? -1
        : sourceStops.findIndex((s) => s.id === currentStopId(patch.removeStop!));
    if (patch.removeStop !== undefined && (removeIndex < 0 || sourceStops.length <= 2))
      throw new CommandRefusal('errors.unsupported', 'gradient stop');
    if (patch.addStop && sourceStops.length >= 10)
      throw new CommandRefusal('errors.unsupported', 'gradient stops');
    if (patch.stop && patch.stop.id === patch.removeStop)
      throw new CommandRefusal('errors.unsupported', 'gradient stop');
    const stopIndex = patch.stop
      ? sourceStops.findIndex((s) => s.id === currentStopId(patch.stop!.id))
      : -1;
    if (patch.stop && stopIndex < 0)
      throw new CommandRefusal('errors.unsupported', 'gradient stop');
    let fill = source.raw;
    if (!source.own) {
      fill = ownFill(batch, part, el, source.spec);
      const copiedStops = elements(child(fill, 'a:gsLst')).filter(
        (c) => c.ns === NS.a && c.local === 'gs'
      );
      const aliases = inheritedStopAliases.get(batch) ?? new Map<string, string>();
      sourceStops.forEach((stop, i) =>
        aliases.set(stopAliasKey(part, id, stop.id), copiedStops[i]!.id)
      );
      inheritedStopAliases.set(batch, aliases);
    }
    if (patch.angle !== undefined) {
      const lin = child(fill, 'a:lin');
      if (!lin) throw new CommandRefusal('errors.unsupported', 'gradient angle');
      set(batch, part, lin, 'ang', (checked(patch.angle, 360) * 60_000) % 21_600_000);
    }
    if (patch.rotWithShape !== undefined)
      set(batch, part, fill, 'rotWithShape', patch.rotWithShape);
    if (patch.stop) {
      const stop = elements(child(fill, 'a:gsLst')).filter(
        (c) => c.ns === NS.a && c.local === 'gs'
      )[stopIndex]!;
      if (patch.stop.position !== undefined)
        set(batch, part, stop, 'pos', checked(patch.stop.position, 1) * 100_000);
      if (patch.stop.color !== undefined || patch.stop.opacity !== undefined)
        patchColor(batch, part, stop, patch.stop.color, patch.stop.opacity);
    }
    const list = child(fill, 'a:gsLst');
    if ((patch.addStop || patch.removeStop !== undefined) && !list)
      throw new CommandRefusal('errors.malformed');
    if (patch.removeStop !== undefined) {
      const stop = elements(list).filter((c) => c.ns === NS.a && c.local === 'gs')[removeIndex]!;
      batch.push({ op: 'removeNode', part, node: stop.id });
    }
    if (patch.addStop) {
      const position = checked(patch.addStop.position, 1);
      const alpha = patch.addStop.opacity === undefined ? 1 : checked(patch.addStop.opacity, 1);
      const color = specFromXml(colorXml(patch.addStop.color));
      const current = batch.element(part, list!.id);
      const later = current.children.findIndex(
        (c) =>
          c.type === 'element' &&
          c.ns === NS.a &&
          c.local === 'gs' &&
          number(c, 'pos') > position * 100_000
      );
      batch.insert({
        op: 'insertNode',
        part,
        parent: current.id,
        index: later < 0 ? current.children.length : later,
        spec: {
          ns: NS.a,
          local: 'gs',
          prefix: 'a',
          attrs: [{ ns: '', local: 'pos', value: String(Math.round(position * 100_000)) }],
          children: [
            {
              ...color,
              children: [
                ...(color.children ?? []),
                {
                  ns: NS.a,
                  local: 'alpha',
                  prefix: 'a',
                  attrs: [{ ns: '', local: 'val', value: String(Math.round(alpha * 100_000)) }],
                },
              ],
            },
          ],
        },
      });
    }
    return;
  }
  const source = effectsSource(batch, part, el);
  const names = patch.kind === 'shadow' ? ['outerShdw', 'innerShdw', 'prstShdw'] : ['reflection'];
  const matches = source.filter((c) => c.ns === NS.a && names.includes(c.local));
  if (matches.length !== 1 || matches[0]!.local === 'prstShdw')
    throw new CommandRefusal('errors.unsupported', patch.kind);
  const list = effectListOf(
    batch,
    part,
    el,
    source.map((s) => ('id' in s ? specOf(s) : s))
  );
  const effect = elements(list).find((c) => c.ns === NS.a && c.local === matches[0]!.local)!;
  if (patch.blur !== undefined)
    set(batch, part, effect, 'blurRad', checked(patch.blur, 150) * 12_700);
  if (patch.distance !== undefined)
    set(batch, part, effect, 'dist', checked(patch.distance, 150) * 12_700);
  if (patch.rotWithShape !== undefined) {
    if (effect.local === 'innerShdw')
      throw new CommandRefusal('errors.unsupported', 'inner shadow rotation');
    set(batch, part, effect, 'rotWithShape', patch.rotWithShape);
  }
  if (patch.kind === 'shadow') {
    if (patch.angle !== undefined)
      set(batch, part, effect, 'dir', (checked(patch.angle, 360) * 60_000) % 21_600_000);
    if (patch.color !== undefined || patch.opacity !== undefined)
      patchColor(batch, part, effect, patch.color, patch.opacity);
  } else {
    for (const [field, xml] of [
      ['startOpacity', 'stA'],
      ['endOpacity', 'endA'],
      ['startPosition', 'stPos'],
      ['endPosition', 'endPos'],
    ] as const)
      if (patch[field] !== undefined)
        set(batch, part, effect, xml, checked(patch[field], 1) * 100_000);
  }
}
