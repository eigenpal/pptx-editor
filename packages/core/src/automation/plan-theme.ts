// The theme colors of a slide, a layout, or a master: the `a:clrScheme` of the theme its master
// relates to (CT_ColorScheme, ECMA-376 Part 1, 20.1.6.2). A write changes the theme part, so
// every slide, layout, and master that uses that theme changes with it.

import { parseTheme, toHex } from '../cascade/index.ts';
import { child, elements, path, xmlPartOf } from '../store/index.ts';
import { specFromXml } from './commands/shape-commands.ts';
import { backgroundOwner } from './plan-backgrounds.ts';
import { colorSpec } from './plan-text.ts';
import type { AutomationRef, AutomationValue } from './protocol.ts';
import { AutomationFailure, badArgument, edit, type PlanContext } from './targets.ts';

/** The twelve colors of CT_ColorScheme, in schema order. */
const SLOTS: ReadonlySet<string> = new Set([
  'dk1',
  'lt1',
  'dk2',
  'lt2',
  'accent1',
  'accent2',
  'accent3',
  'accent4',
  'accent5',
  'accent6',
  'hlink',
  'folHlink',
]);

/** A color slot, or an `invalid-argument` failure. */
function slotArg(slot: unknown): string {
  if (typeof slot !== 'string' || !SLOTS.has(slot)) badArgument('color');
  return slot;
}

/** The theme part of the owner a ref names, or an `unsupported-content` failure. */
function themePartOf(ctx: PlanContext, ref: AutomationRef): string {
  const part = backgroundOwner(ctx, ref).themePart;
  if (!part || !xmlPartOf(ctx.batch.pkg, part))
    throw new AutomationFailure('unsupported-content', 'the master has no theme', 'no-theme');
  return part;
}

/** One theme color as `#RRGGBB`. A theme with no such color reads the default theme's. */
export function readThemeColor(
  ctx: PlanContext,
  ref: AutomationRef,
  slot: unknown
): AutomationValue {
  const name = slotArg(slot);
  const theme = parseTheme(xmlPartOf(ctx.batch.pkg, themePartOf(ctx, ref)));
  const color = theme.colors.get(name);
  if (!color) throw new AutomationFailure('unsupported-content', 'no such theme color', 'no-theme');
  return { kind: 'text', text: toHex({ ...color, a: 1 }).toUpperCase() };
}

/** Replace the color of one slot with `#RRGGBB` (`a:srgbClr`) or a preset name (`a:prstClr`). */
export function setThemeColor(
  ctx: PlanContext,
  ref: AutomationRef,
  slot: unknown,
  color: unknown
): AutomationValue {
  const name = slotArg(slot);
  const xml = colorSpec(color, 'color');
  const part = themePartOf(ctx, ref);
  const scheme = path(xmlPartOf(ctx.batch.pkg, part)?.root, 'a:themeElements', 'a:clrScheme');
  const el = child(scheme, `a:${name}`);
  if (!el)
    throw new AutomationFailure('unsupported-content', 'the theme has no such color', 'no-theme');
  edit(() => {
    // CT_Color holds one color choice (EG_ColorChoice).
    for (const c of elements(el)) ctx.batch.push({ op: 'removeNode', part, node: c.id });
    ctx.batch.insert({ op: 'insertNode', part, parent: el.id, index: 0, spec: specFromXml(xml) });
  });
  return { kind: 'applied' };
}
