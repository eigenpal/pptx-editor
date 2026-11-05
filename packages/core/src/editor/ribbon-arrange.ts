// How a ribbon draws its tabs from the core arrangement and the host's tab overrides. Shared by
// the React and Vue ribbons, so a `Ribbon.Tab` means the same in both, as
// `arrangeToolbarGroups` does for the toolbar's groups.

import type { ChromeGroup } from './chrome-controls.ts';
import { warnRibbonTab } from './dev-warnings.ts';
import { RIBBON_TABS, type RibbonTabLayout } from './ribbon.ts';

/** What a host tab override names. @internal */
export interface RibbonTabOverride {
  /** A built-in tab id changes that tab; any other id adds a host tab. */
  readonly id: string;
  readonly label?: string | undefined;
  readonly labelKey?: string | undefined;
  readonly hidden?: boolean | undefined;
  readonly after?: string | undefined;
}

/** A tab as the ribbon draws it, with the override that changed it. @internal */
export interface ArrangedRibbonTab<S extends RibbonTabOverride> {
  readonly id: string;
  readonly title: string;
  readonly groups: readonly ChromeGroup[];
  readonly contextual: boolean;
  readonly override: S | undefined;
}

const BUILT_IN_TABS: readonly string[] = RIBBON_TABS.map((tab) => tab.id);

/**
 * The tabs to draw: the built-in tabs that show now (`layout`, from `ribbonTabs`), changed by
 * the host's overrides, then the host's tabs. With `preset` false, only the overridden tabs
 * draw: a built-in id keeps its title and groups. A built-in id whose context is off (a
 * contextual tab with nothing selected) draws nothing; it is never a host tab. A setting that
 * cannot work prints a development warning with the adapter's part name (`tabPart`). The Quick
 * Access Toolbar comes back apart from the tab names. @internal
 */
export function arrangeRibbonTabs<S extends RibbonTabOverride>(
  layout: readonly RibbonTabLayout[],
  overrides: ReadonlyMap<string, S>,
  preset: boolean,
  translate: (key: string) => string,
  tabPart: string
): { readonly quickAccess?: ArrangedRibbonTab<S>; readonly tabs: ArrangedRibbonTab<S>[] } {
  const titleOf = (override: S | undefined, fallback: string): string =>
    override?.label ?? (override?.labelKey !== undefined ? translate(override.labelKey) : fallback);
  const drawn: ArrangedRibbonTab<S>[] = [];
  for (const { tab, groups } of layout) {
    const override = overrides.get(tab.id);
    if (override?.hidden || (!preset && !override)) continue;
    drawn.push({
      id: tab.id,
      title: titleOf(override, translate(tab.labelKey)),
      groups,
      contextual: tab.context !== undefined,
      override,
    });
  }
  const known = [...BUILT_IN_TABS, ...overrides.keys()];
  for (const override of overrides.values()) {
    if (BUILT_IN_TABS.includes(override.id)) {
      if (override.after !== undefined)
        warnRibbonTab(tabPart, override.id, 'builtInAfter', BUILT_IN_TABS);
      continue;
    }
    if (override.hidden) {
      warnRibbonTab(tabPart, override.id, 'hiddenUnknown', BUILT_IN_TABS);
      continue;
    }
    if (override.label === undefined && override.labelKey === undefined)
      warnRibbonTab(tabPart, override.id, 'untitled', BUILT_IN_TABS);
    const host: ArrangedRibbonTab<S> = {
      id: override.id,
      title: titleOf(override, override.id),
      groups: [],
      contextual: false,
      override,
    };
    const anchor = override.after;
    if (anchor !== undefined && !known.includes(anchor))
      warnRibbonTab(tabPart, override.id, 'unknownAfter', BUILT_IN_TABS, anchor);
    const at = anchor === undefined ? -1 : drawn.findIndex((d) => d.id === anchor);
    if (at === -1) drawn.push(host);
    else drawn.splice(at + 1, 0, host);
  }
  const quickAccess = drawn.find((d) => d.id === 'quickAccess');
  const tabs = drawn.filter((d) => d.id !== 'quickAccess');
  return quickAccess ? { quickAccess, tabs } : { tabs };
}

/**
 * The tab to show: the wanted one, else Home, else the first. `fellBack` says that the wanted
 * tab is gone (a contextual tab after the selection changed), so the ribbon selects the shown
 * tab and tells the host, and a contextual tab does not come back by itself later. @internal
 */
export function ribbonShownTab<T extends { readonly id: string }>(
  tabs: readonly T[],
  wanted: string
): { readonly shown: T | undefined; readonly fellBack: boolean } {
  const exact = tabs.find((d) => d.id === wanted);
  if (exact) return { shown: exact, fellBack: false };
  const shown = tabs.find((d) => d.id === 'home') ?? tabs[0];
  return { shown, fellBack: shown !== undefined };
}

/**
 * The tab to open after an insert: the first contextual tab that shows for the new selection
 * (Shape Format, Picture Format, or Table Design). Null
 * when no contextual tab shows. @internal
 */
export function ribbonInsertTab<T extends { readonly id: string; readonly contextual: boolean }>(
  tabs: readonly T[]
): string | null {
  return tabs.find((tab) => tab.contextual)?.id ?? null;
}
