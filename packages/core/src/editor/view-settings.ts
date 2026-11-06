// View settings: how the editor shows the deck, not what the deck holds. Smart guides, the
// presentation view (Normal or Slide Sorter), the collapsed sections, and the table border pen
// change no part, take no undo step, and are never saved.

import type { EditorCommand, PresentationView, TableBorderPen } from '../contracts/index.ts';
import { isColorChoice } from '../automation/commands/style-commands.ts';
import { LINE_DASHES } from './line-presets.ts';

/** The pen before any pick: all borders, 1 pt solid in the theme's Text 1. @public */
export const DEFAULT_TABLE_BORDER_PEN: TableBorderPen = {
  target: 'all',
  drawn: false,
  spec: { color: { theme: 'tx1' }, width: 1 },
};

const PEN_TARGETS: ReadonlySet<unknown> = new Set([
  'all',
  'outside',
  'inside',
  'top',
  'bottom',
  'left',
  'right',
]);

/** A deep copy of a pen that no one can change: the pen, its line, and the line's color. */
function frozenPen(pen: TableBorderPen): TableBorderPen {
  const copy = JSON.parse(JSON.stringify(pen)) as TableBorderPen;
  if (typeof copy.spec.color === 'object' && copy.spec.color) Object.freeze(copy.spec.color);
  Object.freeze(copy.spec);
  return Object.freeze(copy);
}

/** Whether a host's pen is one the Borders menu can draw: an edge target and a valid line. */
function validPen(pen: unknown): pen is TableBorderPen {
  if (typeof pen !== 'object' || pen === null) return false;
  const { target, drawn, spec } = pen as Partial<TableBorderPen>;
  if (!PEN_TARGETS.has(target) || typeof drawn !== 'boolean' || !spec) return false;
  const width = spec.width;
  return (
    isColorChoice(spec.color) &&
    typeof width === 'number' &&
    width > 0 &&
    width <= 1584 &&
    (spec.dash === undefined || LINE_DASHES.includes(spec.dash))
  );
}

/** The commands that change a view setting. */
export type ViewCommand = Extract<
  EditorCommand,
  {
    readonly type:
      'toggleSmartGuides' | 'setPresentationView' | 'setSectionsCollapsed' | 'setTableBorderPen';
  }
>;

const VIEWS: ReadonlySet<PresentationView> = new Set(['normal', 'slideSorter']);
const NONE: readonly string[] = [];

export function isViewCommand(command: EditorCommand): command is ViewCommand {
  return (
    command.type === 'toggleSmartGuides' ||
    command.type === 'setPresentationView' ||
    command.type === 'setSectionsCollapsed' ||
    command.type === 'setTableBorderPen'
  );
}

export class ViewSettings {
  /** Whether drags snap to smart guides (View > Smart guides). */
  smartGuides: boolean;
  /** View > Normal or Slide Sorter. */
  presentationView: PresentationView = 'normal';
  /** The collapsed section ids, reference-stable while they do not change. */
  collapsedSections: readonly string[] = NONE;
  /** The table border pen; it stays when another deck opens, as a pen is a setting. */
  tableBorderPen: TableBorderPen = DEFAULT_TABLE_BORDER_PEN;

  constructor(smartGuides: boolean) {
    this.smartGuides = smartGuides;
  }

  /**
   * Why a view command cannot run, or null when it can: a view the editor does not have, or a
   * collapse in a deck without sections. `all` is every section id of the deck.
   */
  refusal(command: ViewCommand, all: readonly string[]): string | null {
    if (command.type === 'setPresentationView' && !VIEWS.has(command.view)) return 'errors.view';
    if (command.type === 'setTableBorderPen' && !validPen(command.pen)) return 'errors.unsupported';
    if (command.type === 'setSectionsCollapsed' && !all.length) return 'disabled.noSections';
    return null;
  }

  /** `refusal` as a refused command result, or null when the command can run. */
  refused(command: ViewCommand, all: readonly string[]): { ok: false; reason: string } | null {
    const reason = this.refusal(command, all);
    return reason ? { ok: false, reason } : null;
  }

  /** Apply a view command; false when `refusal` names a reason. */
  apply(command: ViewCommand, all: readonly string[]): boolean {
    if (this.refusal(command, all)) return false;
    if (command.type === 'toggleSmartGuides') this.smartGuides = !this.smartGuides;
    // A copy the host cannot change after it passed `validPen`.
    else if (command.type === 'setTableBorderPen') this.tableBorderPen = frozenPen(command.pen);
    else if (command.type === 'setPresentationView') this.presentationView = command.view;
    else {
      const ids = new Set(this.collapsedSections);
      for (const id of command.sectionIds ?? all)
        if (command.collapsed) ids.add(id);
        else ids.delete(id);
      // Only sections the deck has: a removed section's id goes.
      const kept = all.filter((id) => ids.has(id));
      this.collapsedSections = kept.length ? kept : NONE;
    }
    return true;
  }

  /** The last collapsed list the snapshot filtered, and what it gave, for a stable reference. */
  private shown: { from: readonly string[]; key: string; out: readonly string[] } | undefined;

  /**
   * The snapshot fields of the view settings. `all` is every section id of the deck now: a
   * section that an undo, a remote edit, or a reload removed is not reported as collapsed.
   */
  snapshot(all: readonly string[]): {
    smartGuides: boolean;
    presentationView: PresentationView;
    collapsedSections: readonly string[];
    tableBorderPen: TableBorderPen;
  } {
    const { smartGuides, presentationView, tableBorderPen } = this;
    return {
      smartGuides,
      presentationView,
      collapsedSections: this.collapsedIn(all),
      tableBorderPen,
    };
  }

  /** The collapsed ids that are sections of the deck, reference-stable while they stay. */
  private collapsedIn(all: readonly string[]): readonly string[] {
    const from = this.collapsedSections;
    if (from === NONE) return NONE;
    const ids = new Set(all);
    if (from.every((id) => ids.has(id))) return from;
    const out = from.filter((id) => ids.has(id));
    const key = out.join('\n');
    if (this.shown?.from === from && this.shown.key === key) return this.shown.out;
    this.shown = { from, key, out: out.length ? out : NONE };
    return this.shown.out;
  }

  /** A new document opens in the Normal view with every section open. */
  reset(): void {
    // The pen keeps its line; it redraws no target until a border pick in the new deck.
    this.tableBorderPen = { ...this.tableBorderPen, drawn: false };
    this.presentationView = 'normal';
    this.collapsedSections = NONE;
  }
}
