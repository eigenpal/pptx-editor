// View settings: how the editor shows the deck, not what the deck holds. Smart guides, the
// presentation view (Normal or Slide Sorter), and the collapsed sections change no part, take
// no undo step, and are never saved.

import type { EditorCommand, PresentationView } from '../contracts/index.ts';

/** The commands that change a view setting. */
export type ViewCommand = Extract<
  EditorCommand,
  { readonly type: 'toggleSmartGuides' | 'setPresentationView' | 'setSectionsCollapsed' }
>;

const VIEWS: ReadonlySet<PresentationView> = new Set(['normal', 'slideSorter']);
const NONE: readonly string[] = [];

export function isViewCommand(command: EditorCommand): command is ViewCommand {
  return (
    command.type === 'toggleSmartGuides' ||
    command.type === 'setPresentationView' ||
    command.type === 'setSectionsCollapsed'
  );
}

export class ViewSettings {
  /** Whether drags snap to smart guides (View > Smart guides). */
  smartGuides: boolean;
  /** View > Normal or Slide Sorter. */
  presentationView: PresentationView = 'normal';
  /** The collapsed section ids, reference-stable while they do not change. */
  collapsedSections: readonly string[] = NONE;

  constructor(smartGuides: boolean) {
    this.smartGuides = smartGuides;
  }

  /**
   * Why a view command cannot run, or null when it can: a view the editor does not have, or a
   * collapse in a deck without sections. `all` is every section id of the deck.
   */
  refusal(command: ViewCommand, all: readonly string[]): string | null {
    if (command.type === 'setPresentationView' && !VIEWS.has(command.view)) return 'errors.view';
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
  } {
    const { smartGuides, presentationView } = this;
    return { smartGuides, presentationView, collapsedSections: this.collapsedIn(all) };
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
    this.presentationView = 'normal';
    this.collapsedSections = NONE;
  }
}
