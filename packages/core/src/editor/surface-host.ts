// What the slide surface asks of `createPptxEditor`: the `SurfaceHost` over the editor's
// state. The editor passes its own calls through (`SurfaceHostBase`); this file adds the ones
// that read the session, the selection, and the editor's parts, and the surface's key handler
// (`createShortcut`).

import { emptyPlaceholderType } from '../automation/commands/placeholder-fill.ts';
import { parentGroupOf } from '../automation/commands/placement.ts';
import type {
  CommandResult,
  EditorCommand,
  EditorMode,
  EditorSelection,
} from '../contracts/index.ts';
import type { SlideLayout } from '../layout/index.ts';
import { relayoutItem } from '../layout/item-layout.ts';
import { elements, shapeTreeOf, type PptxPackage } from '../store/index.ts';
import type { PresenceLink } from './collaboration-presence.ts';
import type { DocumentSession } from './document-session.ts';
import type { FormatPainter } from './format-painter.ts';
import type { EditorHistoryGroups } from './history-groups.ts';
import type { PlaceholderPrompts } from './placeholder-prompts.ts';
import { aspectLockedIn } from './shape-list.ts';
import { altTextAt } from './picture-state.ts';
import type { SlideSurface, SurfaceHost } from './surface.ts';
import type { ViewSettings } from './view-settings.ts';
import { formatKeyCommand } from './format-keys.ts';
import { formatKeyTarget } from './selection-state.ts';
import { listKeyAction } from './list-keys.ts';
import { shapeKeyAction } from './shape-keys.ts';
import { textItemOf } from './table-cells.ts';
import { applyTableKey, tableKeyAction } from './table-keys.ts';
import { isCaretKey, moveTextSelection, type TextInputKind } from './text-input.ts';
import { zoomKeyAction } from './zoom.ts';

/** The surface calls the editor passes through as they are. */
export type SurfaceHostBase = Pick<
  SurfaceHost,
  | 'idPrefix'
  | 'pickImage'
  | 'layout'
  | 'label'
  | 'select'
  | 'copy'
  | 'paste'
  | 'warmCopy'
  | 'cropping'
  | 'textInput'
  | 'shortcut'
  | 'run'
  | 'beginHistoryGroup'
  | 'refuse'
  | 'fitChanged'
  | 'zoom'
  | 'setZoom'
>;

/** The editor state and parts the other surface calls read. */
export interface SurfaceHostState {
  session(): DocumentSession | null;
  selection(): EditorSelection;
  viewing(): boolean;
  /** The current slide's part. */
  part(): string | undefined;
  surface(): SlideSurface | null;
  readonly prompts: PlaceholderPrompts;
  readonly presence: PresenceLink;
  readonly painter: FormatPainter;
  readonly views: ViewSettings;
  readonly groups: EditorHistoryGroups;
  /** One drag step: the steps of one gesture are one undo step. */
  gestureEdit(command: EditorCommand, gesture: string, base?: PptxPackage): void;
}

export function createSurfaceHost(base: SurfaceHostBase, s: SurfaceHostState): SurfaceHost {
  return {
    ...base,
    paintLayout: (l) => s.prompts.paint(s.session(), l, s.selection(), s.viewing()) ?? l,
    selection: () => s.selection(),
    readOnly: () => s.viewing(),
    imageUrl: (p, id) => s.session()?.imageUrl(p, id),
    fillableType: (id) => {
      const pkg = s.session()?.store.package;
      const part = s.part();
      if (!pkg || !part) return undefined;
      const el = elements(shapeTreeOf(pkg, part)).find((c) => c.id === id);
      return el ? emptyPlaceholderType(el) : undefined;
    },
    imageSize: (p, id) => s.session()?.imageSize(p, id),
    altText: (p, id) => altTextAt(s.session()?.store.package, p, id),
    presence: () => s.presence.painted(s.session()?.store.package, s.part()),
    transform: (items, gesture) => {
      const session = s.session();
      if (!session) return;
      // The package when this drag started: a table resize scales from it.
      const base = s.groups.drag(s.surface(), gesture, session.store.package).base ?? undefined;
      s.gestureEdit({ type: 'setTransforms', items }, gesture, base);
    },
    crop: (crop, gesture) => s.gestureEdit({ type: 'cropImage', crop }, gesture),
    adjust: (values, gesture) => s.gestureEdit({ type: 'setAdjustments', values }, gesture),
    painting: () => s.painter.mode !== 'off',
    // A selection made with the pointer while the painter is on takes the copied format.
    pointerUp: () => void (s.painter.mode !== 'off' && base.run({ type: 'pasteFormat' })),
    slideCount: () => s.session()?.info.slides.length ?? 0,
    smartGuides: () => s.views.smartGuides,
    aspectLocked: (id) => aspectLockedIn(s.session()?.store.package, s.part(), id),
    relayout: (item, xfrm) => {
      const session = s.session();
      return session ? relayoutItem(item, xfrm, session.measurer) : null;
    },
  };
}

/** What the surface's keys ask of the editor (`createShortcut`). */
export interface ShortcutHost {
  /** The editor, which a link request names. */
  editor(): unknown;
  layout(): SlideLayout | undefined;
  selection(): EditorSelection;
  /** Set the selection and publish it. */
  setSelection(next: EditorSelection): void;
  mode(): EditorMode;
  exec(command: EditorCommand): CommandResult;
  select(next: EditorSelection): void;
  selectAll(): void;
  textInput(kind: TextInputKind, data?: string): unknown;
  /** Enter text editing at the end of a shape's text. */
  enterText(nodeId: string): boolean;
  surface(): SlideSurface | null;
  fitZoom(): void;
  /** Put the format painter or crop mode away; true when one was on. */
  cancelTools(): boolean;
  /** Close the typing burst: the next edit is an undo step of its own. */
  boundary(): void;
  /** Whether AutoCorrect started a list at the caret, with no change since. */
  autoListed(): boolean;
}

/** The surface's key handler: true when the key was used. */
export function createShortcut(h: ShortcutHost): (key: string, e: KeyboardEvent) => boolean {
  /** Move the caret (arrow keys), optionally extending the selection. */
  const moveCaret = (key: string, e: KeyboardEvent): void => {
    const selection = h.selection();
    if (selection.kind !== 'text') return;
    const item = textItemOf(h.layout(), selection);
    if (!item?.text) return;
    h.boundary();
    h.setSelection(moveTextSelection(item, selection, key, e.shiftKey, e));
  };
  return (key, e) => {
    const selection = h.selection();
    const layout = h.layout;
    const mod = e.metaKey || e.ctrlKey;
    const lower = key.toLowerCase();
    // The format and history keys (format-keys.ts); in text, Ctrl/Cmd+= is subscript.
    const keyed = formatKeyCommand(key, e, formatKeyTarget(layout(), selection));
    if (keyed) return (h.exec(keyed), true);
    // Ctrl/Cmd + = and − zoom, and Ctrl/Cmd + 0 fits the slide to the window (zoom.ts).
    const zoomed = zoomKeyAction(key, mod);
    if (zoomed) return (zoomed === 'fit' ? h.fitZoom() : h.surface()?.zoomBy(zoomed), true);
    // The format painter: Escape puts it away.
    if (key === 'Escape' && h.cancelTools()) return true;
    // Ctrl/Cmd+K asks the chrome for its link popover; Ctrl/Cmd+\\ clears formatting.
    if (mod && lower === 'k') {
      // The detail names the editor, so a toolbar answers only its own editor's request.
      h.surface()?.root.dispatchEvent(
        new CustomEvent('pptx-link-request', { bubbles: true, detail: { editor: h.editor() } })
      );
      return true;
    }
    // Tab, Shift+Tab, and arrows at a cell's edge move between table cells (table-keys.ts).
    const cellKey = tableKeyAction(layout(), selection, key, e.shiftKey, mod || e.altKey);
    if (cellKey) return (applyTableKey(cellKey, h.exec, layout, h.select), true);
    if (selection.kind === 'text') {
      if (isCaretKey(key)) {
        moveCaret(key, e);
        return true;
      }
      if (mod && lower === 'a') return (h.selectAll(), true);
      if (key === 'Escape') {
        const group = parentGroupOf(layout()?.items ?? [], selection.nodeId);
        h.setSelection({
          kind: 'shapes',
          nodeIds: [selection.nodeId],
          ...(group ? { group } : {}),
        });
        h.surface()?.focus();
        return true;
      }
      // Backspace right after AutoCorrect started a list gives the typed marker back (undo).
      const plain = !e.altKey && !e.metaKey && !e.ctrlKey && !e.shiftKey;
      if (key === 'Backspace' && plain && h.autoListed()) return (h.exec({ type: 'undo' }), true);
      // Enter, Tab, and Backspace with lists (list-keys.ts).
      const list = listKeyAction(textItemOf(layout(), selection), selection, key, {
        shift: e.shiftKey,
        other: mod || e.altKey,
      });
      if (list)
        return (
          list.kind === 'input' ? h.textInput(list.input, list.data) : h.exec(list.command),
          true
        );
      return false;
    }
    if (mod && lower === 'a') return (h.selectAll(), true);
    if (mod && lower === 'g') return (h.exec({ type: e.shiftKey ? 'ungroup' : 'group' }), true);
    if (selection.kind === 'shapes' || selection.kind === 'none') {
      // Tab, nudge, delete, duplicate, Escape, and Enter on selected shapes (shape-keys.ts).
      const keys = { shift: e.shiftKey, mod, alt: e.altKey };
      const act = shapeKeyAction(key, keys, selection, layout(), h.mode());
      if (act?.kind === 'leave') return (h.select({ kind: 'none' }), false);
      if (act?.kind === 'exec') return (h.exec(act.command), true);
      if (act?.kind === 'select') return (h.select(act.selection), true);
      if (act?.kind === 'enterText') {
        if (!h.enterText(act.nodeId)) return true;
        if (act.type) h.textInput('text', act.type);
        else h.selectAll();
        return true;
      }
      if (act) return true;
    }
    return false;
  };
}
