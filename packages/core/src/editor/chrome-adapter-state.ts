// The state of the slots an adapter runs itself (the slide background dialog, the clipboard,
// the slide show, and the find bar). Split from `chrome-controls.ts`, which asks it from
// `toolbarCommandState`, so the registry stays under the line cap.

import type { ClipboardPayload, Editor } from '../contracts/index.ts';
import type { ChromeSlotId, ToolbarCommandState } from './chrome-controls.ts';

/** Whether `ids` on `slide` include a frame the clipboard cannot carry (all but a table). */
function holdsUncopiedFrame(editor: Editor, slide: number, ids: readonly string[]): boolean {
  const session = (editor as { readonly session?: { layout(i: number): SlideLayoutLike } | null })
    .session;
  const chosen = new Set(ids);
  // A group carries its children: a chart inside a selected group counts too.
  const frameIn = (item: LayoutItemLike): boolean =>
    (item.element.kind === 'graphicFrame' && item.element.frameKind !== 'table') ||
    (item.children ?? []).some(frameIn);
  // A chosen shape at any depth: a chart picked inside a group (`selection.group`) counts.
  const chosenFrame = (item: LayoutItemLike): boolean =>
    (chosen.has(item.element.nodeId) && frameIn(item)) || (item.children ?? []).some(chosenFrame);
  return (session?.layout(slide)?.items ?? []).some(chosenFrame);
}

interface LayoutItemLike {
  readonly element: { readonly nodeId: string; readonly kind: string; readonly frameKind?: string };
  readonly children?: readonly LayoutItemLike[];
}

interface SlideLayoutLike {
  readonly items: readonly LayoutItemLike[];
}

/** Enabled state of the slots an adapter runs: dialogs, clipboard, show, find. @internal */
export function adapterActionState(editor: Editor, id: ChromeSlotId): ToolbarCommandState {
  const snap = editor.snapshot();
  const off = (reason: string): ToolbarCommandState => ({
    id,
    enabled: false,
    disabledReason: reason,
    active: false,
  });
  const on: ToolbarCommandState = { id, enabled: true, disabledReason: null, active: false };
  if (snap.status !== 'ready') return off('disabled.noEditor');
  const viewing = snap.mode === 'view';
  const sel = snap.selection;
  const textRange =
    sel.kind === 'text' &&
    (sel.anchor.paragraph !== sel.focus.paragraph || sel.anchor.offset !== sel.focus.offset);
  switch (id) {
    case 'slide.background': {
      const can = editor.can({ type: 'setSlideBackground', background: null });
      return can.ok ? on : off(can.reason);
    }
    case 'edit.cut':
      if (viewing) return off('disabled.viewing');
      if (sel.kind === 'cells') return on;
      if (sel.kind === 'text') return textRange ? on : off('contextMenu.noTextSelected');
      if (sel.kind !== 'shapes') return off('disabled.noSelection');
      // The clipboard cannot carry a chart, diagram, or object frame: Cut would lose it.
      return holdsUncopiedFrame(editor, snap.currentSlide, sel.nodeIds)
        ? off('contextMenu.cannotCopyFrame')
        : on;
    case 'edit.copy':
      if (sel.kind === 'cells') return on;
      if (sel.kind === 'text') return textRange ? on : off('contextMenu.noTextSelected');
      if (sel.kind !== 'shapes') return off('disabled.noSelection');
      return holdsUncopiedFrame(editor, snap.currentSlide, sel.nodeIds)
        ? off('contextMenu.cannotCopyFrame')
        : on;
    case 'edit.paste': {
      if (viewing) return off('disabled.viewing');
      // Text and cell paste read the system clipboard; shape paste needs this editor's last copy.
      if (sel.kind === 'text' || sel.kind === 'cells') return on;
      const copied = (editor as { readonly lastCopied?: ClipboardPayload | null }).lastCopied;
      if (!copied) return off('contextMenu.nothingCopied');
      const can = editor.can({ type: 'paste', payload: copied });
      return can.ok ? on : off(can.reason);
    }
    // Find reads any open deck and shows its matches on the slide canvas, which the slide
    // sorter does not show. Replace edits the deck, so it also needs editing mode.
    case 'find.find':
    case 'find.replace':
      if (snap.presentationView === 'slideSorter') return off('disabled.findSorter');
      return id === 'find.replace' && viewing ? off('disabled.viewing') : on;
    case 'view.playFromStart':
    case 'view.playFromCurrent':
      return snap.slides.length > 0 ? on : off('disabled.noEditor');
    default:
      return on;
  }
}
