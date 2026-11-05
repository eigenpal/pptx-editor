/** @jsxImportSource react */
// The packaged link popup, kept for `popups.hyperlink` renderers that forward their props to it.
// It renders `PptxEditor.HyperLink` (`HyperLink.tsx`): over the root's one popup state inside
// `PptxEditor.Root`, else over a state of its own built from its props. `className` goes on the
// panel.

import { useContext, useEffect, useId, useMemo, useSyncExternalStore, type ReactNode } from 'react';
import { hyperlinkPopupFromRequest } from '@pptx-editor.dev/core/editor';
import { usePptxEditor } from './context.ts';
import { HyperLink, HyperlinkPopupContext } from './HyperLink.tsx';
import type { HyperlinkPopupProps } from './popups.tsx';

/**
 * The packaged link popup: `PptxEditor.HyperLink` with its default parts. A `popups.hyperlink`
 * renderer can forward its props to it, or render `PptxEditor.HyperLink` with its own parts.
 * @public
 */
export function HyperlinkPopup(props: HyperlinkPopupProps): ReactNode {
  const shared = useContext(HyperlinkPopupContext);
  const panel = <HyperLink {...(props.className ? { className: props.className } : {})} />;
  return shared ? panel : <OwnState {...props}>{panel}</OwnState>;
}

/** A popup rendered outside its root: a state of its own, from its props. */
function OwnState({ children, ...props }: HyperlinkPopupProps & { children: ReactNode }) {
  const editor = usePptxEditor();
  const errorId = useId();
  const { anchor, url, mode, onClose } = props;
  // One state per opening: the props are what it opened with.
  const controller = useMemo(
    () => hyperlinkPopupFromRequest(editor, { anchor, url, onClose, ...(mode ? { mode } : {}) }),
    [editor, anchor] // eslint-disable-line react-hooks/exhaustive-deps
  );
  useEffect(() => () => controller.dispose(), [controller]);
  const state = useSyncExternalStore(
    controller.subscribe,
    controller.getState,
    controller.getState
  );
  return (
    <HyperlinkPopupContext.Provider value={{ controller, state, errorId }}>
      {children}
    </HyperlinkPopupContext.Provider>
  );
}
