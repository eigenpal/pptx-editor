/** @jsxImportSource react */
// The chrome tooltip bubble, after DOCX Editor's `Tooltip` (packages/react/src/components/ui/
// Tooltip.tsx). The editor root mounts one; it shows the `data-pptx-tooltip` label of the control
// under the pointer, or of the control a keyboard user moved to, after core's delay, with a
// disabled control's reason on its own line. Core's `chrome-tooltip.ts` owns the timing and the
// position, and the Vue twin renders the same markup.

import { useEffect, useLayoutEffect, useRef, useState } from 'react';
import { createPortal } from 'react-dom';
import {
  listenChromeTooltips,
  positionChromeTooltip,
  type ChromeTooltipState,
} from '@pptx-editor.dev/core/editor';

/** The tooltip bubble for every chrome control on the page. @internal */
export function ChromeTooltip() {
  const [state, setState] = useState<ChromeTooltipState | null>(null);
  const ref = useRef<HTMLDivElement>(null);
  useEffect(
    () => (typeof document === 'undefined' ? undefined : listenChromeTooltips(document, setState)),
    []
  );
  // Before paint: the bubble measures itself, then core places it.
  useLayoutEffect(() => {
    if (state && ref.current) positionChromeTooltip(ref.current, state.target);
  }, [state]);
  if (!state) return null;
  return createPortal(
    // Visual only: the control's `aria-label` and `aria-describedby` carry the same text.
    <div ref={ref} className="pptx-tooltip" role="tooltip" aria-hidden="true">
      <span className="pptx-tooltip__label">{state.label}</span>
      {state.reason ? <span className="pptx-tooltip__reason">{state.reason}</span> : null}
    </div>,
    state.scope
  );
}
