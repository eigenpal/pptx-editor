// The React binding of core's side-panel field controller (`live-field.ts`). The Vue twin is
// `packages/vue/src/live-field.ts`.
import { useEffect, useReducer, useRef, useState } from 'react';
import { createLiveField, type LiveField, type LiveFieldHost } from '@pptx-editor.dev/core/editor';

/**
 * Bind one side-panel field. `scope` is the document value and any key that drops the draft
 * when it changes, unless the field's own edit made the change. @internal
 */
export function useLiveField(host: LiveFieldHost, scope: unknown): LiveField {
  const [, render] = useReducer((n: number) => n + 1, 0);
  const latest = useRef(host);
  latest.current = host;
  const [field] = useState(() => createLiveField(() => latest.current, render));
  const [seen, setSeen] = useState(scope);
  if (!Object.is(seen, scope)) {
    setSeen(scope);
    field.sync();
  }
  useEffect(() => () => field.dispose(), [field]);
  return field;
}
