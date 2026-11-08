// One undo step for one continuous gesture on a host control (a slider drag, a held arrow key,
// a native color picker). Adapted from DOCX Editor's
// packages/core/src/editor/bind-history-group.ts (commit 2741998e3), over this editor's
// `beginHistoryGroup()`. The framework's own change handlers apply the values; this binding
// only opens and ends the group.

import type { EditorExecOptions, HistoryGroup } from '../contracts/history-group.ts';

/** Native event policy for one continuous control gesture. @public */
export interface HistoryGroupBindingOptions {
  readonly kind: 'native-color' | 'range' | 'repeat' | 'keyboard';
}

/** Dispose when the element or editor changes. @public */
export interface HistoryGroupBinding {
  /** Options for the next synchronous command: `{ historyGroup }`. */
  options(): EditorExecOptions;
  /** Close the gesture and remove every listener. The edits stay. */
  dispose(): void;
}

/** The editor members a binding uses. */
interface HistoryGroupHost {
  beginHistoryGroup(): HistoryGroup;
}

const STEP_KEYS = [
  'ArrowLeft',
  'ArrowRight',
  'ArrowUp',
  'ArrowDown',
  'Home',
  'End',
  'PageUp',
  'PageDown',
  ' ',
  'Enter',
];

/**
 * Bind the native gesture events of `element`. A native color `change` commits the value.
 * Range and repeat controls end on pointer release or cancel, or on key release, also outside
 * the element. A control with no terminal event ends on the next activation, Escape, blur, or
 * disposal. A terminal event closes the group after it propagates, so the final value joins
 * the same group. @public
 */
export function bindHistoryGroup(
  editor: HistoryGroupHost,
  element: HTMLElement,
  config: HistoryGroupBindingOptions
): HistoryGroupBinding {
  let group: HistoryGroup | undefined;
  let disposed = false;
  let pointer: number | undefined;
  const view = element.ownerDocument.defaultView;
  const setTimer = (run: () => void): number =>
    (view?.setTimeout ?? setTimeout)(run, 0) as unknown as number;
  const clearTimer = (id: number | undefined): void => {
    if (id !== undefined) (view?.clearTimeout ?? clearTimeout)(id);
  };
  let terminalTimer: number | undefined;
  const keys = new Set<string>();
  const removers: (() => void)[] = [];
  const end = () => {
    clearTimer(terminalTimer);
    terminalTimer = undefined;
    group?.end();
    group = undefined;
    pointer = undefined;
    keys.clear();
  };
  const ensure = () => {
    if (disposed) throw new Error('history group binding was disposed');
    if (!group || group.state === 'closed') group = editor.beginHistoryGroup();
    return group;
  };
  const start = () => {
    end();
    ensure();
  };
  const finish = () => {
    const ending = group;
    pointer = undefined;
    keys.clear();
    clearTimer(terminalTimer);
    // Event dispatch can run a microtask checkpoint between listeners. A task deferral keeps
    // the framework's delegated terminal handler in the same gesture.
    terminalTimer = setTimer(() => {
      if (group === ending) end();
    });
  };
  const listen = (
    target: EventTarget,
    type: string,
    listener: (event: Event) => void,
    capture = false
  ) => {
    target.addEventListener(type, listener, capture);
    removers.push(() => target.removeEventListener(type, listener, capture));
  };
  const relevant = (key: string) =>
    config.kind === 'native-color' ? key === 'Enter' || key === ' ' : STEP_KEYS.includes(key);
  listen(
    element,
    'pointerdown',
    (event) => {
      const e = event as PointerEvent;
      if (e.button !== 0 || config.kind === 'keyboard') return;
      start();
      pointer = e.pointerId;
      if (config.kind !== 'native-color') {
        // Capture is optional: synthetic events and detached nodes cannot capture.
        try {
          element.setPointerCapture?.(e.pointerId);
        } catch {
          /* the document listeners still end the gesture */
        }
      }
    },
    true
  );
  listen(
    element,
    'keydown',
    (event) => {
      const e = event as KeyboardEvent;
      if (e.key === 'Escape') {
        finish();
        return;
      }
      if (!relevant(e.key) || e.altKey || e.ctrlKey || e.metaKey) return;
      if (!e.repeat && keys.size === 0) start();
      keys.add(e.key);
      ensure();
    },
    true
  );
  listen(
    element.ownerDocument,
    'keyup',
    (event) => {
      const key = (event as KeyboardEvent).key;
      if (!keys.delete(key)) return;
      if (config.kind !== 'native-color' && keys.size === 0) finish();
    },
    true
  );
  if (config.kind !== 'native-color') {
    for (const type of ['pointerup', 'pointercancel', 'lostpointercapture']) {
      listen(
        type === 'lostpointercapture' ? element : element.ownerDocument,
        type,
        (event) => {
          if (pointer !== undefined && (event as PointerEvent).pointerId === pointer) finish();
        },
        true
      );
    }
  }
  listen(
    element,
    'input',
    () => {
      if (terminalTimer !== undefined) start();
      else ensure();
    },
    true
  );
  listen(element, 'change', () => {
    if (config.kind === 'native-color' || (pointer === undefined && keys.size === 0)) finish();
  });
  listen(element, 'blur', finish, true);
  if (view) listen(view, 'blur', finish);
  return {
    options: () => ({ historyGroup: ensure() }),
    dispose() {
      if (disposed) return;
      disposed = true;
      end();
      for (const remove of removers) remove();
    },
  };
}
