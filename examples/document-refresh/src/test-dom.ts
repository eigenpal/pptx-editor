// Registers happy-dom for `agent-panel.test.tsx` before React loads. The test imports this
// module first, and unregisters the DOM after its last test, so the other tests of this
// example run with no global DOM, as a server does.

import { GlobalRegistrator } from '@happy-dom/global-registrator';

if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();
// React's `act()` warns unless the environment says it runs tests.
(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

/** Remove the global DOM again. */
export function unregisterDom(): Promise<void> {
  return GlobalRegistrator.isRegistered ? GlobalRegistrator.unregister() : Promise.resolve();
}
