// The React demo (`examples/vite`) in happy-dom: the app a person runs with `bun run dev`, with
// its own shell, menus, toolbar, rail, canvas, and notes, on the basic five-slide deck. The
// driver (`scripts/test/demo-dom.ts`) plays clicks, keys, and drags on it; the tests assert on
// the store, the snapshot, and the painted SVG.

import { mock } from 'bun:test';
import { act, cleanup, render } from '@testing-library/react';
import { openDemoWith, type Demo, type DemoOptions } from '../../../../scripts/test/demo-dom.ts';
import * as adapter from '../index.ts';

export * from '../../../../scripts/test/demo-dom.ts';
export { screen } from '@testing-library/react';

// The demo imports the published adapter name; the tests run the workspace source.
mock.module('@pptx-editor.dev/react', () => adapter);

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

// A variable specifier: the type check of this package does not follow it into the demo, which
// imports the adapter by its published name.
const APP = '../../../../examples/vite/src/App.tsx';

/** Open the React demo at a window size, on the basic deck, and wait until it paints. */
export async function openDemo(options: DemoOptions = {}): Promise<Demo> {
  const { App } = (await import(APP)) as { App: () => React.JSX.Element };
  return openDemoWith(
    {
      publicDir: new URL('../../../../examples/vite/public/', import.meta.url),
      mount: () => void render(<App />),
      unmount: cleanup,
      run: (step) => act(async () => void (await step())),
    },
    options
  );
}
