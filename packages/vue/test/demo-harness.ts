// The Vue demo (`examples/vue`) in happy-dom: the app a person runs with `bun run dev:vue`, with
// its own shell, menus, toolbar, rail, canvas, and notes, on the basic five-slide deck. Its
// single-file components compile here with Vue's own compiler (`vue/compiler-sfc`, as Vite's
// plugin does). The driver (`scripts/test/demo-dom.ts`) plays clicks, keys, and drags on it.

import { plugin } from 'bun';
import { mock } from 'bun:test';
import { readFileSync } from 'node:fs';
import { compileScript, parse } from 'vue/compiler-sfc';
import { createApp, nextTick, type App, type Component } from 'vue';
import { openDemoWith, type Demo, type DemoOptions } from '../../../scripts/test/demo-dom.ts';
import * as adapter from '../src/index.ts';

export * from '../../../scripts/test/demo-dom.ts';
export { screen, within } from '@testing-library/react';

// The demo imports the published adapter name; the tests run the workspace source.
mock.module('@pptx-editor.dev/vue', () => adapter);

plugin({
  name: 'vue-sfc',
  setup(build) {
    let n = 0;
    build.onLoad({ filter: /\.vue$/ }, ({ path }) => {
      const { descriptor, errors } = parse(readFileSync(path, 'utf8'), { filename: path });
      if (errors.length) throw errors[0];
      const script = compileScript(descriptor, {
        id: `demo-${++n}`,
        inlineTemplate: true,
        genDefaultAs: '__demo_component',
      });
      return {
        contents: `${script.content}\nexport default __demo_component;\n`,
        loader: 'ts',
      };
    });
  },
});

// A variable specifier: the type check of this package does not follow it into the demo.
const APP = '../../../examples/vue/src/App.vue';

let app: App | null = null;

/** Open the Vue demo at a window size, on the basic deck, and wait until it paints. */
export async function openDemo(options: DemoOptions = {}): Promise<Demo> {
  const { default: Root } = (await import(APP)) as { default: Component };
  return openDemoWith(
    {
      publicDir: new URL('../../../examples/vite/public/', import.meta.url),
      mount: () => {
        const host = document.createElement('div');
        document.body.append(host);
        app = createApp(Root);
        app.mount(host);
      },
      unmount: () => {
        app?.unmount();
        app = null;
      },
      run: async (step) => {
        await step();
        await nextTick();
      },
    },
    options
  );
}
