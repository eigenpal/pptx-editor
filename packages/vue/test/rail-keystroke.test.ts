// A keystroke in a long deck makes one new rail item in both adapters: the other items keep
// their element (React) or vnode (Vue), so their names are not translated again
// (core's `createItemMemo`). A locale change still names every item again.

import { afterEach, describe, expect, test } from 'bun:test';
import { act, createElement, type FunctionComponent } from 'react';
import { cleanup, render } from '@testing-library/react';
import { createApp, defineComponent, h, nextTick, shallowRef, type App, type Component } from 'vue';
import {
  createPptxEditor,
  type ChromeTranslate,
  type PptxEditorInstance,
} from '@pptx-editor.dev/core/editor';
import { textDeck } from '../../../scripts/bench/text-deck.ts';
import { PptxEditorContext } from '../../react/src/context.ts';
import { SlideRail as RSlideRailNamespace } from '../../react/src/index.ts';
import type { SlideRailProps as RSlideRailProps } from '../../react/src/panes.tsx';
import { PptxEditorSlideRail } from '../src/index.ts';
import { providePptxEditor } from '../src/composables.ts';

const RSlideRail = RSlideRailNamespace as unknown as FunctionComponent<RSlideRailProps>;
const SLIDES = 40;

let app: App | null = null;
afterEach(() => {
  cleanup();
  app?.unmount();
  app = null;
  document.body.replaceChildren();
});

async function deck(): Promise<PptxEditorInstance> {
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await editor.load(textDeck(SLIDES))).ok).toBe(true);
  return editor;
}

/** A translation that records the slide numbers it names. */
function countingT(prefix = ''): { t: ChromeTranslate; named: number[] } {
  const named: number[] = [];
  const t: ChromeTranslate = (key, params) => {
    if (key === 'rail.slideLabel' || key === 'rail.hiddenSlideLabel') {
      named.push(Number(params?.index));
      return `${prefix}Slide ${String(params?.index)}`;
    }
    return key;
  };
  return { t, named };
}

/** Type one character at the end of the body text of slide `index`. */
function typeOn(editor: PptxEditorInstance, index: number): void {
  editor.goToSlide(index);
  const lay = editor.session!.layout(index)!;
  const body = lay.items.find((i) => i.kind === 'shape' && i.element.name === 'Body')!;
  const end = body.kind === 'shape' ? (body.text?.paragraphLengths[0] ?? 0) : 0;
  const at = { paragraph: 0, offset: end };
  editor.setSelection({ kind: 'text', nodeId: body.element.nodeId, anchor: at, focus: at });
  expect(editor.input('text', 'x').ok).toBe(true);
}

const items = () => [...document.querySelectorAll('.pptx-rail__item')];
const names = () => items().map((li) => li.querySelector('.pptx-sr-only')?.textContent);

/** The slide numbers named, less the probe of the first slide's labels. */
const itemNames = (named: number[]) => named.filter((n) => n !== 1);

describe('a keystroke renders one rail item', () => {
  test('React', async () => {
    const editor = await deck();
    editor.goToSlide(9);
    const { t, named } = countingT();
    const view = render(
      createElement(PptxEditorContext.Provider, { value: editor }, createElement(RSlideRail, { t }))
    );
    expect(items()).toHaveLength(SLIDES);
    const before = items();
    named.length = 0;
    act(() => typeOn(editor, 9));
    // Slide 10 changed. No other item was named again, and every item element stayed.
    expect(itemNames(named)).toEqual([10]);
    expect(items()).toEqual(before);
    // A new translation names every item again.
    const other = countingT('Folie ');
    view.rerender(
      createElement(
        PptxEditorContext.Provider,
        { value: editor },
        createElement(RSlideRail, { t: other.t })
      )
    );
    expect(names().every((n) => n?.startsWith('Folie '))).toBe(true);
    editor.destroy();
  });

  test('Vue', async () => {
    const editor = await deck();
    editor.goToSlide(9);
    const { t, named } = countingT();
    const translate = shallowRef<ChromeTranslate>(t);
    const host = document.createElement('div');
    document.body.append(host);
    app = createApp(
      defineComponent({
        setup() {
          providePptxEditor(shallowRef(editor));
          return () => h(PptxEditorSlideRail, { t: translate.value });
        },
      }) as Component
    );
    // Each rail item's instance, and the items that render again. A patch that compares the
    // props of an item, even one that then skips its render, gives the instance a new vnode;
    // with the memo mark, Vue skips the item before that (`rail-memo.ts`).
    type Instance = { vnode: unknown; props: { slide: { index: number } } };
    const instances: Instance[] = [];
    const updated: number[] = [];
    app.mixin({
      mounted(this: { $: Instance; $options: { name?: string } }) {
        if (this.$options.name === 'PptxRailItem') instances.push(this.$);
      },
      updated(this: { $: Instance; $options: { name?: string } }) {
        if (this.$options.name === 'PptxRailItem') updated.push(this.$.props.slide.index);
      },
    });
    app.mount(host);
    await nextTick();
    expect(items()).toHaveLength(SLIDES);
    expect(instances).toHaveLength(SLIDES);
    const before = items();
    const vnodes = instances.map((i) => i.vnode);
    named.length = 0;
    typeOn(editor, 9);
    await nextTick();
    expect(itemNames(named)).toEqual([10]);
    expect(items()).toEqual(before);
    // Only slide 10's item rendered again; Vue did not patch any other item.
    expect(updated).toEqual([9]);
    const patched = instances.filter((inst, i) => inst.vnode !== vnodes[i]);
    expect(patched.map((i) => i.props.slide.index)).toEqual([9]);
    translate.value = countingT('Folie ').t;
    await nextTick();
    expect(names().every((n) => n?.startsWith('Folie '))).toBe(true);
    editor.destroy();
  });
});
