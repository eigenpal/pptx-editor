/** @jsxImportSource vue */
// The find bar: the Vue twin of the React adapter's `FindBar.tsx`. Find and replace over the
// engine's `editor.find`. Ctrl/Cmd+F opens it to find and Ctrl/Cmd+H to replace. Enter and
// Shift+Enter step through the matches; Escape closes it and returns the focus.
//
// The toolbar's `find` group and a host's menu open it with `requestFindBar`: an event that
// names the editor, so a bar answers only its own editor. The requests, keys, counter, actions,
// and focus rules are core's (`overlays.ts`), so the two adapters cannot drift.

import {
  defineComponent,
  inject,
  onScopeDispose,
  provide,
  ref,
  shallowRef,
  watch,
  type InjectionKey,
  type Ref,
  type ExtractPublicPropTypes,
} from 'vue';
import {
  dispatchFindRequest,
  findAsYouType,
  findBarOpener,
  findBarView,
  findFieldTakesFocus,
  focusAfterFindBar,
  listenFindRequests,
  onFindShown,
  onNewDocument,
  reserveFindBarSpace,
  runFindBarAction,
  type FindBarAction,
  type FindOptions,
  type PptxEditorInstance,
} from '@pptx-editor.dev/core/editor';
import { ToolButton } from './chrome.tsx';
import { useEditorState, usePptxEditor } from './composables.ts';
import { chromeTranslateProp, useChromeTranslateScope } from './headless-hooks.ts';
import { usePopupEntry } from './popups.ts';

/** What a replacement find bar receives, as React's `FindPopupProps`. @public */
export interface FindPopupProps {
  /** Opened to replace (Ctrl/Cmd+H), not only to find. */
  readonly replace: boolean;
  /** Close the bar. Run `editor.find` and `editor.replace` for the search itself. */
  readonly onClose: () => void;
}

/** Open the find bar of `editor`: to find, or with the replace row when `replace`. @public */
export function requestFindBar(editor: PptxEditorInstance | null, replace = false): void {
  dispatchFindRequest(editor, replace);
}

/** How many find bars a root has mounted, so Find shows as disabled while there is none. */
const findBarsKey: InjectionKey<Ref<number>> = Symbol('pptxFindBars');

/** Provide the find bar count of one root and return it. @internal */
export function provideFindBars(): Readonly<Ref<number>> {
  const count = ref(0);
  provide(findBarsKey, count);
  return count;
}

/** Count this find bar in its root while it is mounted. */
function useRegisterFindBar(): void {
  const count = inject(findBarsKey, null);
  if (!count) return;
  count.value += 1;
  onScopeDispose(() => {
    count.value -= 1;
  });
}

const pptxEditorFindBarProps = {
  ...chromeTranslateProp,
  /** A class name added to the find bar, as React's `className`. A host `class` also works. */
  className: { type: String, default: undefined },
} as const;

/** The props of `PptxEditorFindBar`, the same name as React's `FindBarProps`. @public */
export type FindBarProps = ExtractPublicPropTypes<typeof pptxEditorFindBarProps>;

/**
 * The find and replace bar, the twin of React's `PptxEditor.FindBar`. Render one inside the
 * editor's viewport, over the canvas. A host `class` goes on the bar. @public
 */
export const PptxEditorFindBar = defineComponent({
  name: 'PptxEditorFindBar',
  inheritAttrs: false,
  props: pptxEditorFindBarProps,
  setup(props, { attrs }) {
    const editorRef = usePptxEditor();
    const { t } = useChromeTranslateScope(() => props.t);
    // Find rows and controls are on only while a find bar can answer them.
    useRegisterFindBar();
    // `popups.find`: a renderer replaces the packaged bar here; `false` turns find off.
    const entry = usePopupEntry('find');
    const root = ref<HTMLDivElement | null>(null);
    // Finds the editor scope when a host renderer replaces the packaged bar.
    const anchor = ref<HTMLSpanElement | null>(null);
    const findInput = ref<HTMLInputElement | null>(null);
    let returnTo: HTMLElement | null = null;
    const open = ref(false);
    const replaceShown = ref(false);
    const focusCount = ref(0);
    const query = ref('');
    const replacement = ref('');
    const matchCase = ref(false);
    const wholeWord = ref(false);
    const status = shallowRef<string | null>(null);
    // The counter reads the deck and the selection: render again when either moves.
    const revision = useEditorState((s) => s.revision);
    const selection = useEditorState((s) => s.selection);
    const currentSlide = useEditorState((s) => s.currentSlide);
    // A step inside the speaker notes moves neither: find tells when it shows another match.
    const shownCount = ref(0);
    watch(
      editorRef,
      (editor, _old, onCleanup) => {
        if (editor)
          onCleanup(
            onFindShown(editor.find, () => {
              shownCount.value += 1;
            })
          );
      },
      { immediate: true }
    );
    const editable = useEditorState((s) => s.status === 'ready' && s.mode === 'edit');
    const options = (): FindOptions => ({ matchCase: matchCase.value, wholeWord: wholeWord.value });

    const show = (replace: boolean): void => {
      if (!open.value) {
        returnTo = findBarOpener();
        // A fresh Find shows only the find row; Replace opens both.
        replaceShown.value = replace;
      } else if (replace) replaceShown.value = true;
      open.value = true;
      focusCount.value += 1;
    };

    watch(
      focusCount,
      () => {
        findInput.value?.focus();
        findInput.value?.select();
      },
      { flush: 'post' }
    );

    watch(
      () => [editorRef.value, entry() === false] as const,
      ([editor, off], _old, onCleanup) => {
        if (off || !editor) return;
        // The request for this editor, and Ctrl/Cmd+F and Ctrl/Cmd+H while the focus is in this
        // editor (or nowhere).
        onCleanup(listenFindRequests(editor, () => root.value ?? anchor.value, show));
      },
      { immediate: true }
    );

    // The open bar keeps its height free above the slide (core's `reserveFindBarSpace`).
    watch(
      root,
      (bar, _old, onCleanup) => {
        if (bar) onCleanup(reserveFindBarSpace(bar));
      },
      { flush: 'post' }
    );

    // Another document: the bar closes and forgets the search, which names the old deck's text.
    watch(
      editorRef,
      (editor, _old, onCleanup) => {
        if (!editor) return;
        onCleanup(
          onNewDocument(editor, () => {
            open.value = false;
            query.value = '';
            replacement.value = '';
            status.value = null;
          })
        );
      },
      { immediate: true }
    );

    const run = (action: FindBarAction): void => {
      const editor = editorRef.value;
      if (!editor || !query.value) return;
      status.value = runFindBarAction(t, editor, action, query.value, replacement.value, options());
      // The buttons turn off with no match left: focus stays in the bar, in the find field.
      if (findFieldTakesFocus(editor, action, query.value, options())) findInput.value?.focus();
    };

    const close = (): void => {
      open.value = false;
      status.value = null;
      const back = returnTo;
      returnTo = null;
      focusAfterFindBar(editorRef.value, root.value, root.value ?? anchor.value, back);
    };

    const onBarKey = (e: KeyboardEvent): void => {
      if (e.key !== 'Escape') return;
      e.preventDefault();
      e.stopPropagation();
      close();
    };

    return () => {
      const current = entry();
      if (current === false) return null;
      // Render again when the deck or the selection moves.
      void [revision.value, selection.value, currentSlide.value, shownCount.value];
      if (typeof current === 'function')
        return [
          <span ref={anchor} hidden />,
          open.value ? current({ replace: replaceShown.value, onClose: close }) : null,
        ];
      const { matches, counter, inNotes } = findBarView(
        t,
        editorRef.value,
        open.value,
        query.value,
        options()
      );
      return (
        <div
          {...attrs}
          ref={root}
          class={['pptx-find', attrs.class, props.className]}
          role="search"
          aria-label={t('find.label')}
          hidden={!open.value}
          onKeydown={onBarKey}
        >
          {open.value
            ? [
                <div class="pptx-find__row">
                  <ToolButton
                    label={t('find.showReplace')}
                    icon={replaceShown.value ? 'keyboard_arrow_up' : 'keyboard_arrow_down'}
                    pressed={replaceShown.value}
                    onClick={() => {
                      replaceShown.value = !replaceShown.value;
                    }}
                  />
                  <input
                    ref={findInput}
                    class="pptx-find__input"
                    type="text"
                    aria-label={t('find.find')}
                    placeholder={t('find.findPlaceholder')}
                    value={query.value}
                    onInput={(e: Event) => {
                      query.value = (e.target as HTMLInputElement).value;
                      status.value = null;
                      findAsYouType(editorRef.value, query.value, options());
                    }}
                    onKeydown={(e: KeyboardEvent) => {
                      if (e.key !== 'Enter') return;
                      e.preventDefault();
                      run(e.shiftKey ? 'previous' : 'next');
                    }}
                  />
                  <span class="pptx-find__count" aria-live="polite">
                    {counter}
                  </span>
                  <ToolButton
                    label={t('find.previous')}
                    icon="keyboard_arrow_up"
                    disabled={!matches.length}
                    onClick={() => run('previous')}
                  />
                  <ToolButton
                    label={t('find.next')}
                    icon="keyboard_arrow_down"
                    disabled={!matches.length}
                    onClick={() => run('next')}
                  />
                  <ToolButton label={t('find.close')} icon="close" onClick={close} />
                </div>,
                replaceShown.value ? (
                  <div class="pptx-find__row pptx-find__row--replace">
                    <input
                      class="pptx-find__input"
                      type="text"
                      aria-label={t('find.replace')}
                      placeholder={t('find.replacePlaceholder')}
                      value={replacement.value}
                      disabled={!editable.value}
                      onInput={(e: Event) => {
                        replacement.value = (e.target as HTMLInputElement).value;
                      }}
                      onKeydown={(e: KeyboardEvent) => {
                        if (e.key !== 'Enter') return;
                        e.preventDefault();
                        run('replace');
                      }}
                    />
                    <button
                      type="button"
                      class="pptx-button"
                      disabled={!editable.value || !matches.length}
                      onClick={() => run('replace')}
                    >
                      {t('find.replaceOne')}
                    </button>
                    <button
                      type="button"
                      class="pptx-button"
                      disabled={!editable.value || !matches.length}
                      onClick={() => run('replaceAll')}
                    >
                      {t('find.replaceAll')}
                    </button>
                  </div>
                ) : null,
                <div class="pptx-find__options">
                  <label class="pptx-find__option">
                    <input
                      type="checkbox"
                      checked={matchCase.value}
                      onChange={(e: Event) => {
                        matchCase.value = (e.target as HTMLInputElement).checked;
                      }}
                    />
                    {t('find.matchCase')}
                  </label>
                  <label class="pptx-find__option">
                    <input
                      type="checkbox"
                      checked={wholeWord.value}
                      onChange={(e: Event) => {
                        wholeWord.value = (e.target as HTMLInputElement).checked;
                      }}
                    />
                    {t('find.wholeWord')}
                  </label>
                </div>,
                inNotes ? <p class="pptx-find__hint">{t('find.inNotes')}</p> : null,
                status.value ? (
                  <p class="pptx-find__status" role="status">
                    {status.value}
                  </p>
                ) : null,
              ]
            : null}
        </div>
      );
    };
  },
});
