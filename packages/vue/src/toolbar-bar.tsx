/** @jsxImportSource vue */
// The bar element of the toolbar and of each ribbon tab: the groups with separators, the fixed
// block of loose host children, and the More menu that takes the groups that do not fit. The
// Vue twin of React's `PresetToolbar` and `HandBar` (`Toolbar.tsx`): the same markup, classes,
// and ARIA, with the fit rule in core (`toolbarCollapsedIds`, `observeToolbarFit`).

import {
  Fragment,
  defineComponent,
  onBeforeUnmount,
  onMounted,
  onUpdated,
  ref,
  shallowRef,
  useId,
  watch,
  type PropType,
  type VNode,
  type VNodeArrayChildren,
  type VNodeChild,
} from 'vue';
import {
  createToolbarFitMemory,
  formatSelectionHasLineEnds,
  formatTargetKey,
  observeToolbarFit,
  toolbarCollapsedIds,
} from '@pptx-editor.dev/core/editor';
import { Icon, Popover, keepFocus } from './chrome.tsx';
import { useEditorState, useTranslation } from './composables.ts';
import { useRovingToolbar } from './toolbar-measure.ts';

/** One group on the bar: built-in or host, already resolved against the overrides. @internal */
export interface BarItem {
  readonly id: string;
  readonly label: string | undefined;
  readonly content: () => VNodeArrayChildren;
  readonly priority: number;
  readonly pinned: boolean;
  readonly className: string | undefined;
  /** The host's content for the group in the More menu. */
  readonly overflowContent: (() => VNodeChild) | undefined;
  /**
   * What a built-in group shows in the More menu when it has `overflowContent`: its own
   * controls, then that content, as DOCX Editor's `hostPanelContent`. A host group's
   * `overflowContent` replaces its controls.
   */
  readonly menuContent?: (() => VNodeArrayChildren) | undefined;
}

const sep = (): VNode => <span class="pptx-toolbar__sep" aria-hidden="true" />;

/** A `preset` false bar: the host's children, wrapped, with roving focus. @internal */
export const HandBar = defineComponent({
  name: 'PptxHandBar',
  setup(_, { slots }) {
    const { t } = useTranslation();
    const root = ref<HTMLElement | null>(null);
    const roving = useRovingToolbar(root);
    return () => (
      <div
        ref={root}
        class="pptx-toolbar pptx-toolbar--wrap"
        role="toolbar"
        aria-label={t('toolbar.ariaLabel')}
        data-overflow="wrap"
        onMousedown={keepFocus}
        onFocusin={roving.onFocusin}
        onFocusout={roving.onFocusout}
        onKeydown={roving.onKeydown}
      >
        {slots.default?.()}
      </div>
    );
  },
});

/** The preset bar element, with the More menu. @internal */
export const PresetToolbar = defineComponent({
  name: 'PptxPresetToolbar',
  props: {
    items: { type: Array as PropType<readonly BarItem[]>, required: true as const },
    extras: { type: Array as PropType<readonly VNode[]>, required: true as const },
    overflow: { type: Boolean, required: true as const },
    label: { type: String, required: true as const },
  },
  setup(props) {
    const { t } = useTranslation();
    // Contextual controls change the width; refit when they appear or go.
    const context = useEditorState((s) => `${formatTargetKey(s)}|${formatSelectionHasLineEnds(s)}`);
    const status = useEditorState((s) => s.status);
    const root = ref<HTMLElement | null>(null);
    const more = ref<HTMLElement | null>(null);
    const fixed = ref<HTMLElement | null>(null);
    const panelId = useId();
    // Widths and the collapsed set are keyed by group id, so a group that comes or goes does
    // not hand its width to its neighbour.
    const memory = createToolbarFitMemory();
    const collapsed = shallowRef<readonly string[]>([]);
    const open = ref(false);
    const roving = useRovingToolbar(root);

    let stopObserving: (() => void) | null = null;
    const fit = (): void => {
      const el = root.value;
      if (!el) return;
      if (!props.overflow) {
        if (collapsed.value.length) collapsed.value = [];
        return;
      }
      const next = toolbarCollapsedIds(el, props.items, memory, more.value, fixed.value);
      if (collapsed.value.join() !== next.join()) collapsed.value = next;
    };
    // After every render, as React's layout effect: measure, then watch the bar, its parent,
    // its groups, and the fixed block for size changes. A render that changes nothing leaves
    // the collapsed set as it is, so this settles.
    const refit = (): void => {
      stopObserving?.();
      stopObserving = null;
      fit();
      if (props.overflow && root.value) stopObserving = observeToolbarFit(root.value, fit);
    };
    onMounted(refit);
    onUpdated(refit);
    watch([context, status], refit, { flush: 'post' });
    onBeforeUnmount(() => stopObserving?.());
    // The More menu closes for good when nothing is in it, so it does not reopen by itself.
    watch(
      () => collapsed.value.length,
      (n) => {
        if (!n) open.value = false;
      }
    );

    const group = (item: BarItem, index: number, inMenu: boolean): VNode => (
      <div
        key={item.id}
        class={['pptx-toolbar__group', item.className]}
        role="group"
        aria-label={item.label}
        data-group={item.id}
        data-group-index={inMenu ? undefined : index}
      >
        {/* The group's aria-label already names it: the heading is for sight only. */}
        {inMenu ? (
          <div class="pptx-toolbar__overflow-heading" aria-hidden="true">
            {item.label}
          </div>
        ) : null}
        {inMenu && item.overflowContent
          ? item.menuContent
            ? item.menuContent()
            : item.overflowContent()
          : item.content()}
      </div>
    );

    return () => {
      const items = props.items;
      const hiddenIds = new Set(props.overflow ? collapsed.value : []);
      const visible = items.map((_, i) => i).filter((i) => !hiddenIds.has(items[i]!.id));
      const hidden = items.map((_, i) => i).filter((i) => hiddenIds.has(items[i]!.id));
      return (
        <div
          ref={root}
          class={['pptx-toolbar', props.overflow ? null : 'pptx-toolbar--wrap']}
          role="toolbar"
          aria-label={props.label}
          data-overflow={props.overflow ? 'menu' : 'wrap'}
          onMousedown={keepFocus}
          onFocusin={roving.onFocusin}
          onFocusout={roving.onFocusout}
          onKeydown={roving.onKeydown}
        >
          {visible.map((index, i) => (
            <Fragment key={items[index]!.id}>
              {i > 0 ? sep() : null}
              {group(items[index]!, index, false)}
            </Fragment>
          ))}
          {props.extras.length ? (
            <>
              {visible.length ? sep() : null}
              <div ref={fixed} class="pptx-toolbar__group pptx-toolbar__fixed" data-fixed="">
                {props.extras}
              </div>
            </>
          ) : null}
          {hidden.length ? (
            <>
              {visible.length || props.extras.length ? sep() : null}
              <button
                ref={more}
                type="button"
                class="pptx-toolbar__button"
                aria-label={t('toolbar.more')}
                data-pptx-tooltip={t('toolbar.more')}
                aria-haspopup="dialog"
                aria-expanded={open.value}
                data-active={open.value ? '' : undefined}
                aria-controls={open.value ? panelId : undefined}
                onMousedown={keepFocus}
                onClick={() => (open.value = !open.value)}
                onKeydown={(event: KeyboardEvent) => {
                  // As DOCX Editor's More button: ArrowDown opens the panel at its first control.
                  if (event.key !== 'ArrowDown') return;
                  event.preventDefault();
                  open.value = true;
                }}
              >
                <Icon name="more_vert" />
              </button>
              {open.value ? (
                <Popover
                  id={panelId}
                  anchor={more.value}
                  align="end"
                  role="dialog"
                  focusFirst
                  onClose={() => (open.value = false)}
                >
                  {{
                    default: () => (
                      <div class="pptx-toolbar__overflow" onMousedown={keepFocus}>
                        {hidden.map((i) => group(items[i]!, i, true))}
                      </div>
                    ),
                  }}
                </Popover>
              ) : null}
            </>
          ) : null}
        </div>
      );
    };
  },
});
