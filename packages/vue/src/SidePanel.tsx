/** @jsxImportSource vue */
// The side panel (Format, Chart Data, Diagram, Selection, Animations, and Comments): the Vue twin of `SidePanel` in the React adapter's
// `panes.tsx`, with the same markup, classes, and ARIA. The Format tab writes the selected
// element's position, size, rotation, and text fitting through engine commands.

import {
  computed,
  defineComponent,
  nextTick,
  onMounted,
  ref,
  watch,
  type CSSProperties,
  type ExtractPublicPropTypes,
  type PropType,
  type StyleValue,
  type SlotsType,
  type VNodeChild,
} from 'vue';
import {
  listenPaneRequests,
  takePaneRequest,
  SIDE_PANEL_TAB_ICONS,
  visibleSidePanelTabs,
  type SidePanelTab,
  panelHistoryKey,
  radioGroupStep,
} from '@pptx-editor.dev/core/editor';
import type { TranslationKey } from '@pptx-editor.dev/i18n';
import { Icon, ToolButton, keepFocus } from './chrome.tsx';
import { useEditorState, usePptxEditor, useTranslation } from './composables.ts';
import { PptxEditorSelectionPane } from './SelectionPane.tsx';
import { AnimationsPane } from './AnimationsPane.tsx';
import { ChartDataPane } from './ChartDataPane.tsx';
import { DiagramTextPane } from './DiagramTextPane.tsx';
import { FormatPanel } from './FormatPanel.tsx';

const sidePanelProps = {
  /** The panel's width in CSS pixels. */
  paneWidth: { type: Number, default: undefined },
  /** Closes the panel. When given, the panel shows a close button, as React's `onClose`. */
  onClose: { type: Function as PropType<() => void>, default: undefined },
  /** A class name added to the panel, as React's `className`. A host `class` also works. */
  className: { type: String, default: undefined },
  /** Inline styles for the panel, as React's `style`; a `width` here wins over `paneWidth`. */
  style: { type: Object as PropType<CSSProperties>, default: undefined },
  /** The tab that shows, controlled. Use `v-model:tab`, or listen to `tabChange`. */
  tab: { type: String as PropType<SidePanelTab>, default: undefined },
} as const;

/** A tab's packaged pane. The Comments tab has none: a review pane fills its slot. */
function packagedPane(tab: SidePanelTab, t: (key: TranslationKey) => string): VNodeChild {
  if (tab === 'format') return <FormatPanel />;
  if (tab === 'chart') return <ChartDataPane />;
  if (tab === 'diagram') return <DiagramTextPane />;
  if (tab === 'selection') return <PptxEditorSelectionPane />;
  if (tab === 'animations') return <AnimationsPane />;
  return (
    <div class="pptx-empty">
      <p class="pptx-empty__title">{t('panel.comments.emptyTitle')}</p>
      <p class="pptx-empty__body">{t('panel.comments.emptyBody')}</p>
    </div>
  );
}

/**
 * Props of `PptxEditorSidePanel`, the twin of React's `SidePanelProps`. `onClose` shows a close
 * button that calls it; `className` (or a host `class`) and `style` go on the panel, and a
 * `width` in `style` wins over `paneWidth`. @public
 */
export type SidePanelProps = ExtractPublicPropTypes<typeof sidePanelProps>;

/**
 * The side panel: Format · Chart Data · Diagram · Selection · Animations · Comments, the twin of React's `PptxEditor.SidePanel`. The host
 * shows or hides the panel by rendering it. @public
 */
export const PptxEditorSidePanel = defineComponent({
  name: 'PptxEditorSidePanel',
  inheritAttrs: false,
  props: sidePanelProps,
  // One named slot per tab (`#comments`, `#format`, ...): its content replaces the tab's
  // packaged pane, the twin of React's `panes`. The Comments tab has no packaged pane: put the
  // comments pane of `@pptx-editor.dev/pro/vue` in `#comments`.
  slots: Object as SlotsType<{ [K in SidePanelTab]?: () => VNodeChild }>,
  emits: {
    tabChange: (_tab: SidePanelTab) => true,
    'update:tab': (_tab: SidePanelTab) => true,
  },
  setup(props, { attrs, emit, slots }) {
    const { t } = useTranslation();
    const editor = usePptxEditor();
    const own = ref<SidePanelTab>(props.tab ?? 'format');
    const tab = computed<SidePanelTab>(() => props.tab ?? own.value);
    // The Chart Data and Diagram tabs show while their object is selected, or while open.
    const chartSelected = useEditorState((s) => !!s.chart);
    const diagramSelected = useEditorState((s) => !!s.diagram);
    const tabs = computed(() =>
      visibleSidePanelTabs(chartSelected.value, tab.value, diagramSelected.value)
    );
    const setTab = (next: SidePanelTab): void => {
      own.value = next;
      emit('tabChange', next);
      emit('update:tab', next);
    };
    // The toolbar's Selection pane and Format shape slots ask for a tab (`requestPane`). The
    // focus goes to that tab, so a keyboard user lands in the panel.
    const root = ref<HTMLElement | null>(null);
    // A request (a comment marker, the Selection pane control) lands the focus in the panel: on
    // the pane's current item when it marks one (`aria-current`), else on the tab. Focus that a
    // pane already took for the request (a new comment draft) stays.
    const focusTab = (pane: SidePanelTab): void =>
      void nextTick(() => {
        const body = root.value?.querySelector('.pptx-panel__body');
        if (body?.contains(body.ownerDocument.activeElement)) return;
        const current = root.value?.querySelector<HTMLElement>(
          '.pptx-panel__body [aria-current="true"]'
        );
        (current ?? root.value?.querySelector<HTMLElement>(`[data-tab="${pane}"]`))?.focus();
      });
    watch(
      editor,
      (e, _old, onCleanup) => {
        if (!e) return;
        onCleanup(
          listenPaneRequests(e, (pane, request) => {
            takePaneRequest(e);
            setTab(pane);
            if (request.focus) focusTab(pane);
          })
        );
      },
      { immediate: true }
    );
    // A host that shows a closed panel for a request mounts it after the request: take it now.
    onMounted(() => {
      const request = takePaneRequest(editor.value);
      if (!request) return;
      setTab(request.pane);
      if (request.focus) focusTab(request.pane);
    });
    return () => {
      const { class: hostClass, style: hostStyle, ...rest } = attrs;
      const style = [hostStyle as StyleValue, props.style];
      return (
        <aside
          ref={root}
          aria-label={t('panel.ariaLabel')}
          {...rest}
          class={['pptx-panel', hostClass, props.className]}
          style={
            (props.paneWidth === undefined
              ? style
              : [{ width: `${props.paneWidth}px` }, ...style]) as StyleValue
          }
          onMousedown={keepFocus}
        >
          <div class="pptx-panel__head">
            <div
              class="pptx-tabs"
              role="tablist"
              // One Tab stop: the selected tab; the arrow keys move between tabs.
              onKeydown={(e: KeyboardEvent) => {
                const next = radioGroupStep(tabs.value, tab.value, e.key);
                if (!next) return;
                e.preventDefault();
                setTab(next);
                (e.currentTarget as HTMLElement)
                  .querySelector<HTMLElement>(`[data-tab="${next}"]`)
                  ?.focus();
              }}
            >
              {tabs.value.map((id) => (
                <button
                  key={id}
                  type="button"
                  role="tab"
                  class="pptx-tabs__tab"
                  data-tab={id}
                  aria-selected={tab.value === id}
                  // The row shows icons; the selected tab also shows its label. Each tab keeps
                  // its label as its name and tooltip.
                  aria-label={t(`panel.tabs.${id}`)}
                  data-pptx-tooltip={t(`panel.tabs.${id}`)}
                  tabindex={tab.value === id ? 0 : -1}
                  onClick={() => setTab(id)}
                >
                  <Icon name={SIDE_PANEL_TAB_ICONS[id]} />
                  {tab.value === id ? (
                    <span class="pptx-tabs__label" aria-hidden="true">
                      {t(`panel.tabs.${id}`)}
                    </span>
                  ) : null}
                </button>
              ))}
            </div>
            {props.onClose ? (
              <ToolButton label={t('panel.close')} icon="close" onClick={() => props.onClose?.()} />
            ) : null}
          </div>
          <div
            class="pptx-panel__body"
            role="tabpanel"
            onKeydown={(e: KeyboardEvent) => panelHistoryKey(e, editor.value)}
          >
            {slots[tab.value]?.() ?? packagedPane(tab.value, t)}
          </div>
        </aside>
      );
    };
  },
});
