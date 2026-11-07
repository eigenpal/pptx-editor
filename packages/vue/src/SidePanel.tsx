/** @jsxImportSource vue */
// The side panel (Format and Comments): the Vue twin of `SidePanel` in the React adapter's
// `panes.tsx`, with the same markup, classes, and ARIA. The Format tab writes the selected
// element's position, size, rotation, and text fitting through engine commands.

import {
  computed,
  defineComponent,
  ref,
  watch,
  type CSSProperties,
  type ExtractPublicPropTypes,
  type PropType,
  type StyleValue,
} from 'vue';
import {
  elementKindLabelKey,
  elementRotates,
  formatPanelSize,
  listenPaneRequests,
  TABLE_OPTION_ROWS,
  SIDE_PANEL_TABS,
  type SidePanelTab,
  panelHistoryKey,
  platformShortcuts,
  radioGroupStep,
} from '@pptx-editor.dev/core/editor';
import type { TranslationKey } from '@pptx-editor.dev/i18n';
import { ToolButton, keepFocus } from './chrome.tsx';
import { useEditorDialog } from './popups.ts';
import { useEditorState, usePptxEditor, useTranslation } from './composables.ts';
import { PptxEditorSelectionPane } from './SelectionPane.tsx';

const EMU_PER_INCH = 914_400;
const inches = (emu: number): string => (emu / EMU_PER_INCH).toFixed(2);

/** A number field that commits on Enter or blur and shows the document value otherwise. */
const PanelNumberField = defineComponent({
  name: 'PptxPanelNumberField',
  props: {
    label: { type: String, required: true },
    value: { type: String, required: true },
    unit: { type: String, required: true },
    disabled: { type: Boolean, required: true },
    onCommit: { type: Function as PropType<(value: number) => void>, required: true },
  },
  setup(props) {
    const draft = ref<string | null>(null);
    const commit = (): void => {
      // An empty field keeps the value: `Number('')` is 0, which would move the shape.
      const d = draft.value;
      const n = d === null || d.trim() === '' ? NaN : Number(d.replace(',', '.'));
      if (Number.isFinite(n) && draft.value !== props.value) props.onCommit(n);
      draft.value = null;
    };
    return () => (
      <label class="pptx-field">
        {props.label}
        <span class="pptx-field__control">
          <input
            inputmode="decimal"
            aria-label={props.label}
            value={draft.value ?? props.value}
            disabled={props.disabled}
            // Uncommitted typing keeps Ctrl/Cmd+Z for the field's own undo (`panelHistoryKey`).
            data-draft={draft.value !== null}
            onInput={(e: Event) => {
              draft.value = (e.target as HTMLInputElement).value;
            }}
            onBlur={commit}
            onKeydown={(e: KeyboardEvent) => {
              // Enter commits and keeps the focus here, so the next key still reaches the panel.
              if (e.key === 'Enter') commit();
              if (e.key === 'Escape') draft.value = null;
            }}
          />
          <span class="pptx-field__unit">{props.unit}</span>
        </span>
      </label>
    );
  },
});

/** Position, size, rotation, and text fitting of the selected element. */
const FormatTab = defineComponent({
  name: 'PptxFormatTab',
  setup() {
    const { t } = useTranslation();
    const editor = usePptxEditor();
    const geometry = useEditorState((s) => s.geometry);
    const table = useEditorState((s) => s.table);
    const readOnly = useEditorState((s) => s.mode === 'view');
    const selected = useEditorState((s) =>
      s.selection.kind === 'shapes' ? s.selection.nodeIds.length : 0
    );
    const background = useEditorDialog('background');
    return () => {
      const g = geometry.value;
      if (!g)
        return selected.value > 1 ? (
          <div class="pptx-empty">
            <p class="pptx-empty__title">
              {t('panel.format.multipleTitle', { count: selected.value })}
            </p>
            <p class="pptx-empty__body">{platformShortcuts(t('panel.format.multipleBody'))}</p>
          </div>
        ) : (
          [
            <div class="pptx-panel__section">
              <h3 class="pptx-panel__title">{t('panel.format.slide')}</h3>
              {background.isAvailable.value ? (
                <button
                  type="button"
                  class="pptx-button"
                  disabled={readOnly.value}
                  onClick={(e: MouseEvent) => background.open(e.currentTarget as HTMLElement)}
                >
                  {t('menu.background')}
                </button>
              ) : null}
            </div>,
            <div class="pptx-empty">
              <p class="pptx-empty__title">{t('panel.format.emptyTitle')}</p>
              <p class="pptx-empty__body">{t('panel.format.emptyBody')}</p>
            </div>,
          ]
        );
      const move = (patch: Partial<Record<'x' | 'y' | 'cx' | 'cy' | 'rot', number>>): void => {
        editor.value?.exec({
          type: 'setTransform',
          nodeId: g.nodeId,
          x: patch.x ?? g.x,
          y: patch.y ?? g.y,
          cx: Math.max(0, patch.cx ?? g.cx),
          cy: Math.max(0, patch.cy ?? g.cy),
          rot: patch.rot ?? g.rot,
        });
      };
      const inch = (v: number): number => Math.round(v * EMU_PER_INCH);
      const unit = t('panel.format.inchUnit');
      const tableSection = () => {
        const tb = table.value;
        if (!tb) return null;
        return (
          <div class="pptx-panel__section">
            <h3 class="pptx-panel__title">{t('panel.format.table')}</h3>
            <div class="pptx-panel__grid">
              <PanelNumberField
                label={t('panel.format.rowHeight')}
                value={inches(tb.rowHeight)}
                unit={unit}
                disabled={readOnly.value}
                onCommit={(v: number) =>
                  editor.value?.exec({ type: 'setRowHeight', height: inch(v) })
                }
              />
              <PanelNumberField
                label={t('panel.format.columnWidth')}
                value={inches(tb.columnWidth)}
                unit={unit}
                disabled={readOnly.value}
                onCommit={(v: number) =>
                  editor.value?.exec({ type: 'setColumnWidth', width: inch(v) })
                }
              />
            </div>
            <div class="pptx-radios" role="group" aria-label={t('panel.format.styleOptions')}>
              {TABLE_OPTION_ROWS.map(({ option, labelKey }) => (
                <label key={option} class="pptx-radio">
                  <input
                    type="checkbox"
                    checked={tb.options.includes(option)}
                    disabled={readOnly.value}
                    onChange={() => editor.value?.exec({ type: 'toggleTableStyleOption', option })}
                  />
                  {t(labelKey)}
                </label>
              ))}
            </div>
          </div>
        );
      };
      const field = (key: 'x' | 'y' | 'cx' | 'cy', labelKey: TranslationKey) => (
        <PanelNumberField
          label={t(labelKey)}
          value={inches(g[key])}
          unit={unit}
          disabled={readOnly.value}
          // A locked aspect ratio sizes both sides from one.
          onCommit={(v: number) =>
            move(
              key === 'cx' || key === 'cy' ? formatPanelSize(g, key, inch(v)) : { [key]: inch(v) }
            )
          }
        />
      );
      return [
        <div class="pptx-panel__section">
          <h3 class="pptx-panel__title">{t('panel.format.element')}</h3>
          <span class="pptx-chip">{t(elementKindLabelKey(g, table.value))}</span>
        </div>,
        <div class="pptx-panel__section">
          <h3 class="pptx-panel__title">{t('panel.format.size')}</h3>
          <div class="pptx-panel__grid">
            {field('cx', 'panel.format.width')}
            {field('cy', 'panel.format.height')}
          </div>
          <label class="pptx-radio">
            <input
              type="checkbox"
              checked={g.lockAspectRatio === true}
              disabled={readOnly.value}
              onChange={(e: Event) =>
                editor.value?.exec({
                  type: 'setLockAspectRatio',
                  lock: (e.currentTarget as HTMLInputElement).checked,
                })
              }
            />
            {t('imageProperties.lockAspect')}
          </label>
        </div>,
        <div class="pptx-panel__section">
          <h3 class="pptx-panel__title">{t('panel.format.position')}</h3>
          <div class="pptx-panel__grid">
            {field('x', 'panel.format.x')}
            {field('y', 'panel.format.y')}
          </div>
        </div>,
        elementRotates(g, table.value) ? (
          <div class="pptx-panel__section">
            <h3 class="pptx-panel__title">{t('panel.format.rotation')}</h3>
            <div class="pptx-panel__grid">
              <PanelNumberField
                label={t('panel.format.angle')}
                value={String(Math.round(g.rot * 10) / 10)}
                unit="°"
                disabled={readOnly.value}
                onCommit={(v: number) => move({ rot: ((v % 360) + 360) % 360 })}
              />
            </div>
          </div>
        ) : null,
        tableSection(),
        g.textFitting ? (
          <div class="pptx-panel__section">
            <h3 class="pptx-panel__title">{t('panel.format.textFitting')}</h3>
            <div class="pptx-radios" role="radiogroup" aria-label={t('panel.format.textFitting')}>
              {(['none', 'shrink', 'shape'] as const).map((mode) => (
                <label key={mode} class="pptx-radio">
                  <input
                    type="radio"
                    name="pptx-text-fitting"
                    checked={g.textFitting === mode}
                    disabled={readOnly.value}
                    onChange={() => editor.value?.exec({ type: 'setAutofit', mode })}
                  />
                  {t(`panel.format.fitting.${mode}`)}
                </label>
              ))}
            </div>
          </div>
        ) : null,
      ];
    };
  },
});

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

/**
 * Props of `PptxEditorSidePanel`, the twin of React's `SidePanelProps`. `onClose` shows a close
 * button that calls it; `className` (or a host `class`) and `style` go on the panel, and a
 * `width` in `style` wins over `paneWidth`. @public
 */
export type SidePanelProps = ExtractPublicPropTypes<typeof sidePanelProps>;

/**
 * The side panel: Format · Selection · Comments, the twin of React's `PptxEditor.SidePanel`. The host
 * shows or hides the panel by rendering it. @public
 */
export const PptxEditorSidePanel = defineComponent({
  name: 'PptxEditorSidePanel',
  inheritAttrs: false,
  props: sidePanelProps,
  emits: {
    tabChange: (_tab: SidePanelTab) => true,
    'update:tab': (_tab: SidePanelTab) => true,
  },
  setup(props, { attrs, emit }) {
    const { t } = useTranslation();
    const editor = usePptxEditor();
    const own = ref<SidePanelTab>(props.tab ?? 'format');
    const tab = computed<SidePanelTab>(() => props.tab ?? own.value);
    const setTab = (next: SidePanelTab): void => {
      own.value = next;
      emit('tabChange', next);
      emit('update:tab', next);
    };
    // The toolbar's Selection pane and Format pane slots ask for a tab (`requestPane`).
    watch(
      editor,
      (e, _old, onCleanup) => {
        if (e) onCleanup(listenPaneRequests(e, setTab));
      },
      { immediate: true }
    );
    return () => {
      const { class: hostClass, style: hostStyle, ...rest } = attrs;
      const style = [hostStyle as StyleValue, props.style];
      return (
        <aside
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
                const next = radioGroupStep(SIDE_PANEL_TABS, tab.value, e.key);
                if (!next) return;
                e.preventDefault();
                setTab(next);
                (e.currentTarget as HTMLElement)
                  .querySelector<HTMLElement>(`[data-tab="${next}"]`)
                  ?.focus();
              }}
            >
              {SIDE_PANEL_TABS.map((id) => (
                <button
                  key={id}
                  type="button"
                  role="tab"
                  class="pptx-tabs__tab"
                  data-tab={id}
                  aria-selected={tab.value === id}
                  tabindex={tab.value === id ? 0 : -1}
                  onClick={() => setTab(id)}
                >
                  {t(`panel.tabs.${id}`)}
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
            {tab.value === 'format' ? <FormatTab /> : null}
            {tab.value === 'selection' ? <PptxEditorSelectionPane /> : null}
            {tab.value === 'comments' ? (
              <div class="pptx-empty">
                <p class="pptx-empty__title">{t('panel.comments.emptyTitle')}</p>
                <p class="pptx-empty__body">{t('panel.comments.emptyBody')}</p>
              </div>
            ) : null}
          </div>
        </aside>
      );
    };
  },
});
