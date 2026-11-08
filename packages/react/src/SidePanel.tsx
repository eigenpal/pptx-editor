/** @jsxImportSource react */
// The side panel: Format, Chart Data, Diagram, Selection (the selection pane), Animations, and Comments. The Format
// tab writes the selected element's position, size, rotation, and text fitting through engine
// commands.
// The Vue twin is `packages/vue/src/SidePanel.tsx`.

import { useEffect, useRef, useState, type CSSProperties, type ReactNode } from 'react';
import {
  SIDE_PANEL_TAB_ICONS,
  TABLE_OPTION_ROWS,
  elementKindLabelKey,
  elementRotates,
  formatPanelSize,
  listenPaneRequests,
  takePaneRequest,
  visibleSidePanelTabs,
  panelHistoryKey,
  platformShortcuts,
  radioGroupStep,
  type SidePanelTab,
  formatPanelKeys,
} from '@pptx-editor.dev/core/editor';
import type { TranslationKey } from '@pptx-editor.dev/i18n';
import { Icon, ToolButton, keepFocus } from './chrome.tsx';
import { useEditorDialog } from './popups.tsx';
import { useEditorState, usePptxEditor, useTranslation } from './context.ts';
import { paneStyle } from './pane-state.ts';
import { SelectionPane } from './SelectionPane.tsx';
import { AnimationsPane } from './AnimationsPane.tsx';
import { ChartDataPane } from './ChartDataPane.tsx';
import { DiagramTextPane } from './DiagramTextPane.tsx';

const EMU_PER_INCH = 914_400;
const inches = (emu: number): string => (emu / EMU_PER_INCH).toFixed(2);

/** A number field that commits on Enter or blur and shows the document value otherwise. */
function NumberField({
  label,
  value,
  unit,
  disabled,
  onCommit,
}: {
  label: string;
  value: string;
  unit: string;
  disabled: boolean;
  onCommit: (value: number) => void;
}) {
  const [draft, setDraft] = useState<string | null>(null);
  const commit = (): void => {
    // An empty field keeps the value: `Number('')` is 0, which would move the shape.
    const n = draft === null || draft.trim() === '' ? NaN : Number(draft.replace(',', '.'));
    if (Number.isFinite(n) && draft !== value) onCommit(n);
    setDraft(null);
  };
  return (
    <label className="pptx-field">
      {label}
      <span className="pptx-field__control">
        <input
          inputMode="decimal"
          aria-label={label}
          value={draft ?? value}
          disabled={disabled}
          // Uncommitted typing keeps Ctrl/Cmd+Z for the field's own undo (`panelHistoryKey`).
          data-draft={draft !== null}
          onChange={(e) => setDraft(e.target.value)}
          onBlur={commit}
          onKeyDown={(e) => {
            // Enter commits and keeps the focus here, so the next key still reaches the panel.
            if (e.key === 'Enter') commit();
            if (e.key === 'Escape') setDraft(null);
          }}
        />
        <span className="pptx-field__unit">{unit}</span>
      </span>
    </label>
  );
}

/** Position, size, rotation, and text fitting of the selected element. */
function FormatTab() {
  const { t } = useTranslation();
  const editor = usePptxEditor();
  const geometry = useEditorState((s) => s.geometry);
  const table = useEditorState((s) => s.table);
  const readOnly = useEditorState((s) => s.mode === 'view');
  // A layout or a master in the slide master view names itself in the heading.
  const panelTitle = useEditorState((s) => formatPanelKeys(s).title);
  const panelBody = useEditorState((s) => formatPanelKeys(s).body);
  const selected = useEditorState((s) =>
    s.selection.kind === 'shapes' ? s.selection.nodeIds.length : 0
  );
  const background = useEditorDialog('background');
  if (!geometry)
    return selected > 1 ? (
      <div className="pptx-empty">
        <p className="pptx-empty__title">{t('panel.format.multipleTitle', { count: selected })}</p>
        <p className="pptx-empty__body">{platformShortcuts(t('panel.format.multipleBody'))}</p>
      </div>
    ) : (
      <>
        <div className="pptx-panel__section">
          <h3 className="pptx-panel__title">{t(panelTitle)}</h3>
          {background.isAvailable ? (
            <button
              type="button"
              className="pptx-button"
              disabled={readOnly}
              onClick={(e) => background.open(e.currentTarget)}
            >
              {t('menu.background')}
            </button>
          ) : null}
        </div>
        <div className="pptx-empty">
          <p className="pptx-empty__title">{t('panel.format.emptyTitle')}</p>
          <p className="pptx-empty__body">{t(panelBody)}</p>
        </div>
      </>
    );
  const g = geometry;
  const move = (patch: Partial<Record<'x' | 'y' | 'cx' | 'cy' | 'rot', number>>): void => {
    editor?.exec({
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
  const field = (key: 'x' | 'y' | 'cx' | 'cy', labelKey: TranslationKey) => (
    <NumberField
      label={t(labelKey)}
      value={inches(g[key])}
      unit={unit}
      disabled={readOnly}
      // A locked aspect ratio sizes both sides from one.
      onCommit={(v) =>
        move(key === 'cx' || key === 'cy' ? formatPanelSize(g, key, inch(v)) : { [key]: inch(v) })
      }
    />
  );
  return (
    <>
      <div className="pptx-panel__section">
        <h3 className="pptx-panel__title">{t('panel.format.element')}</h3>
        <span className="pptx-chip">{t(elementKindLabelKey(g, table))}</span>
      </div>
      <div className="pptx-panel__section">
        <h3 className="pptx-panel__title">{t('panel.format.size')}</h3>
        <div className="pptx-panel__grid">
          {field('cx', 'panel.format.width')}
          {field('cy', 'panel.format.height')}
        </div>
        <label className="pptx-radio">
          <input
            type="checkbox"
            checked={g.lockAspectRatio === true}
            disabled={readOnly}
            onChange={(e) =>
              editor?.exec({ type: 'setLockAspectRatio', lock: e.currentTarget.checked })
            }
          />
          {t('imageProperties.lockAspect')}
        </label>
      </div>
      <div className="pptx-panel__section">
        <h3 className="pptx-panel__title">{t('panel.format.position')}</h3>
        <div className="pptx-panel__grid">
          {field('x', 'panel.format.x')}
          {field('y', 'panel.format.y')}
        </div>
      </div>
      {elementRotates(g, table) ? (
        <div className="pptx-panel__section">
          <h3 className="pptx-panel__title">{t('panel.format.rotation')}</h3>
          <div className="pptx-panel__grid">
            <NumberField
              label={t('panel.format.angle')}
              value={String(Math.round(g.rot * 10) / 10)}
              unit="°"
              disabled={readOnly}
              onCommit={(v) => move({ rot: ((v % 360) + 360) % 360 })}
            />
          </div>
        </div>
      ) : null}
      {table ? (
        <div className="pptx-panel__section">
          <h3 className="pptx-panel__title">{t('panel.format.table')}</h3>
          <div className="pptx-panel__grid">
            <NumberField
              label={t('panel.format.rowHeight')}
              value={inches(table.rowHeight)}
              unit={unit}
              disabled={readOnly}
              onCommit={(v) => editor?.exec({ type: 'setRowHeight', height: inch(v) })}
            />
            <NumberField
              label={t('panel.format.columnWidth')}
              value={inches(table.columnWidth)}
              unit={unit}
              disabled={readOnly}
              onCommit={(v) => editor?.exec({ type: 'setColumnWidth', width: inch(v) })}
            />
          </div>
          <div className="pptx-radios" role="group" aria-label={t('panel.format.styleOptions')}>
            {TABLE_OPTION_ROWS.map(({ option, labelKey }) => (
              <label key={option} className="pptx-radio">
                <input
                  type="checkbox"
                  checked={table.options.includes(option)}
                  disabled={readOnly}
                  onChange={() => editor?.exec({ type: 'toggleTableStyleOption', option })}
                />
                {t(labelKey)}
              </label>
            ))}
          </div>
        </div>
      ) : null}
      {g.textFitting ? (
        <div className="pptx-panel__section">
          <h3 className="pptx-panel__title">{t('panel.format.textFitting')}</h3>
          <div className="pptx-radios" role="radiogroup" aria-label={t('panel.format.textFitting')}>
            {(['none', 'shrink', 'shape'] as const).map((mode) => (
              <label key={mode} className="pptx-radio">
                <input
                  type="radio"
                  name="pptx-text-fitting"
                  checked={g.textFitting === mode}
                  disabled={readOnly}
                  onChange={() => editor?.exec({ type: 'setAutofit', mode })}
                />
                {t(`panel.format.fitting.${mode}`)}
              </label>
            ))}
          </div>
        </div>
      ) : null}
    </>
  );
}

/** A tab's packaged pane. The Comments tab has none: a review pane fills it (`panes`). */
function PackagedPane({ tab }: { readonly tab: SidePanelTab }) {
  const { t } = useTranslation();
  if (tab === 'format') return <FormatTab />;
  if (tab === 'chart') return <ChartDataPane />;
  if (tab === 'diagram') return <DiagramTextPane />;
  if (tab === 'selection') return <SelectionPane />;
  if (tab === 'animations') return <AnimationsPane />;
  return (
    <div className="pptx-empty">
      <p className="pptx-empty__title">{t('panel.comments.emptyTitle')}</p>
      <p className="pptx-empty__body">{t('panel.comments.emptyBody')}</p>
    </div>
  );
}

/** Props of `SidePanel`. The host shows or hides the panel by rendering it. @public */
export interface SidePanelProps {
  /** Shows a close button that calls it. */
  readonly onClose?: () => void;
  /** A class name added to the panel. */
  readonly className?: string;
  /** The panel's width in CSS pixels. */
  readonly paneWidth?: number;
  /** Inline styles for the panel; a `width` here wins over `paneWidth`. */
  readonly style?: CSSProperties;
  /** The tab that shows. Set it with `onTabChange` to control the tab; omit it to start on Format. */
  readonly tab?: SidePanelTab;
  /** Called with the tab the user picks. */
  readonly onTabChange?: (tab: SidePanelTab) => void;
  /**
   * Content that replaces a tab's packaged pane, by tab. The **Comments** tab has no packaged
   * pane: pass the comments pane of `@pptx-editor.dev/pro/react` as `panes.comments`. Without
   * it, the tab says that comments are kept when you save.
   */
  readonly panes?: Partial<Record<SidePanelTab, ReactNode>>;
}

/** The side panel: Format · Chart Data · Diagram · Selection · Animations · Comments. `panes` fills a tab. @public */
export function SidePanel({
  onClose,
  className,
  paneWidth,
  style,
  tab: shown,
  onTabChange,
  panes,
}: SidePanelProps) {
  const { t } = useTranslation();
  const editor = usePptxEditor();
  const [own, setOwn] = useState<SidePanelTab>(shown ?? 'format');
  const tab = shown ?? own;
  // The Chart Data and Diagram tabs show while their object is selected, or while open.
  const chartSelected = useEditorState((s) => !!s.chart);
  const diagramSelected = useEditorState((s) => !!s.diagram);
  const tabs = visibleSidePanelTabs(chartSelected, tab, diagramSelected);
  const setTab = (next: SidePanelTab): void => {
    setOwn(next);
    onTabChange?.(next);
  };
  // The toolbar's Selection pane and Format shape slots ask for a tab (`requestPane`). The
  // focus goes to that tab, so a keyboard user lands in the panel.
  const latest = useRef(setTab);
  latest.current = setTab;
  const root = useRef<HTMLElement>(null);
  const [asked, setAsked] = useState<{ pane: SidePanelTab } | null>(null);
  useEffect(
    () =>
      editor
        ? listenPaneRequests(editor, (pane) => {
            takePaneRequest(editor);
            latest.current(pane);
            setAsked({ pane });
          })
        : undefined,
    [editor]
  );
  // A host that shows a closed panel for a request mounts it after the request: take it now.
  useEffect(() => {
    const pane = takePaneRequest(editor);
    if (!pane) return;
    latest.current(pane);
    setAsked({ pane });
  }, [editor]);
  // A request (a comment marker, the Selection pane control) lands the focus in the panel: on the
  // pane's current item when it marks one (`aria-current`), else on the tab. Focus that a pane
  // already took for the request (a new comment draft) stays.
  useEffect(() => {
    if (!asked) return;
    const body = root.current?.querySelector('.pptx-panel__body');
    if (body?.contains(body.ownerDocument.activeElement)) return;
    const current = root.current?.querySelector<HTMLElement>(
      '.pptx-panel__body [aria-current="true"]'
    );
    (current ?? root.current?.querySelector<HTMLElement>(`[data-tab="${asked.pane}"]`))?.focus();
  }, [asked]);
  return (
    <aside
      ref={root}
      className={['pptx-panel', className].filter(Boolean).join(' ')}
      aria-label={t('panel.ariaLabel')}
      style={paneStyle(paneWidth, style)}
      onMouseDown={keepFocus}
    >
      <div className="pptx-panel__head">
        <div
          className="pptx-tabs"
          role="tablist"
          // One Tab stop: the selected tab; the arrow keys move between tabs.
          onKeyDown={(e) => {
            const next = radioGroupStep(tabs, tab, e.key);
            if (!next) return;
            e.preventDefault();
            setTab(next);
            e.currentTarget.querySelector<HTMLElement>(`[data-tab="${next}"]`)?.focus();
          }}
        >
          {tabs.map((id) => (
            <button
              key={id}
              type="button"
              role="tab"
              className="pptx-tabs__tab"
              data-tab={id}
              aria-selected={tab === id}
              // The row shows icons; the selected tab also shows its label. Each tab keeps its
              // label as its name and tooltip.
              aria-label={t(`panel.tabs.${id}`)}
              data-pptx-tooltip={t(`panel.tabs.${id}`)}
              tabIndex={tab === id ? 0 : -1}
              onClick={() => setTab(id)}
            >
              <Icon name={SIDE_PANEL_TAB_ICONS[id]} />
              {tab === id ? (
                <span className="pptx-tabs__label" aria-hidden="true">
                  {t(`panel.tabs.${id}`)}
                </span>
              ) : null}
            </button>
          ))}
        </div>
        {onClose ? <ToolButton label={t('panel.close')} icon="close" onClick={onClose} /> : null}
      </div>
      <div
        className="pptx-panel__body"
        role="tabpanel"
        onKeyDown={(e) => panelHistoryKey(e, editor)}
      >
        {panes?.[tab] !== undefined ? panes[tab] : <PackagedPane tab={tab} />}
      </div>
    </aside>
  );
}
