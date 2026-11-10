/** @jsxImportSource react */
// The side panel: Format, Chart Data, Diagram, Selection (the selection pane), Animations, and Comments. The Format
// tab writes the selected element's position, size, rotation, and text fitting through engine
// commands.
// The Vue twin is `packages/vue/src/SidePanel.tsx`.

import { useEffect, useRef, useState, type CSSProperties, type ReactNode } from 'react';
import {
  SIDE_PANEL_TAB_ICONS,
  listenPaneRequests,
  takePaneRequest,
  visibleSidePanelTabs,
  panelHistoryKey,
  radioGroupStep,
  type SidePanelTab,
} from '@pptx-editor.dev/core/editor';
import { Icon, ToolButton, keepFocus } from './chrome.tsx';
import { useEditorState, usePptxEditor, useTranslation } from './context.ts';
import { paneStyle } from './pane-state.ts';
import { SelectionPane } from './SelectionPane.tsx';
import { AnimationsPane } from './AnimationsPane.tsx';
import { ChartDataPane } from './ChartDataPane.tsx';
import { DiagramTextPane } from './DiagramTextPane.tsx';
import { FormatPanel } from './FormatPanel.tsx';

/** A tab's packaged pane. The Comments tab has none: a review pane fills it (`panes`). */
function PackagedPane({ tab }: { readonly tab: SidePanelTab }) {
  const { t } = useTranslation();
  if (tab === 'format') return <FormatPanel />;
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
        ? listenPaneRequests(editor, (pane, request) => {
            takePaneRequest(editor);
            latest.current(pane);
            if (request.focus) setAsked({ pane });
          })
        : undefined,
    [editor]
  );
  // A host that shows a closed panel for a request mounts it after the request: take it now.
  useEffect(() => {
    const request = takePaneRequest(editor);
    if (!request) return;
    latest.current(request.pane);
    if (request.focus) setAsked({ pane: request.pane });
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
