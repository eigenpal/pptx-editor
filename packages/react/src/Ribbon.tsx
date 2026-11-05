/** @jsxImportSource react */
// `PptxEditor.Ribbon`: tabs over chrome groups. The ribbon is a layer over the toolbar: each tab panel is the toolbar's preset
// bar with that tab's groups, so a tab keeps every rung of the customization ladder
// (`Toolbar.Slot`, named parts, `Toolbar.Group`, `Toolbar.Action`, `hidden`, the More menu).
// DOCX Editor has no ribbon; its equivalent of a tab is a `Toolbar` with `preset={false}` and
// that tab's `Toolbar.Group`s. `PptxEditor.Toolbar` stays as the flat form of the same groups.
//
// The arrangement comes from core (`RIBBON_TABS`, `ribbonTabs`). The ribbon's only state is the
// selected tab: controlled with `tab`, or uncontrolled with `defaultTab`.

import {
  Children,
  Fragment,
  isValidElement,
  useEffect,
  useId,
  useRef,
  useState,
  type KeyboardEvent as ReactKeyboardEvent,
  type ReactNode,
} from 'react';
import {
  arrangeRibbonTabs,
  ribbonInsertTab,
  ribbonShownTab,
  ribbonTabs,
  translateKeyOrText,
  warnRibbonChild,
  type EditorSnapshot,
  type RibbonTabId,
  type RibbonTabLayout,
} from '@pptx-editor.dev/core/editor';
import { keepFocus } from './chrome.tsx';
import { useEditorState, useTranslation } from './context.ts';
import { PresetBar } from './Toolbar.tsx';

/** Props for `PptxEditor.Ribbon.Tab`. @public */
export interface RibbonTabProps {
  /** A built-in tab id changes that tab; any other id adds a host tab. */
  readonly id: RibbonTabId | (string & {});
  /** The tab's title. Wins over `labelKey`. Built-in tabs keep their own title without it. */
  readonly label?: string;
  /** A translation key for the tab's title. */
  readonly labelKey?: string;
  /** Remove the tab. Its commands keep working from menus, keys, and the API. */
  readonly hidden?: boolean;
  /** For a host tab: the tab it follows. Without it, host tabs follow the built-in ones. */
  readonly after?: RibbonTabId | (string & {});
  /** A class name added to the tab's panel. */
  readonly className?: string;
  /**
   * `Toolbar.Slot`, named parts, `Toolbar.Group`, and `Toolbar.Action` for the tab's panel, as
   * toolbar children: they change the tab's groups, or are the groups of a host tab.
   */
  readonly children?: ReactNode;
}

/** One tab of the ribbon. Inside `PptxEditor.Ribbon` it is a description the ribbon reads. @public */
export function RibbonTab(_props: RibbonTabProps): ReactNode {
  return null;
}

/** Props for `PptxEditor.Ribbon`. @public */
export interface RibbonProps {
  /** A class name added to the ribbon. */
  readonly className?: string;
  /**
   * `true` (default): the tabs of `RIBBON_TABS`, changed by `Ribbon.Tab` children. `false`:
   * only the `Ribbon.Tab` children.
   */
  readonly preset?: boolean;
  /** The selected tab, controlled. Use with `onTabChange`. */
  readonly tab?: RibbonTabId | (string & {});
  /** The tab selected first, uncontrolled. Default `home`. */
  readonly defaultTab?: RibbonTabId | (string & {});
  /** Called when the user selects another tab. */
  readonly onTabChange?: (tab: string) => void;
  /** `false`: a tab's groups wrap instead of moving into the More menu. Default `true`. */
  readonly overflow?: boolean;
  /** `Ribbon.Tab` children. */
  readonly children?: ReactNode;
}

type NodeProps = Record<string, unknown> & { children?: ReactNode };

const flatten = (nodes: ReactNode): ReactNode[] =>
  Children.toArray(nodes).flatMap((node) =>
    isValidElement<NodeProps>(node) && node.type === Fragment
      ? flatten(node.props.children)
      : [node]
  );

/** Whether two tab layouts show the same tabs and groups. */
function sameLayout(a: readonly RibbonTabLayout[], b: readonly RibbonTabLayout[]): boolean {
  const key = (l: readonly RibbonTabLayout[]) =>
    l.map((t) => `${t.tab.id}:${t.groups.map((g) => g.id).join(',')}`).join('|');
  return key(a) === key(b);
}

const selectTabs = (s: EditorSnapshot): RibbonTabLayout[] => ribbonTabs(s);

/**
 * The ribbon: the Quick Access Toolbar, the tab names, and the selected tab's groups, with
 * contextual tabs (Shape Format, Table Design, …) while the selection needs them.
 *
 * @example
 * ```tsx
 * <PptxEditor.Ribbon defaultTab="home">
 *   <PptxEditor.Ribbon.Tab id="home">
 *     <PptxEditor.Toolbar.Strike hidden />
 *   </PptxEditor.Ribbon.Tab>
 *   <PptxEditor.Ribbon.Tab id="view" hidden />
 *   <PptxEditor.Ribbon.Tab id="review" label="Review">
 *     <PptxEditor.Toolbar.Group id="legal" label="Legal">
 *       <PptxEditor.Toolbar.Action label="Ask legal" onSelect={askLegal} />
 *     </PptxEditor.Toolbar.Group>
 *   </PptxEditor.Ribbon.Tab>
 * </PptxEditor.Ribbon>
 * ```
 *
 * @public
 */
export function Ribbon({
  className,
  preset = true,
  tab,
  defaultTab = 'home',
  onTabChange,
  overflow = true,
  children,
}: RibbonProps): ReactNode {
  const { t } = useTranslation();
  const baseId = useId();
  const layout = useEditorState(selectTabs, sameLayout);
  const [own, setOwn] = useState<string>(defaultTab);
  const names = useRef(new Map<string, HTMLButtonElement>());
  const root = useRef<HTMLDivElement>(null);
  // Whether the keyboard focus is in the ribbon. A removed panel takes the focus with it and
  // sends no `focusout` with a target, so the flag stays set until the focus goes elsewhere.
  const focusInside = useRef(false);
  const specs = new Map<string, RibbonTabProps>();
  for (const child of flatten(children)) {
    if (isValidElement<NodeProps>(child) && child.type === RibbonTab) {
      const props = child.props as unknown as RibbonTabProps;
      specs.set(props.id, props);
    } else if (
      child !== null &&
      child !== undefined &&
      typeof child !== 'boolean' &&
      !(typeof child === 'string' && child.trim() === '')
    )
      warnRibbonChild('PptxEditor.Ribbon', 'Ribbon.Tab');
  }
  const { quickAccess, tabs } = arrangeRibbonTabs(
    layout,
    specs,
    preset,
    (key) => translateKeyOrText(t, key),
    'Ribbon.Tab'
  );
  const wanted = tab ?? own;
  const { shown: current, fellBack } = ribbonShownTab(tabs, wanted);
  const shownId = current?.id;
  // A tab that went away (a contextual tab after the selection changed) gives way to Home: the
  // ribbon selects Home, so the gone tab does not come back by itself when its context does,
  // and a controlled host hears of it. The focus moves to the shown tab name.
  useEffect(() => {
    if (!fellBack || shownId === undefined) return;
    if (tab === undefined) setOwn(shownId);
    onTabChange?.(shownId);
    const active = root.current?.ownerDocument.activeElement;
    if (focusInside.current && (!active || active === active.ownerDocument.body))
      names.current.get(shownId)?.focus();
    // Only a change of the shown tab runs this; the callback is the host's latest.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [fellBack, shownId]);
  const select = (id: string): void => {
    if (tab === undefined) setOwn(id);
    onTabChange?.(id);
  };
  // An insert opens the new object's Format tab.
  const insertions = useEditorState((s) => s.insertions);
  const seen = useRef(insertions);
  const insertTab = ribbonInsertTab(tabs);
  useEffect(() => {
    if (insertions === seen.current) return;
    seen.current = insertions;
    if (insertTab !== null && insertTab !== wanted) select(insertTab);
    // Only a new insert runs this.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [insertions]);
  const onTabKey = (e: ReactKeyboardEvent): void => {
    const index = tabs.findIndex((d) => d.id === current?.id);
    const next =
      e.key === 'ArrowRight'
        ? tabs[(index + 1) % tabs.length]
        : e.key === 'ArrowLeft'
          ? tabs[(index - 1 + tabs.length) % tabs.length]
          : e.key === 'Home'
            ? tabs[0]
            : e.key === 'End'
              ? tabs[tabs.length - 1]
              : undefined;
    if (!next) return;
    e.preventDefault();
    select(next.id);
    names.current.get(next.id)?.focus();
  };
  const tabId = (id: string) => `${baseId}-tab-${id}`;
  const panelId = `${baseId}-panel`;
  return (
    <div
      ref={root}
      className={['pptx-ribbon', className].filter(Boolean).join(' ')}
      onFocus={() => {
        focusInside.current = true;
      }}
      onBlur={(e) => {
        const next = e.relatedTarget as Node | null;
        if (next && !e.currentTarget.contains(next)) focusInside.current = false;
      }}
    >
      <div className="pptx-ribbon__bar">
        {quickAccess ? (
          <div className="pptx-ribbon__qat">
            <PresetBar groups={quickAccess.groups} label={quickAccess.title} overflow={false}>
              {quickAccess.override?.children}
            </PresetBar>
          </div>
        ) : null}
        <div
          role="tablist"
          aria-label={t('ribbon.ariaLabel')}
          className="pptx-ribbon__tabs"
          onKeyDown={onTabKey}
        >
          {tabs.map((d) => {
            const selected = d.id === current?.id;
            return (
              <button
                key={d.id}
                ref={(el) => {
                  if (el) names.current.set(d.id, el);
                  else names.current.delete(d.id);
                }}
                type="button"
                role="tab"
                id={tabId(d.id)}
                className="pptx-ribbon__tab"
                aria-selected={selected}
                aria-controls={selected ? panelId : undefined}
                tabIndex={selected ? 0 : -1}
                data-tab={d.id}
                data-contextual={d.contextual ? '' : undefined}
                onMouseDown={keepFocus}
                onClick={() => select(d.id)}
              >
                {d.title}
              </button>
            );
          })}
        </div>
      </div>
      {current ? (
        <div
          role="tabpanel"
          id={panelId}
          aria-labelledby={tabId(current.id)}
          className={['pptx-ribbon__panel', current.override?.className].filter(Boolean).join(' ')}
          data-tab={current.id}
        >
          <PresetBar
            key={current.id}
            groups={current.groups}
            label={current.title}
            overflow={overflow}
          >
            {current.override?.children}
          </PresetBar>
        </div>
      ) : null}
    </div>
  );
}
