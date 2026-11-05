/** @jsxImportSource react */
// The named parts of the slide rail's action row (`SlideRail.AddSlide`, `SlideRail.Delete`, …),
// after DOCX Editor's `Navigation` parts and this adapter's `StatusBar` parts. Each part takes
// `className`, `hidden`, `icon`, and `asChild` (the one element child renders the button). A named part child of `SlideRail` replaces its part in
// place, other children go after Delete, and `preset={false}` renders only the children. The
// parts read the rail's state from `RailActionsContext`, so they render only inside a rail.

import {
  Children,
  Fragment,
  createContext,
  isValidElement,
  useContext,
  useState,
  type ReactElement,
  type ReactNode,
} from 'react';
import type {
  PptxEditorInstance,
  RailOrientationChrome,
  SlideStep,
} from '@pptx-editor.dev/core/editor';
import { railOverflowPoint, warnAsChildChild } from '@pptx-editor.dev/core/editor';
import type { TranslationKey } from '@pptx-editor.dev/i18n';
import { ToolButton } from './chrome.tsx';
import { ContextMenu, MenuCloseContext, MenuRow, type ContextMenuPoint } from './ContextMenu.tsx';
import type { HostElementProps } from './slot-override.ts';

/** Props of every named part of the slide rail's action row. @public */
export interface SlideRailPartProps {
  /** A class name added to the button. */
  readonly className?: string;
  /** Remove the button from the rail. */
  readonly hidden?: boolean;
  /** Icon content in place of the packaged icon, about 18 pixels. */
  readonly icon?: ReactNode;
  /**
   * Render the one element child as the button: it gets the part's label, state, disabled
   * reason, and click, as `SlideRail.Action` and `Toolbar.Action` do.
   */
  readonly asChild?: boolean;
  /** The element for `asChild`. */
  readonly children?: ReactNode;
}

/** What the rail hands its action parts. @internal */
export interface RailActions {
  readonly t: (key: TranslationKey) => string;
  readonly editor: PptxEditorInstance | null;
  /** Why the slide buttons are off, or null when they work. */
  readonly why: string | null;
  readonly selected: readonly number[];
  readonly count: number;
  readonly chrome: RailOrientationChrome;
  readonly step: (step: -1 | 1) => SlideStep | null;
  readonly moveBy: (step: -1 | 1) => void;
  readonly galleryOpen: boolean;
  readonly toggleGallery: () => void;
  readonly listId: string;
  /** Whether the rail is open: the action row shows, or the folded strip. */
  readonly open: boolean;
  /** The rail's `toggle`: `false` removes Hide thumbnails and Show thumbnails. */
  readonly toggle: boolean;
  readonly collapse: () => void;
  readonly expand: () => void;
}

/** The rail's state for its action parts. @internal */
export const RailActionsContext = createContext<RailActions | null>(null);

/** Inside the overflow menu of a narrow rail, the parts render as menu rows. */
const RailOverflowContext = createContext(false);

interface RailButtonProps {
  readonly part: SlideRailPartProps;
  /** The part name in warnings, as `SlideRail.AddSlide`. */
  readonly name: string;
  readonly base: string;
  readonly label: string;
  readonly icon: string;
  readonly onClick: () => void;
  readonly disabled?: boolean;
  readonly disabledReason?: string | null;
  readonly hasPopup?: 'menu';
  readonly expanded?: boolean;
  readonly controls?: string;
}

/**
 * One rail button: the packaged icon or the host's, and the host class after the part class.
 * With `asChild`, the one element child renders the button; otherwise the packaged button
 * renders, with a warning.
 */
function RailButton({ part, name, base, icon, ...rest }: RailButtonProps): ReactNode {
  // In the overflow menu of a narrow rail: the same action as a menu row.
  if (useContext(RailOverflowContext))
    return (
      <MenuRow
        label={rest.label}
        icon={part.icon === undefined ? icon : part.icon}
        disabled={!!rest.disabled}
        reason={rest.disabled ? (rest.disabledReason ?? null) : null}
        className={[base, part.className].filter(Boolean).join(' ')}
        onSelect={rest.onClick}
      />
    );
  const only = part.asChild ? Children.toArray(part.children) : [];
  const host =
    only.length === 1 && isValidElement<HostElementProps>(only[0])
      ? (only[0] as ReactElement<HostElementProps>)
      : null;
  if (part.asChild && !host) warnAsChildChild(rest.label, { slot: name, group: 'SlideRail' });
  return (
    <ToolButton
      {...rest}
      host={host}
      icon={part.icon === undefined ? icon : undefined}
      className={[base, part.className].filter(Boolean).join(' ')}
    >
      {part.icon}
    </ToolButton>
  );
}

/** Add slide, after the current slide. @public */
export function SlideRailAddSlide(props: SlideRailPartProps): ReactNode {
  const rail = useContext(RailActionsContext);
  if (!rail || props.hidden) return null;
  return (
    <RailButton
      part={props}
      name="SlideRail.AddSlide"
      base="pptx-rail__add"
      label={rail.t('rail.addSlide')}
      icon="add"
      disabled={!!rail.why}
      disabledReason={rail.why}
      onClick={() => rail.editor?.exec({ type: 'addSlide' })}
    />
  );
}

/** The layout gallery: adds a slide with the layout the person picks. @public */
export function SlideRailChooseLayout(props: SlideRailPartProps): ReactNode {
  const rail = useContext(RailActionsContext);
  if (!rail || props.hidden) return null;
  return (
    <RailButton
      part={props}
      name="SlideRail.ChooseLayout"
      base="pptx-rail__layouts"
      label={rail.t('rail.chooseLayout')}
      icon="note_add"
      hasPopup="menu"
      expanded={rail.galleryOpen}
      disabled={!!rail.why}
      disabledReason={rail.why}
      onClick={rail.toggleGallery}
    />
  );
}

/** Duplicate the selected slides. @public */
export function SlideRailDuplicate(props: SlideRailPartProps): ReactNode {
  const rail = useContext(RailActionsContext);
  if (!rail || props.hidden) return null;
  return (
    <RailButton
      part={props}
      name="SlideRail.Duplicate"
      base="pptx-rail__duplicate"
      label={rail.t('rail.duplicateSlide')}
      icon="content_copy"
      disabled={!!rail.why}
      disabledReason={rail.why}
      onClick={() => rail.editor?.exec({ type: 'duplicateSlides', indices: rail.selected })}
    />
  );
}

/** Move the selected slides one place earlier (up, or left in a horizontal rail). @public */
export function SlideRailMoveEarlier(props: SlideRailPartProps): ReactNode {
  const rail = useContext(RailActionsContext);
  if (!rail || props.hidden) return null;
  return (
    <RailButton
      part={props}
      name="SlideRail.MoveEarlier"
      base="pptx-rail__move"
      label={rail.t(rail.chrome.moveEarlierLabel)}
      icon={rail.chrome.moveEarlierIcon}
      disabled={!!rail.why || rail.step(-1) === null}
      disabledReason={rail.why ?? rail.t('disabled.slideEdge')}
      onClick={() => rail.moveBy(-1)}
    />
  );
}

/** Move the selected slides one place later (down, or right in a horizontal rail). @public */
export function SlideRailMoveLater(props: SlideRailPartProps): ReactNode {
  const rail = useContext(RailActionsContext);
  if (!rail || props.hidden) return null;
  return (
    <RailButton
      part={props}
      name="SlideRail.MoveLater"
      base="pptx-rail__move"
      label={rail.t(rail.chrome.moveLaterLabel)}
      icon={rail.chrome.moveLaterIcon}
      disabled={!!rail.why || rail.step(1) === null}
      disabledReason={rail.why ?? rail.t('disabled.slideEdge')}
      onClick={() => rail.moveBy(1)}
    />
  );
}

/** Delete the selected slides. The last slide stays. @public */
export function SlideRailDelete(props: SlideRailPartProps): ReactNode {
  const rail = useContext(RailActionsContext);
  if (!rail || props.hidden) return null;
  return (
    <RailButton
      part={props}
      name="SlideRail.Delete"
      base="pptx-rail__delete"
      label={rail.t('rail.deleteSlide')}
      icon="delete"
      disabled={!!rail.why || rail.selected.length >= rail.count}
      disabledReason={rail.why ?? rail.t('disabled.lastSlide')}
      onClick={() => rail.editor?.exec({ type: 'deleteSlides', indices: rail.selected })}
    />
  );
}

/** Hide thumbnails: folds the rail to a narrow strip. `toggle={false}` on the rail removes it. @public */
export function SlideRailCollapse(props: SlideRailPartProps): ReactNode {
  const rail = useContext(RailActionsContext);
  if (!rail || props.hidden || !rail.toggle || !rail.open) return null;
  return (
    <RailButton
      part={props}
      name="SlideRail.Collapse"
      base="pptx-rail__collapse"
      label={rail.t('rail.collapse')}
      icon={rail.chrome.collapseIcon}
      expanded
      controls={rail.listId}
      onClick={rail.collapse}
    />
  );
}

/**
 * Show thumbnails: the one button of the folded rail strip, which opens the rail again. A
 * child of `SlideRail` replaces it in place; `hidden` leaves the strip with no button.
 * `toggle={false}` on the rail removes it. @public
 */
export function SlideRailExpand(props: SlideRailPartProps): ReactNode {
  const rail = useContext(RailActionsContext);
  if (!rail || props.hidden || !rail.toggle || rail.open) return null;
  return (
    <RailButton
      part={props}
      name="SlideRail.Expand"
      base="pptx-rail__collapse pptx-rail__expand"
      label={rail.t('rail.expand')}
      icon={rail.chrome.expandIcon}
      expanded={false}
      onClick={rail.expand}
    />
  );
}

/** The overflow button of the rail whose slide list has the id `listId`. */
function overflowButtonOf(listId: string): HTMLElement | null {
  if (typeof document === 'undefined') return null;
  const rail = document.getElementById(listId)?.closest('.pptx-rail');
  return rail?.querySelector<HTMLElement>('.pptx-rail__more') ?? null;
}

/**
 * The overflow button of a narrow rail: the secondary actions (layouts, duplicate, move,
 * delete) in one menu, as the toolbar's More menu holds the controls that do not fit.
 */
function SlideRailOverflow({ children }: { readonly children: ReactNode }): ReactNode {
  const rail = useContext(RailActionsContext);
  const [menu, setMenu] = useState<{
    readonly at: ContextMenuPoint;
    readonly owner: HTMLElement;
  } | null>(null);
  if (!rail) return null;
  const close = (): void => setMenu(null);
  return (
    <>
      <ToolButton
        label={rail.t('rail.moreActions')}
        icon="more_vert"
        className="pptx-rail__more"
        hasPopup="menu"
        expanded={menu !== null}
        onClick={() => {
          const owner = overflowButtonOf(rail.listId);
          setMenu((open) => (open || !owner ? null : { at: railOverflowPoint(owner), owner }));
        }}
      />
      {menu ? (
        <ContextMenu
          at={menu.at}
          owner={menu.owner}
          // Focus goes to the first row, as the toolbar's menus take focus when they open.
          keyboard
          label={rail.t('rail.moreActions')}
          onClose={close}
        >
          <MenuCloseContext.Provider value={close}>
            <RailOverflowContext.Provider value>{children}</RailOverflowContext.Provider>
          </MenuCloseContext.Provider>
        </ContextMenu>
      ) : null}
    </>
  );
}

/** The named parts, in their packaged order. Collapse is last and follows the host children. */
const PARTS = [
  SlideRailAddSlide,
  SlideRailChooseLayout,
  SlideRailDuplicate,
  SlideRailMoveEarlier,
  SlideRailMoveLater,
  SlideRailDelete,
] as const;

type NodeProps = Record<string, unknown> & { children?: ReactNode };

const flatten = (nodes: ReactNode): ReactNode[] =>
  Children.toArray(nodes).flatMap((node) =>
    isValidElement<NodeProps>(node) && node.type === Fragment
      ? flatten(node.props.children)
      : [node]
  );

/**
 * A child that renders no button in the action row: a part with `hidden`, Show thumbnails
 * (it renders only in the folded strip), and Hide thumbnails under `toggle={false}`.
 */
const isHidden = (node: ReactNode, toggle: boolean): boolean =>
  isValidElement<NodeProps>(node) &&
  (node.props.hidden === true ||
    node.type === SlideRailExpand ||
    (!toggle && node.type === SlideRailCollapse));

/**
 * The folded strip's content: the host's `SlideRail.Expand` child, or the packaged one.
 * @internal
 */
export function railExpandContent(children: ReactNode): ReactNode {
  const host = flatten(children).find(
    (node) => isValidElement(node) && node.type === SlideRailExpand
  );
  return host ?? <SlideRailExpand />;
}

/**
 * The rail's action row content and the number of buttons it shows: the packaged parts with
 * the host's named part children in their places and other children after Delete, or with
 * `preset` false only the children. A `folded` (narrow) rail shows New slide and one overflow
 * menu for the other parts. @internal
 */
export function railActionContent(
  children: ReactNode,
  preset: boolean,
  toggle: boolean,
  folded = false
): { readonly content: ReactNode; readonly shown: number } {
  const nodes = flatten(children);
  if (!preset)
    return { content: children, shown: nodes.filter((node) => !isHidden(node, toggle)).length };
  const overrides = new Map<unknown, ReactNode>();
  const extras: ReactNode[] = [];
  for (const child of nodes) {
    const type = isValidElement(child) ? child.type : null;
    if (type && ((PARTS as readonly unknown[]).includes(type) || type === SlideRailCollapse))
      overrides.set(type, child);
    else if (type !== SlideRailExpand) extras.push(child);
  }
  const parts: ReactNode[] = PARTS.map((Part) => overrides.get(Part) ?? <Part />);
  // A narrow rail keeps New slide and folds the rest into one menu; a hidden part stays hidden.
  const [add, ...secondary] = parts;
  const folds = folded && secondary.some((node) => !isHidden(node, toggle));
  const list = folds
    ? [
        add,
        <SlideRailOverflow key="overflow">
          {secondary.map((node, i) => (
            <Fragment key={i}>{node}</Fragment>
          ))}
        </SlideRailOverflow>,
        ...extras,
      ]
    : [...parts, ...extras];
  if (toggle) list.push(overrides.get(SlideRailCollapse) ?? <SlideRailCollapse />);
  return {
    content: (
      <>
        {list.map((node, i) => (
          <Fragment key={i}>{node}</Fragment>
        ))}
      </>
    ),
    shown: list.filter((node) => !isHidden(node, toggle)).length,
  };
}
