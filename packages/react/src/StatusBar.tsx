/** @jsxImportSource react */
// `PptxEditor.StatusBar`: the status bar under the slide (openspec/
// changes/editor-ui-parity/design.md, "Ground truth": slide x of y, notes, views, zoom).
// The parts follow the customization ladder: each takes `className` and `hidden`, a named part
// child replaces its part in place, other children follow the parts, and `preset={false}`
// renders only the children. The view buttons and zoom are the editor's own `ViewSwitch` and
// `ZoomPill`, so the bar adds no second control for the same action.

import { Children, Fragment, isValidElement, type ReactNode } from 'react';
import { slideStatus, warnAsChildChild } from '@pptx-editor.dev/core/editor';
import { Icon, keepFocus } from './chrome.tsx';
import { ToolbarAction } from './toolbar-action.tsx';
import { useEditorState, useTranslation } from './context.ts';
import { useNotesPane } from './pane-state.ts';
import { ZoomPill } from './panes.tsx';
import { ViewSwitch } from './SlideSorter.tsx';

/** Props of every status bar part. @public */
export interface StatusBarPartProps {
  /** A class name added to the part. */
  readonly className?: string;
  /** Remove the part from the status bar. */
  readonly hidden?: boolean;
}

/** "Slide 2 of 5". @public */
export function StatusBarSlideNumber({ className, hidden }: StatusBarPartProps): ReactNode {
  const { t } = useTranslation();
  const status = useEditorState(
    (s) => slideStatus(s),
    (a, b) => a?.current === b?.current && a?.total === b?.total
  );
  if (hidden) return null;
  return (
    <span
      className={['pptx-statusbar__slide', className].filter(Boolean).join(' ')}
      data-part="slideNumber"
      aria-live="polite"
    >
      {status ? t('statusBar.slide', status) : null}
    </span>
  );
}

/** The part name in warnings about `StatusBar.Notes`. */
const NOTES_PART = { slot: 'StatusBar.Notes', group: 'StatusBar' } as const;

/** Props of `StatusBar.Notes`: the part props, and the ladder's `icon` and `asChild`. @public */
export interface StatusBarNotesProps extends StatusBarPartProps {
  /** Icon content in place of the packaged icon. */
  readonly icon?: ReactNode;
  /** Render the one element child as the button: it gets the label, state, and click. */
  readonly asChild?: boolean;
  /** The element for `asChild`. */
  readonly children?: ReactNode;
}

/** The Notes button: shows or hides the speaker notes (`useNotesPane`). @public */
export function StatusBarNotes({
  className,
  hidden,
  icon,
  asChild,
  children,
}: StatusBarNotesProps): ReactNode {
  const { t } = useTranslation();
  const notes = useNotesPane();
  if (hidden) return null;
  // One element child, or the packaged Notes button with its icon (and a warning that names
  // this part).
  const only = asChild ? Children.toArray(children) : [];
  const host = only.length === 1 && isValidElement(only[0]);
  if (asChild && !host) warnAsChildChild('notes', NOTES_PART);
  return (
    <span className="pptx-statusbar__notes-part" data-part="notes">
      <ToolbarAction
        label={t('statusBar.notes')}
        icon={host ? undefined : (icon ?? <Icon name="edit_note" />)}
        className={['pptx-statusbar__notes', className].filter(Boolean).join(' ')}
        active={notes.open}
        onSelect={notes.toggle}
        asChild={host}
      >
        {host ? children : null}
      </ToolbarAction>
    </span>
  );
}

/**
 * The view buttons: the editor's view switch (`view.presentationView`: Normal and Slide
 * Sorter), the same control as the default editor's status area. @public
 */
export function StatusBarViews({ className, hidden }: StatusBarPartProps): ReactNode {
  if (hidden) return null;
  return (
    <span
      className={['pptx-statusbar__views', className].filter(Boolean).join(' ')}
      data-part="views"
    >
      <ViewSwitch />
    </span>
  );
}

/** Zoom out, the zoom level with its menu, and zoom in: the editor's `ZoomPill`. @public */
export function StatusBarZoom({ className, hidden }: StatusBarPartProps): ReactNode {
  if (hidden) return null;
  return (
    <span
      className={['pptx-statusbar__zoom', className].filter(Boolean).join(' ')}
      data-part="zoom"
    >
      <ZoomPill />
    </span>
  );
}

const PARTS = [StatusBarSlideNumber, StatusBarNotes, StatusBarViews, StatusBarZoom] as const;

/** Props for `PptxEditor.StatusBar`. @public */
export interface StatusBarProps {
  /** A class name added to the bar. */
  readonly className?: string;
  /**
   * `true` (default): slide number, notes, views, and zoom, with named part children replacing
   * their parts in place and other children after the notes button. `false`: only the children.
   */
  readonly preset?: boolean;
  /** Named part children replace their parts; other children go after **Notes**. */
  readonly children?: ReactNode;
}

type NodeProps = Record<string, unknown> & { children?: ReactNode };

const flatten = (nodes: ReactNode): ReactNode[] =>
  Children.toArray(nodes).flatMap((node) =>
    isValidElement<NodeProps>(node) && node.type === Fragment
      ? flatten(node.props.children)
      : [node]
  );

/**
 * The status bar: "Slide x of y", the Notes button, the view buttons, and zoom.
 *
 * @example
 * ```tsx
 * <PptxEditor.StatusBar>
 *   <PptxEditor.StatusBar.Views hidden />
 *   <span className="acme-saved">Saved</span>
 * </PptxEditor.StatusBar>
 * ```
 *
 * @public
 */
export function StatusBar({ className, preset = true, children }: StatusBarProps): ReactNode {
  const { t } = useTranslation();
  const bar = (content: ReactNode) => (
    <div
      role="group"
      aria-label={t('statusBar.ariaLabel')}
      className={['pptx-statusbar', 'pptx-statusbar--bar', className].filter(Boolean).join(' ')}
      onMouseDown={keepFocus}
    >
      {content}
    </div>
  );
  if (!preset) return bar(children);
  const overrides = new Map<unknown, ReactNode>();
  const extras: ReactNode[] = [];
  for (const child of flatten(children)) {
    if (isValidElement(child) && (PARTS as readonly unknown[]).includes(child.type))
      overrides.set(child.type, child);
    else extras.push(child);
  }
  const part = (Part: (typeof PARTS)[number]) => overrides.get(Part) ?? <Part />;
  return bar(
    <>
      {part(StatusBarSlideNumber)}
      <span className="pptx-statusbar__spacer" />
      {part(StatusBarNotes)}
      {extras}
      {part(StatusBarViews)}
      {part(StatusBarZoom)}
    </>
  );
}
