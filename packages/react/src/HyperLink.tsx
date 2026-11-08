/** @jsxImportSource react */
// `PptxEditor.HyperLink`: the link popup with named parts, after DOCX Editor's
// `DocxEditor.HyperLink` (packages/react/src/editor/DocxEditorHyperLink.tsx at 6907dc963).
// A click into linked text opens it in reading mode under the caret: the target, Copy, Edit,
// and Unlink. Ctrl/Cmd+K, the link control, Insert > Link, and Edit open it in edit mode: the
// URL field, Apply, and Cancel. The state is the root's (core's `createHyperlinkPopupController`),
// so the toolbar, the keyboard, and a host panel agree.
//
// CUSTOMIZATION LADDER, the toolbar's five rungs:
//   1. `className` / `data-mode`   restyle the default parts with CSS
//   2. `icon`                      swap an action's glyph (Copy, Edit, Unlink)
//   3. `asChild`                   put a part's wiring on your own element
//   4. in-place part override      a `<HyperLink.Copy>` child replaces that part; `hidden`
//                                  removes it; `preset={false}` drops the defaults
//   5. `useHyperlinkPopup()`       the raw state and actions, for a UI of your own
//
// Render it from `popups.hyperlink` to change it; the root renders the default one. A renderer
// shows reading mode only when registered with `definePopup(Component, { reading: true })`.
// F6 moves focus from the slide text into an open reading panel.

import {
  Children,
  createContext,
  Fragment,
  isValidElement,
  useContext,
  useEffect,
  useMemo,
  useRef,
  useState,
  type ReactElement,
  type ReactNode,
} from 'react';
import {
  customShowDraft,
  customShowDraftUrl,
  fileDraftUrl,
  customShowLinkLabel,
  linkKindDraft,
  linkKindOf,
  linkKindsFor,
  linkSelectionRect,
  mailtoParts,
  mailtoUrl,
  slideLinkChoices,
  slideLinkLabel,
  type LinkKind,
  type MailtoParts,
  openHyperlinkPopup,
  requestLinkPopup,
  translateKeyOrText,
  type HyperlinkPopupController,
  type HyperlinkPopupState,
} from '@pptx-editor.dev/core/editor';
import type { CustomShowSummary } from '@pptx-editor.dev/core/contracts';
import { Icon, Popover } from './chrome.tsx';
import { useEditorState, usePptxEditor, useTranslation } from './context.ts';
import { renderAsChild, type HostElementProps } from './slot-override.ts';

/** What `useHyperlinkPopup` returns. @public */
export interface UseHyperlinkPopupResult {
  readonly state: HyperlinkPopupState;
  /**
   * Open reading mode over the link at a text caret, or the URL field when there is none, under
   * `anchor` (default: the caret). DOCX Editor's argument order: `link` is `null` for the URL
   * field at once; the panel shows the link at the caret, so a given `link` only asks for
   * reading mode.
   */
  open(link?: { readonly url: string } | null, anchor?: HTMLElement | null): void;
  /** Open the URL field for the selection, as Ctrl/Cmd+K and the link control do. */
  openAtCaret(): void;
  close(): void;
  /** Copy the safe target. Resolves false when there is nothing safe to copy. */
  copy(): Promise<boolean>;
  /** Switch from reading to editing. */
  beginEdit(): void;
  setUrl(url: string): void;
  /** Apply the draft. False when the engine refused it; `state.error` says why. */
  commitEdit(): boolean;
  /** Remove the link, keeping its text. */
  unlink(): boolean;
  /** Open the safe target in a new tab, with no opener. False when there is none. */
  openTarget(): boolean;
}

/** The root's link popup state, published by `PptxEditor.Root`. @internal */
export interface HyperlinkPopupContextValue {
  readonly controller: HyperlinkPopupController;
  readonly state: HyperlinkPopupState;
  /** The id of the error message, so the URL field names it. */
  readonly errorId: string;
}

/** @internal */
export const HyperlinkPopupContext = createContext<HyperlinkPopupContextValue | null>(null);

/** No custom shows: one array, so the selector result stays the same. */
const NO_SHOWS: readonly CustomShowSummary[] = [];

const CLOSED_STATE: HyperlinkPopupState = {
  mode: 'closed',
  url: '',
  href: null,
  draft: '',
  error: null,
  copied: false,
  canEdit: false,
  anchor: null,
};

/**
 * The link popup's state and actions: the root's one state, so a host panel, the toolbar, and
 * Ctrl/Cmd+K agree. Call it in a part of the panel or in a `popups.hyperlink` renderer, which the
 * root renders inside the state; elsewhere it is closed and inert. @public
 */
export function useHyperlinkPopup(): UseHyperlinkPopupResult {
  const ctx = useContext(HyperlinkPopupContext);
  const editor = usePptxEditor();
  const c = ctx?.controller;
  return {
    state: ctx?.state ?? CLOSED_STATE,
    open: (link, anchor) => {
      if (c && editor) openHyperlinkPopup(c, editor, link, anchor);
    },
    openAtCaret: () => {
      if (editor) requestLinkPopup(editor);
    },
    close: () => c?.close(),
    copy: () => c?.copy() ?? Promise.resolve(false),
    beginEdit: () => c?.beginEdit(),
    setUrl: (url) => c?.setUrl(url),
    commitEdit: () => c?.commitEdit() ?? false,
    unlink: () => c?.unlink() ?? false,
    openTarget: () => c?.openTarget() ?? false,
  };
}

/** Shared props of every part. @public */
export interface HyperLinkPartProps {
  readonly className?: string;
  /** Put this part's wiring on the one child element instead of the default element. */
  readonly asChild?: boolean;
  /** Render nothing: inside the default arrangement this removes the part. */
  readonly hidden?: boolean;
  readonly children?: ReactNode;
}

/** Props of the action parts, which also take an icon. @public */
export interface HyperLinkActionProps extends HyperLinkPartProps {
  /** The glyph; falls back to `children`, then to the part's default icon. */
  readonly icon?: ReactNode;
}

/** Props of `PptxEditor.HyperLink`. @public */
export interface HyperLinkProps extends HyperLinkPartProps {
  /** Render the packaged parts. `false` renders only the panel and your children. */
  readonly preset?: boolean;
}

/** Keeps the caret: a press in the panel must not move it. Fields are exempt. */
function guardMousedown(event: { target: EventTarget | null; preventDefault(): void }): void {
  const tag = (event.target as HTMLElement | null)?.tagName;
  if (tag === 'INPUT' || tag === 'SELECT' || tag === 'TEXTAREA') return;
  event.preventDefault();
}

const join = (base: string, extra?: string): string => (extra ? `${base} ${extra}` : base);

/** The part's own element, or its wiring on the host's child (`asChild`). */
function render(
  asChild: boolean | undefined,
  children: ReactNode,
  props: HostElementProps,
  fallback: () => ReactElement
): ReactElement | null {
  if (!asChild) return fallback();
  const only = Children.toArray(children);
  const child = only.length === 1 ? only[0] : null;
  return isValidElement<HostElementProps>(child) ? renderAsChild(child, props) : fallback();
}

/** The link popup panel. @public */
function HyperLinkRoot({ className, asChild, hidden, preset = true, children }: HyperLinkProps) {
  const ctx = useContext(HyperlinkPopupContext);
  const { t } = useTranslation();
  const state = ctx?.state ?? CLOSED_STATE;
  const editor = usePptxEditor();
  const open = state.mode !== 'closed';
  // Capture the text box before the field takes focus. Keep it while the draft changes.
  const anchorRect = useMemo(
    () => (open && editor && state.anchor ? linkSelectionRect(editor) : null),
    [editor, open, state.anchor]
  );
  if (!ctx || !open) return null;
  // `hidden` drops the packaged panel only: the gestures stay, so a host renders its own UI.
  if (hidden) return <>{children}</>;
  const body = preset ? <HyperLinkPreset>{children}</HyperLinkPreset> : children;
  const reading = state.mode === 'reading';
  const inner = { className: 'pptx-link', 'data-mode': state.mode };
  return (
    <Popover
      anchor={state.anchor}
      anchorRect={anchorRect}
      role="dialog"
      label={t(reading ? 'link.label' : state.url ? 'link.edit' : 'link.insert')}
      onClose={() => ctx.controller.close()}
      // Reading leaves the keys in the slide text; editing takes them in its field.
      passive={reading}
      field={!reading}
      className={join('pptx-link-popup', className)}
    >
      {render(asChild, children, inner, () => (
        <div {...inner}>{body}</div>
      ))}
    </Popover>
  );
}

type PartName = 'Url' | 'Copy' | 'Edit' | 'Unlink' | 'Fields' | 'Error' | 'Apply' | 'Cancel';
type Marked = { pptxHyperLinkPart?: PartName };

/** Each part child by its name, so the preset swaps it in place; other children follow. */
function partOverrides(children: ReactNode): {
  parts: Partial<Record<PartName, ReactNode>>;
  extra: ReactNode[];
} {
  const parts: Partial<Record<PartName, ReactNode>> = {};
  const extra: ReactNode[] = [];
  const visit = (node: ReactNode): void => {
    if (Array.isArray(node)) return void node.forEach(visit);
    if (!isValidElement(node)) {
      if (node !== null && node !== undefined && node !== false) extra.push(node);
      return;
    }
    // A fragment groups overrides; it is not content.
    if (node.type === Fragment) return visit((node.props as { children?: ReactNode }).children);
    const name = (node.type as Marked).pptxHyperLinkPart;
    if (name) parts[name] = node;
    else extra.push(node);
  };
  visit(children);
  return { parts, extra };
}

/** The default arrangement of the mode. A part child replaces its default copy. */
function HyperLinkPreset({ children }: { children?: ReactNode }) {
  const { state } = useHyperlinkPopup();
  const { parts, extra } = partOverrides(children);
  const take = (name: PartName, fallback: ReactNode): ReactNode =>
    name in parts ? parts[name] : fallback;
  if (state.mode === 'editing')
    return (
      <>
        {/* An existing link keeps its target, Copy, and Open while it is edited. */}
        {state.url ? (
          <div className="pptx-link__row">
            {take('Url', <HyperLinkUrl />)}
            <div className="pptx-link__tools">{take('Copy', <HyperLinkCopy />)}</div>
          </div>
        ) : null}
        {take('Fields', <HyperLinkFields />)}
        {take('Error', <HyperLinkError />)}
        <div className="pptx-link__actions">
          {/* Changing a link can also remove it, as the reading row can. */}
          {state.url && state.canEdit ? take('Unlink', <HyperLinkUnlink />) : null}
          {take('Cancel', <HyperLinkCancel />)}
          {take('Apply', <HyperLinkApply />)}
        </div>
        {extra}
      </>
    );
  return (
    <>
      <div className="pptx-link__row">
        {take('Url', <HyperLinkUrl />)}
        <div className="pptx-link__tools">
          {take('Copy', <HyperLinkCopy />)}
          {/* An action that can never run is absent, not disabled: a read-only editor. */}
          {state.canEdit ? take('Edit', <HyperLinkEdit />) : null}
          {state.canEdit ? take('Unlink', <HyperLinkUnlink />) : null}
        </div>
      </div>
      {take('Error', <HyperLinkError />)}
      {extra}
    </>
  );
}

/** The target. A safe link opens in a new tab; any other shows as text. @public */
function HyperLinkUrl({ className, asChild, hidden, children }: HyperLinkPartProps) {
  const { state, openTarget } = useHyperlinkPopup();
  const { t } = useTranslation();
  const shows = useEditorState((s) => s.customShows ?? NO_SHOWS);
  if (hidden || !state.url) return null;
  // A custom show link names its show; it opens in a slide show, not here.
  const show = customShowLinkLabel(t, state.url, shows);
  if (show) {
    const props = { className: join('pptx-link__url', className), 'data-inert': '' };
    return render(asChild, children, props, () => <span {...props}>{children ?? show}</span>);
  }
  const slide = slideLinkLabel(t, state.url);
  if (slide) {
    const props: HostElementProps = {
      type: 'button',
      className: join('pptx-link__url pptx-link__slide', className),
      'aria-label': t('link.goToTarget', { target: slide }),
      title: t('link.goToTarget', { target: slide }),
      onMouseDown: guardMousedown,
      onClick: () => void openTarget(),
    };
    return render(asChild, children, props, () => <button {...props}>{children ?? slide}</button>);
  }
  const text = <bdi dir="ltr">{state.url}</bdi>;
  if (!state.href) {
    const props = {
      className: join('pptx-link__url', className),
      'data-inert': '',
      title: t('link.unsafe'),
    };
    return render(asChild, children, props, () => <span {...props}>{children ?? text}</span>);
  }
  const props = {
    className: join('pptx-link__url', className),
    href: state.href,
    target: '_blank',
    rel: 'noopener noreferrer',
    title: t('link.open'),
    onMouseDown: guardMousedown,
  };
  return render(asChild, children, props, () => (
    <a {...props}>
      {children ?? text}
      <Icon name="open_in_new" />
    </a>
  ));
}
HyperLinkUrl.pptxHyperLinkPart = 'Url' as const;

/** An icon button of the reading row. */
function ActionPart(
  { className, asChild, hidden, children, icon }: HyperLinkActionProps,
  part: string,
  label: string,
  glyph: string,
  onClick: () => void,
  extra: HostElementProps = {}
): ReactElement | null {
  if (hidden) return null;
  const props: HostElementProps = {
    type: 'button',
    className: join(`pptx-link__action pptx-link__${part}`, className),
    'aria-label': label,
    title: label,
    onMouseDown: guardMousedown,
    onClick,
    ...extra,
  };
  return render(asChild, children, props, () => (
    <button {...props}>{icon ?? children ?? <Icon name={glyph} />}</button>
  ));
}

/** Copy the safe target. @public */
function HyperLinkCopy(props: HyperLinkActionProps) {
  const { state, copy } = useHyperlinkPopup();
  const { t } = useTranslation();
  // A slide link names a place in this deck: there is nothing to paste elsewhere.
  if (!state.href || state.href.startsWith('#')) return null;
  const label = t(state.copied ? 'link.copied' : 'link.copy');
  return ActionPart(props, 'copy', label, 'content_copy', () => void copy(), {
    ...(state.copied ? { 'data-copied': '' } : {}),
  });
}
HyperLinkCopy.pptxHyperLinkPart = 'Copy' as const;

/** Switch to the URL field. @public */
function HyperLinkEdit(props: HyperLinkActionProps) {
  const { beginEdit } = useHyperlinkPopup();
  const { t } = useTranslation();
  return ActionPart(props, 'edit', t('link.edit'), 'edit', beginEdit);
}
HyperLinkEdit.pptxHyperLinkPart = 'Edit' as const;

/** Remove the link, keeping its text. @public */
function HyperLinkUnlink(props: HyperLinkActionProps) {
  const { unlink } = useHyperlinkPopup();
  const { t } = useTranslation();
  return ActionPart(props, 'unlink', t('link.remove'), 'link_off', () => void unlink());
}
HyperLinkUnlink.pptxHyperLinkPart = 'Unlink' as const;

const KIND_LABEL = {
  web: 'link.kindWeb',
  email: 'link.kindEmail',
  slide: 'link.kindSlide',
  customShow: 'link.kindCustomShow',
  file: 'link.kindFile',
} as const;

/**
 * The link's target: **Link to** (a web address, an email address, a slide or a custom show of
 * this presentation, or another file by its relative path) and its fields. Each kind writes the
 * one draft that Apply sends. Enter applies. @public
 */
function HyperLinkFields({ className, hidden }: HyperLinkPartProps) {
  const ctx = useContext(HyperlinkPopupContext);
  const { state, setUrl, commitEdit } = useHyperlinkPopup();
  const { t } = useTranslation();
  const slides = useEditorState((s) => s.slides.length);
  const shows = useEditorState((s) => s.customShows ?? NO_SHOWS);
  const [kind, setKind] = useState<LinkKind>(() => linkKindOf(state.draft));
  // The kinds stay as they opened: Custom show only for a deck with one, or a link to one.
  const [kinds] = useState(() => linkKindsFor(shows, state.draft));
  const show = customShowDraft(state.draft);
  const [mail, setMail] = useState(() => mailtoParts(state.draft));
  const field = useRef<HTMLInputElement & HTMLSelectElement>(null);
  // After the popover is placed: a browser does not focus into a hidden panel. A change of kind
  // keeps focus on Link to, so the arrow keys can step through the kinds.
  useEffect(() => {
    field.current?.focus({ preventScroll: true });
    if (field.current instanceof HTMLInputElement) field.current.select();
  }, []);
  if (hidden) return null;
  const invalid = {
    'aria-invalid': state.error ? true : undefined,
    'aria-describedby': state.error && ctx ? ctx.errorId : undefined,
  } as const;
  const onEnter = (e: { key: string; preventDefault(): void }): void => {
    if (e.key !== 'Enter') return;
    // The Enter that applies must not reach the element focus moves to next.
    e.preventDefault();
    commitEdit();
  };
  const setMailField = (next: MailtoParts): void => {
    setMail(next);
    setUrl(mailtoUrl(next.address, next.subject, next.rest));
  };
  return (
    <div className={join('pptx-link__fields', className)}>
      <label className="pptx-field pptx-link__kind">
        {t('link.kind')}
        <select
          value={kind}
          onChange={(e) => {
            const next = e.target.value as LinkKind;
            setKind(next);
            setMail({ address: '', subject: '', rest: '' });
            setUrl(linkKindDraft(next, shows));
          }}
        >
          {kinds.map((k) => (
            <option key={k} value={k}>
              {t(KIND_LABEL[k])}
            </option>
          ))}
        </select>
      </label>
      {kind === 'web' ? (
        <input
          ref={field}
          className="pptx-font-picker__search pptx-link__field"
          type="url"
          dir="ltr"
          aria-label={t('link.label')}
          placeholder={t('link.placeholder')}
          value={state.draft}
          {...invalid}
          onChange={(e) => setUrl(e.target.value)}
          onKeyDown={onEnter}
        />
      ) : kind === 'email' ? (
        <>
          <input
            ref={field}
            className="pptx-font-picker__search pptx-link__field"
            type="email"
            dir="ltr"
            aria-label={t('link.emailAddress')}
            placeholder={t('link.emailAddress')}
            value={mail.address}
            {...invalid}
            onChange={(e) => setMailField({ ...mail, address: e.target.value })}
            onKeyDown={onEnter}
          />
          <input
            className="pptx-font-picker__search pptx-link__field"
            type="text"
            aria-label={t('link.emailSubject')}
            placeholder={t('link.emailSubject')}
            value={mail.subject}
            onChange={(e) => setMailField({ ...mail, subject: e.target.value })}
            onKeyDown={onEnter}
          />
        </>
      ) : kind === 'customShow' ? (
        <>
          <select
            ref={field}
            className="pptx-link__field"
            aria-label={t('link.customShow')}
            value={show ? String(show.id) : ''}
            {...invalid}
            onChange={(e) =>
              setUrl(customShowDraftUrl(Number(e.target.value), show?.returns ?? true))
            }
            onKeyDown={onEnter}
          >
            {/* A link to a show the deck no longer has keeps its id as a choice. */}
            {show && !shows.some((s) => s.id === show.id) ? (
              <option value={String(show.id)}>{customShowLinkLabel(t, state.draft, shows)}</option>
            ) : null}
            {shows.map((s) => (
              <option key={s.id} value={String(s.id)}>
                {s.name || String(s.id)}
              </option>
            ))}
          </select>
          <label className="pptx-radio pptx-link__return">
            <input
              type="checkbox"
              checked={show?.returns ?? true}
              disabled={!show}
              onChange={(e) => show && setUrl(customShowDraftUrl(show.id, e.target.checked))}
            />
            {t('link.returnAfterShow')}
          </label>
        </>
      ) : kind === 'file' ? (
        <input
          ref={field}
          className="pptx-font-picker__search pptx-link__field"
          type="text"
          dir="ltr"
          aria-label={t('link.filePath')}
          placeholder={t('link.filePlaceholder')}
          value={state.draft}
          {...invalid}
          onChange={(e) => setUrl(fileDraftUrl(e.target.value))}
          onKeyDown={onEnter}
        />
      ) : (
        <select
          ref={field}
          className="pptx-link__field"
          aria-label={t('link.slide')}
          value={state.draft}
          {...invalid}
          onChange={(e) => setUrl(e.target.value)}
          onKeyDown={onEnter}
        >
          {slideLinkChoices(slides).map((url) => (
            <option key={url} value={url}>
              {slideLinkLabel(t, url)}
            </option>
          ))}
        </select>
      )}
    </div>
  );
}
HyperLinkFields.pptxHyperLinkPart = 'Fields' as const;

/** Why the last Apply or Unlink was refused. @public */
function HyperLinkError({ className, hidden }: HyperLinkPartProps) {
  const ctx = useContext(HyperlinkPopupContext);
  const { t } = useTranslation();
  const error = ctx?.state.error;
  if (hidden || !error) return null;
  return (
    <p id={ctx.errorId} role="alert" className={join('pptx-link__error', className)}>
      {translateKeyOrText(t, error)}
    </p>
  );
}
HyperLinkError.pptxHyperLinkPart = 'Error' as const;

/** Apply the draft. @public */
function HyperLinkApply({ className, asChild, hidden, children }: HyperLinkPartProps) {
  const { state, commitEdit } = useHyperlinkPopup();
  const { t } = useTranslation();
  if (hidden) return null;
  const props: HostElementProps = {
    type: 'button',
    className: join('pptx-button pptx-button--primary pptx-link__apply', className),
    disabled: state.draft.trim().length === 0,
    onMouseDown: guardMousedown,
    onClick: () => void commitEdit(),
  };
  return render(asChild, children, props, () => (
    <button {...props}>{children ?? t('link.apply')}</button>
  ));
}
HyperLinkApply.pptxHyperLinkPart = 'Apply' as const;

/** Close without applying. @public */
function HyperLinkCancel({ className, asChild, hidden, children }: HyperLinkPartProps) {
  const { close } = useHyperlinkPopup();
  const { t } = useTranslation();
  if (hidden) return null;
  const props: HostElementProps = {
    type: 'button',
    className: join('pptx-button pptx-link__cancel', className),
    onMouseDown: guardMousedown,
    onClick: close,
  };
  return render(asChild, children, props, () => (
    <button {...props}>{children ?? t('link.cancel')}</button>
  ));
}
HyperLinkCancel.pptxHyperLinkPart = 'Cancel' as const;

/** The link popup compound. @public */
export interface HyperLinkNamespace {
  (props: HyperLinkProps): ReactNode;
  readonly Url: typeof HyperLinkUrl;
  readonly Copy: typeof HyperLinkCopy;
  readonly Edit: typeof HyperLinkEdit;
  readonly Unlink: typeof HyperLinkUnlink;
  readonly Fields: typeof HyperLinkFields;
  readonly Error: typeof HyperLinkError;
  readonly Apply: typeof HyperLinkApply;
  readonly Cancel: typeof HyperLinkCancel;
}

/** `PptxEditor.HyperLink`: the link popup with named parts. @public */
export const HyperLink: HyperLinkNamespace = Object.assign(HyperLinkRoot, {
  Url: HyperLinkUrl,
  Copy: HyperLinkCopy,
  Edit: HyperLinkEdit,
  Unlink: HyperLinkUnlink,
  Fields: HyperLinkFields,
  Error: HyperLinkError,
  Apply: HyperLinkApply,
  Cancel: HyperLinkCancel,
});
