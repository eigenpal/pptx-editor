/** @jsxImportSource vue */
// `PptxEditorHyperLink`: the Vue twin of the React adapter's `HyperLink.tsx`, the link popup
// with named parts after DOCX Editor's `DocxEditor.HyperLink`. A click into linked text opens
// it in reading mode (the target, Copy, Edit, Unlink); Ctrl/Cmd+K, the link control, Insert >
// Link, and Edit open it in edit mode (the URL field, Apply, Cancel). The state is the root's
// (core's `createHyperlinkPopupController`). The same ladder as React: `class` / `className`
// and `data-mode`, an `icon` slot on the actions, `asChild`, in-place part override with
// `hidden` and `preset: false`, and `useHyperlinkPopup()`. A renderer shows reading mode only
// when registered with `definePopup(Component, { reading: true })`. F6 moves focus from the
// slide text into an open reading panel.

import {
  cloneVNode,
  defineComponent,
  Fragment,
  inject,
  isVNode,
  onMounted,
  ref,
  shallowRef,
  watch,
  type ExtractPublicPropTypes,
  type VNode,
  type VNodeArrayChildren,
} from 'vue';
import {
  LINK_KINDS,
  linkKindOf,
  linkSelectionRect,
  mailtoParts,
  mailtoUrl,
  slideLinkChoices,
  slideLinkLabel,
  translateKeyOrText,
  type HyperlinkPopupState,
  type LinkKind,
  type MailtoParts,
} from '@pptx-editor.dev/core/editor';
import { Icon, Popover } from './chrome.tsx';
import { useEditorState, usePptxEditor, useTranslation } from './composables.ts';
import { hyperlinkPopupKey, useHyperlinkPopup } from './useHyperlinkPopup.ts';

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

const partProps = {
  /** A class name added to the part (`class` works too). */
  className: { type: String, default: undefined },
  /** Put this part's wiring on the one child element of the default slot. */
  asChild: { type: Boolean, default: false },
  /** Render nothing: inside the default arrangement this removes the part. */
  hidden: { type: Boolean, default: false },
} as const;

/** Shared props of every part, the twin of React's `HyperLinkPartProps`. @public */
export type HyperLinkPartProps = ExtractPublicPropTypes<typeof partProps>;
/** Props of the action parts; the glyph is the `icon` slot. @public */
export type HyperLinkActionProps = HyperLinkPartProps;

const rootProps = {
  ...partProps,
  /** Render the packaged parts. `false` renders only the panel and the default slot. */
  preset: { type: Boolean, default: true },
} as const;

/** Props of `PptxEditorHyperLink`. @public */
export type HyperLinkProps = ExtractPublicPropTypes<typeof rootProps>;

/** Keeps the caret: a press in the panel must not move it. Fields are exempt. */
function guardMousedown(event: MouseEvent): void {
  const tag = (event.target as HTMLElement | null)?.tagName;
  if (tag === 'INPUT' || tag === 'SELECT' || tag === 'TEXTAREA') return;
  event.preventDefault();
}

/** The top-level vnodes of `children`, with fragments opened. */
function flatten(children: VNodeArrayChildren | undefined, out: VNode[] = []): VNode[] {
  for (const child of children ?? []) {
    if (Array.isArray(child)) flatten(child, out);
    else if (isVNode(child)) {
      if (child.type === Fragment && Array.isArray(child.children))
        flatten(child.children as VNodeArrayChildren, out);
      else out.push(child);
    }
  }
  return out;
}

/** The part's own element, or its wiring on the host's one child (`asChild`). */
function wire(
  asChild: boolean,
  children: VNode[] | undefined,
  props: Record<string, unknown>,
  fallback: () => VNode
): VNode {
  const only = asChild ? flatten(children) : [];
  const child = only.length === 1 && typeof only[0]!.type !== 'symbol' ? only[0] : undefined;
  return child ? cloneVNode(child, props, true) : fallback();
}

type PartName = 'Url' | 'Copy' | 'Edit' | 'Unlink' | 'Fields' | 'Error' | 'Apply' | 'Cancel';

/** A part component, marked with its name so the preset swaps a host copy in place. */
function part<P>(name: PartName, component: P): P {
  return Object.assign(component as object, { pptxHyperLinkPart: name }) as P;
}

const HyperLinkUrl = part(
  'Url',
  defineComponent({
    name: 'PptxEditorHyperLinkUrl',
    props: partProps,
    setup(props, { slots }) {
      const { state, openTarget } = useHyperlinkPopup();
      const { t } = useTranslation();
      return () => {
        const s = state.value;
        if (props.hidden || !s.url) return null;
        const children = slots.default?.();
        const slide = slideLinkLabel(t, s.url);
        if (slide) {
          const attrs = {
            type: 'button' as const,
            class: ['pptx-link__url pptx-link__slide', props.className],
            'aria-label': t('link.goToTarget', { target: slide }),
            title: t('link.goToTarget', { target: slide }),
            onMousedown: guardMousedown,
            onClick: () => void openTarget(),
          };
          return wire(props.asChild, children, attrs, () => (
            <button {...attrs}>{children ?? slide}</button>
          ));
        }
        const text = () => <bdi dir="ltr">{s.url}</bdi>;
        if (!s.href) {
          const attrs = {
            class: ['pptx-link__url', props.className],
            'data-inert': '',
            title: t('link.unsafe'),
          };
          return wire(props.asChild, children, attrs, () => (
            <span {...attrs}>{children ?? text()}</span>
          ));
        }
        const attrs = {
          class: ['pptx-link__url', props.className],
          href: s.href,
          target: '_blank',
          rel: 'noopener noreferrer',
          title: t('link.open'),
          onMousedown: guardMousedown,
        };
        return wire(props.asChild, children, attrs, () => (
          <a {...attrs}>
            {children ?? text()}
            <Icon name="open_in_new" />
          </a>
        ));
      };
    },
  })
);

/** An icon button of the reading row. */
function actionPart(
  name: 'Copy' | 'Edit' | 'Unlink',
  cls: string,
  glyph: string,
  use: () => { label: () => string; run: () => void; shown?: () => boolean; extra?: () => object }
) {
  return part(
    name,
    defineComponent({
      name: `PptxEditorHyperLink${name}`,
      props: partProps,
      setup(props, { slots }) {
        const action = use();
        return () => {
          if (props.hidden || action.shown?.() === false) return null;
          const label = action.label();
          const attrs = {
            type: 'button' as const,
            class: [`pptx-link__action pptx-link__${cls}`, props.className],
            'aria-label': label,
            title: label,
            onMousedown: guardMousedown,
            onClick: action.run,
            ...action.extra?.(),
          };
          const children = slots.default?.();
          return wire(props.asChild, children, attrs, () => (
            <button {...attrs}>{slots.icon?.() ?? children ?? <Icon name={glyph} />}</button>
          ));
        };
      },
    })
  );
}

const HyperLinkCopy = actionPart('Copy', 'copy', 'content_copy', () => {
  const { state, copy } = useHyperlinkPopup();
  const { t } = useTranslation();
  return {
    label: () => t(state.value.copied ? 'link.copied' : 'link.copy'),
    run: () => void copy(),
    // A slide link names a place in this deck: there is nothing to paste elsewhere.
    shown: () => !!state.value.href && !state.value.href.startsWith('#'),
    extra: () => (state.value.copied ? { 'data-copied': '' } : {}),
  };
});

const HyperLinkEdit = actionPart('Edit', 'edit', 'edit', () => {
  const { beginEdit } = useHyperlinkPopup();
  const { t } = useTranslation();
  return { label: () => t('link.edit'), run: beginEdit };
});

const HyperLinkUnlink = actionPart('Unlink', 'unlink', 'link_off', () => {
  const { unlink } = useHyperlinkPopup();
  const { t } = useTranslation();
  return { label: () => t('link.remove'), run: () => void unlink() };
});

const HyperLinkFields = part(
  'Fields',
  defineComponent({
    name: 'PptxEditorHyperLinkFields',
    props: partProps,
    setup(props) {
      const provision = inject(hyperlinkPopupKey, null);
      const { state, setUrl, commitEdit } = useHyperlinkPopup();
      const { t } = useTranslation();
      const slides = useEditorState((s) => s.slides.length);
      const kind = ref<LinkKind>(linkKindOf(state.value.draft));
      const mail = shallowRef(mailtoParts(state.value.draft));
      const field = ref<HTMLInputElement | HTMLSelectElement | null>(null);
      // After the popover is placed: a browser does not focus into a hidden panel. A change of
      // kind keeps focus on Link to, so the arrow keys can step through the kinds.
      onMounted(() => {
        field.value?.focus({ preventScroll: true });
        if (field.value instanceof HTMLInputElement) field.value.select();
      });
      const onEnter = (e: KeyboardEvent): void => {
        if (e.key !== 'Enter') return;
        // The Enter that applies must not reach the element focus moves to next.
        e.preventDefault();
        commitEdit();
      };
      const setMail = (next: MailtoParts): void => {
        mail.value = next;
        setUrl(mailtoUrl(next.address, next.subject, next.rest));
      };
      const KIND_LABEL = {
        web: 'link.kindWeb',
        email: 'link.kindEmail',
        slide: 'link.kindSlide',
      } as const;
      return () => {
        if (props.hidden) return null;
        const s = state.value;
        const invalid = {
          'aria-invalid': s.error ? true : undefined,
          'aria-describedby': s.error && provision ? provision.errorId : undefined,
        };
        const value = (e: Event): string => (e.target as HTMLInputElement).value;
        return (
          <div class={['pptx-link__fields', props.className]}>
            <label class="pptx-field pptx-link__kind">
              {t('link.kind')}
              <select
                value={kind.value}
                onChange={(e: Event) => {
                  const next = value(e) as LinkKind;
                  kind.value = next;
                  mail.value = { address: '', subject: '', rest: '' };
                  setUrl(next === 'slide' ? '#next-slide' : '');
                }}
              >
                {LINK_KINDS.map((k) => (
                  <option key={k} value={k}>
                    {t(KIND_LABEL[k])}
                  </option>
                ))}
              </select>
            </label>
            {kind.value === 'web' ? (
              <input
                ref={field}
                class="pptx-font-picker__search pptx-link__field"
                type="url"
                dir="ltr"
                aria-label={t('link.label')}
                placeholder={t('link.placeholder')}
                value={s.draft}
                {...invalid}
                onInput={(e: Event) => setUrl(value(e))}
                onKeydown={onEnter}
              />
            ) : kind.value === 'email' ? (
              <>
                <input
                  ref={field}
                  class="pptx-font-picker__search pptx-link__field"
                  type="email"
                  dir="ltr"
                  aria-label={t('link.emailAddress')}
                  placeholder={t('link.emailAddress')}
                  value={mail.value.address}
                  {...invalid}
                  onInput={(e: Event) => setMail({ ...mail.value, address: value(e) })}
                  onKeydown={onEnter}
                />
                <input
                  class="pptx-font-picker__search pptx-link__field"
                  type="text"
                  aria-label={t('link.emailSubject')}
                  placeholder={t('link.emailSubject')}
                  value={mail.value.subject}
                  onInput={(e: Event) => setMail({ ...mail.value, subject: value(e) })}
                  onKeydown={onEnter}
                />
              </>
            ) : (
              <select
                ref={field}
                class="pptx-link__field"
                aria-label={t('link.slide')}
                value={s.draft}
                {...invalid}
                onChange={(e: Event) => setUrl(value(e))}
                onKeydown={onEnter}
              >
                {slideLinkChoices(slides.value).map((url) => (
                  <option key={url} value={url}>
                    {slideLinkLabel(t, url)}
                  </option>
                ))}
              </select>
            )}
          </div>
        );
      };
    },
  })
);

const HyperLinkError = part(
  'Error',
  defineComponent({
    name: 'PptxEditorHyperLinkError',
    props: partProps,
    setup(props) {
      const provision = inject(hyperlinkPopupKey, null);
      const { t } = useTranslation();
      return () => {
        const error = provision?.state.value.error;
        if (props.hidden || !error) return null;
        return (
          <p id={provision.errorId} role="alert" class={['pptx-link__error', props.className]}>
            {translateKeyOrText(t, error)}
          </p>
        );
      };
    },
  })
);

/** A text button of the edit row. */
function textPart(
  name: 'Apply' | 'Cancel',
  cls: string,
  labelKey: 'link.apply' | 'link.cancel',
  primary: boolean
) {
  return part(
    name,
    defineComponent({
      name: `PptxEditorHyperLink${name}`,
      props: partProps,
      setup(props, { slots }) {
        const { state, commitEdit, close } = useHyperlinkPopup();
        const { t } = useTranslation();
        return () => {
          if (props.hidden) return null;
          const attrs = {
            type: 'button' as const,
            class: [
              primary ? 'pptx-button pptx-button--primary' : 'pptx-button',
              cls,
              props.className,
            ],
            disabled: primary ? state.value.draft.trim().length === 0 : undefined,
            onMousedown: guardMousedown,
            onClick: primary ? () => void commitEdit() : close,
          };
          const children = slots.default?.();
          return wire(props.asChild, children, attrs, () => (
            <button {...attrs}>{children ?? t(labelKey)}</button>
          ));
        };
      },
    })
  );
}

const HyperLinkApply = textPart('Apply', 'pptx-link__apply', 'link.apply', true);
const HyperLinkCancel = textPart('Cancel', 'pptx-link__cancel', 'link.cancel', false);

/** Each part child by its name, so the preset swaps it in place; other children follow. */
function partOverrides(children: VNode[] | undefined): {
  parts: Partial<Record<PartName, VNode>>;
  extra: VNode[];
} {
  const parts: Partial<Record<PartName, VNode>> = {};
  const extra: VNode[] = [];
  for (const node of flatten(children)) {
    const name = (node.type as { pptxHyperLinkPart?: PartName } | null)?.pptxHyperLinkPart;
    if (name) parts[name] = node;
    else extra.push(node);
  }
  return { parts, extra };
}

/** The default arrangement of the mode, with host parts in place. */
function preset(state: HyperlinkPopupState, children: VNode[] | undefined): VNodeArrayChildren {
  const { parts, extra } = partOverrides(children);
  const take = (name: PartName, fallback: () => VNode): VNode | null =>
    name in parts ? (parts[name] ?? null) : fallback();
  if (state.mode === 'editing')
    return [
      // An existing link keeps its target, Copy, and Open while it is edited.
      state.url ? (
        <div class="pptx-link__row">
          {take('Url', () => (
            <HyperLinkUrl />
          ))}
          <div class="pptx-link__tools">
            {take('Copy', () => (
              <HyperLinkCopy />
            ))}
          </div>
        </div>
      ) : null,
      take('Fields', () => <HyperLinkFields />),
      take('Error', () => <HyperLinkError />),
      <div class="pptx-link__actions">
        {state.url && state.canEdit ? take('Unlink', () => <HyperLinkUnlink />) : null}
        {take('Cancel', () => (
          <HyperLinkCancel />
        ))}
        {take('Apply', () => (
          <HyperLinkApply />
        ))}
      </div>,
      ...extra,
    ];
  return [
    <div class="pptx-link__row">
      {take('Url', () => (
        <HyperLinkUrl />
      ))}
      <div class="pptx-link__tools">
        {take('Copy', () => (
          <HyperLinkCopy />
        ))}
        {/* An action that can never run is absent, not disabled: a read-only editor. */}
        {state.canEdit ? take('Edit', () => <HyperLinkEdit />) : null}
        {state.canEdit ? take('Unlink', () => <HyperLinkUnlink />) : null}
      </div>
    </div>,
    take('Error', () => <HyperLinkError />),
    ...extra,
  ];
}

const HyperLinkRoot = defineComponent({
  name: 'PptxEditorHyperLink',
  inheritAttrs: false,
  props: rootProps,
  setup(props, { slots, attrs }) {
    const provision = inject(hyperlinkPopupKey, null);
    const { t } = useTranslation();
    const editorRef = usePptxEditor();
    const selectionRect = shallowRef<ReturnType<typeof linkSelectionRect>>(null);
    watch(
      () => [editorRef.value, provision?.state.value.mode, provision?.state.value.anchor] as const,
      ([editor, mode, anchor], previous) => {
        if (!mode || mode === 'closed') {
          selectionRect.value = null;
        } else if (
          !previous ||
          previous[1] === 'closed' ||
          previous[0] !== editor ||
          previous[2] !== anchor
        ) {
          // Keep the selection box while the field takes focus and its draft changes.
          selectionRect.value = editor ? linkSelectionRect(editor) : null;
        }
      },
      { immediate: true, flush: 'sync' }
    );
    return () => {
      const state = provision?.state.value ?? CLOSED_STATE;
      if (!provision || state.mode === 'closed') return null;
      const children = slots.default?.();
      // `hidden` drops the packaged panel only: the gestures stay, so a host renders its own UI.
      if (props.hidden) return children ?? null;
      const reading = state.mode === 'reading';
      const inner = { class: 'pptx-link', 'data-mode': state.mode };
      const body = props.preset ? preset(state, children) : children;
      const extra = [props.className, attrs.class as string | undefined].filter(Boolean).join(' ');
      return (
        <Popover
          anchor={state.anchor}
          anchorRect={selectionRect.value}
          role="dialog"
          label={t(reading ? 'link.label' : state.url ? 'link.edit' : 'link.insert')}
          onClose={() => provision.controller.close()}
          // Reading leaves the keys in the slide text; editing takes them in its field.
          passive={reading}
          field={!reading}
          className={extra ? `pptx-link-popup ${extra}` : 'pptx-link-popup'}
        >
          {{
            default: () => wire(props.asChild, children, inner, () => <div {...inner}>{body}</div>),
          }}
        </Popover>
      );
    };
  },
});

/** `PptxEditorHyperLink`: the link popup with named parts, the twin of React's. @public */
export const PptxEditorHyperLink = Object.assign(HyperLinkRoot, {
  Url: HyperLinkUrl,
  Copy: HyperLinkCopy,
  Edit: HyperLinkEdit,
  Unlink: HyperLinkUnlink,
  Fields: HyperLinkFields,
  Error: HyperLinkError,
  Apply: HyperLinkApply,
  Cancel: HyperLinkCancel,
});
