// Ported from DOCX Editor packages/vue/src/editor/dialog-parts.ts at 27720ad5
// (`createDialogComposition`, `NativeDialog`). PPTX changes:
// the names follow the React twin (`createDialogParts`, `DialogFrame`), and the draft lives in
// core.
//
// Dialog named parts over a draft controller: the Vue twin of the React adapter's
// `dialog-parts.tsx`. A packaged dialog renders its default
// arrangement with each element marked `data-pptx-part`; a host part with the same name
// replaces it in place, keeps its wiring, and adds the host's class, style, and attributes.
// The draft lives in core (`createDialogDraft`), so both adapters apply the same way.

import {
  Fragment,
  cloneVNode,
  defineComponent,
  h,
  inject,
  isVNode,
  onBeforeUnmount,
  onMounted,
  provide,
  ref,
  shallowRef,
  watch,
  type CSSProperties,
  type DefineComponent,
  type InjectionKey,
  type PropType,
  type Ref,
  type ShallowRef,
  type VNode,
  type VNodeArrayChildren,
} from 'vue';
import {
  DIALOG_FIELD_ATTRIBUTE,
  DIALOG_PART_ATTRIBUTE,
  createDialogDraft,
  dialogInitialFocus,
  dialogPartKey,
  dialogPartLabel,
  fromNestedDialog,
  isDialogBackdropPoint,
  returnAppliesFrom,
  trapTabWithin,
  warnDialogPart,
  type DialogDraftOptions,
  type DialogErrors,
} from '@pptx-editor.dev/core/editor';
import { usePptxEditor } from './composables.ts';
import { useChromeDark } from './color-mode.ts';

/** A dialog's draft and actions, as its parts and a host's own fields read them. @public */
export interface UseDialogReturn<Fields extends object> {
  /** The draft values. */
  readonly values: Readonly<Ref<Fields>>;
  /** Translated messages per field, and `form` for the whole draft. */
  readonly errors: Readonly<
    Ref<Readonly<Partial<Record<Extract<keyof Fields, string> | 'form', string>>>>
  >;
  /** False when the dialog cannot apply now (no document, viewing mode). */
  readonly isEnabled: Readonly<Ref<boolean>>;
  /** Change one field of the draft. */
  setValue<K extends keyof Fields>(name: K, value: Fields[K]): void;
  /** Validate and commit the draft. The dialog closes when the document took it. */
  apply(): void;
  /** Discard the draft and close. */
  cancel(): void;
}

/** A draft as Vue refs, re-read on every change. @internal */
export function useDialogDraft<Fields extends object>(
  options: DialogDraftOptions<Fields>
): {
  readonly values: ShallowRef<Fields>;
  readonly errors: ShallowRef<DialogErrors<Fields>>;
  setValue<K extends keyof Fields>(name: K, value: Fields[K]): void;
  apply(): boolean;
  reset(): void;
} {
  const draft = createDialogDraft(options);
  const values = shallowRef(draft.values) as ShallowRef<Fields>;
  const errors = shallowRef(draft.errors) as ShallowRef<DialogErrors<Fields>>;
  const off = draft.subscribe(() => {
    values.value = draft.values;
    errors.value = draft.errors;
  });
  onBeforeUnmount(off);
  return {
    values,
    errors,
    setValue: draft.setValue,
    apply: draft.apply,
    reset: draft.reset,
  };
}

/**
 * Presentation props of a dialog part, the twin of React's `DialogPartProps`. Other attributes
 * (`id`, `aria-*`, `data-*`) reach the rendered element; the part's own wiring wins every
 * collision. @public
 */
export interface DialogPartProps {
  /** A class name added to the part (`class` works too). */
  readonly className?: string;
  /** Styles added to the part. */
  readonly style?: CSSProperties;
  /** Remove the part from the arrangement. */
  readonly hidden?: boolean;
  /** Apply the part's wiring to the single child element. */
  readonly asChild?: boolean;
  /** The id of the part's element. */
  readonly id?: string;
  /** The tooltip of the part's element. */
  readonly title?: string;
  /** The accessible name of the part's element. */
  readonly 'aria-label'?: string;
  /** `data-*` attributes, passed to the part's element. */
  readonly [attribute: `data-${string}`]: unknown;
}

/**
 * Layout customization for a packaged dialog, the twin of React's `DialogCustomizationProps`
 * and DOCX Editor's Vue one. Vue passes children through the default slot, not a prop. @public
 */
export interface DialogCustomizationProps {
  /** A class name added to the dialog surface (`class` works too). */
  readonly className?: string;
  /** Styles added to the dialog surface. */
  readonly style?: CSSProperties;
  /** `true` (default): the default arrangement with host parts in place. `false`: only the slot. */
  readonly preset?: boolean;
}

/** The parts a `createDialogParts` call returns. @internal */
export interface DialogParts<State, Name extends string> {
  readonly Header: DefineComponent<DialogPartProps>;
  readonly Title: DefineComponent<DialogPartProps>;
  readonly Body: DefineComponent<DialogPartProps>;
  readonly Footer: DefineComponent<DialogPartProps>;
  readonly Apply: DefineComponent<DialogPartProps>;
  readonly Cancel: DefineComponent<DialogPartProps>;
  readonly Reset: DefineComponent<DialogPartProps>;
  readonly Error: DefineComponent<DialogPartProps>;
  /** One field of the dialog, by its field name. */
  readonly Field: DefineComponent<DialogPartProps & { readonly name: Name }>;
  /** The dialog's draft, for a child of the dialog. Throws outside it. */
  readonly useState: () => State;
  /** A dialog-specific part, matched to the default element with this `data-pptx-part`. */
  readonly part: (name: string) => DefineComponent<DialogPartProps>;
  /** Provide the state; returns the renderer of the default arrangement with host parts. */
  readonly provideParts: (
    state: State
  ) => (
    defaults: VNode,
    children: VNodeArrayChildren | undefined,
    preset: boolean
  ) => VNodeArrayChildren;
}

const partProps = {
  /** A class name added to the part. */
  className: { type: String, default: undefined },
  /** Styles added to the part. */
  style: { type: Object as PropType<CSSProperties>, default: undefined },
  /** Remove the part from the arrangement. */
  hidden: { type: Boolean, default: false },
  /** Apply the part's wiring to the single child element. */
  asChild: { type: Boolean, default: false },
} as const;

interface Composition {
  readonly defaults: ShallowRef<Map<string, VNode>>;
  readonly overrides: Map<string, VNode>;
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

/** A vnode's part key, read from its markers. */
const markerOf = (node: VNode): string | null =>
  dialogPartKey(node.props?.[DIALOG_PART_ATTRIBUTE], node.props?.[DIALOG_FIELD_ATTRIBUTE]);

/** The part a host component stands for, read from its static marker. */
const partOf = (node: VNode): string | null => {
  const part = (node.type as { pptxDialogPart?: string } | null)?.pptxDialogPart;
  return part === undefined ? null : dialogPartKey(part, node.props?.name);
};

/**
 * Build the named parts of one dialog: `Header`, `Title`, `Body`, `Footer`, `Apply`, `Cancel`,
 * `Reset`, `Error`, and `Field` (by its `name`). The dialog calls `provideParts(state)` in its
 * setup and renders through the function it returns. @internal
 */
export function createDialogParts<State, Name extends string = string>(
  dialogName: string
): DialogParts<State, Name> {
  const stateKey: InjectionKey<State> = Symbol(dialogName);
  const compositionKey: InjectionKey<Composition> = Symbol(`${dialogName}.parts`);
  const label = (key: string): string => dialogPartLabel(dialogName, key);
  function useState(): State {
    const state = inject(stateKey, null);
    if (!state) throw new Error(`${dialogName} parts must be rendered inside ${dialogName}.`);
    return state;
  }
  function makePart(part: string): DefineComponent<DialogPartProps> {
    const component = Object.assign(
      defineComponent({
        name: `${dialogName}${part[0]!.toUpperCase()}${part.slice(1)}`,
        inheritAttrs: false,
        props: {
          ...partProps,
          /** The field this part stands for. `Field` parts only. */
          name: { type: String, default: undefined },
        },
        setup(props, { attrs, slots }) {
          const composition = inject(compositionKey, null);
          return () => {
            const key = dialogPartKey(part, props.name) ?? '';
            if (!composition) {
              warnDialogPart(label(key), 'outside');
              return null;
            }
            if (props.hidden) return null;
            const children = slots.default?.();
            const original = composition.defaults.value.get(key);
            if (!original) {
              warnDialogPart(label(key), children?.length ? 'childrenOnly' : 'unknown');
              return children ?? null;
            }
            // The part's own wiring outranks the host's attributes, so a stray `onClick` or
            // `data-pptx-part` cannot detach the control from the dialog. The React twin
            // applies the same order.
            // A default element's `undefined` (a `title` set only when disabled) leaves the
            // host's value in place, as the React twin.
            const wiring: Record<string, unknown> = {
              ...attrs,
              ...Object.fromEntries(
                Object.entries(original.props ?? {}).filter(([, value]) => value !== undefined)
              ),
            };
            // `asChild` needs exactly one element: text and comments have symbol types.
            const only = props.asChild ? flatten(children) : [];
            const child =
              only.length === 1 && typeof only[0]!.type !== 'symbol' ? only[0] : undefined;
            if (props.asChild && !child) warnDialogPart(label(key), 'asChild');
            if (child) {
              // The host owns the look: no packaged `pptx-dialog__*` classes or styles.
              delete wiring.class;
              delete wiring.style;
              return cloneVNode(child, {
                ...wiring,
                class: [attrs.class, props.className],
                style: props.style,
              });
            }
            const merged = {
              ...wiring,
              class: [original.props?.class, attrs.class, props.className],
              style: [original.props?.style, props.style],
            };
            // A misused `asChild` falls back to the packaged part with its packaged content.
            const content = props.asChild ? undefined : children;
            return h(
              original.type as string,
              merged,
              content ?? (original.children as VNodeArrayChildren | string | null) ?? undefined
            );
          };
        },
      }),
      { pptxDialogPart: part }
    );
    return component as unknown as DefineComponent<DialogPartProps>;
  }
  /**
   * Provide the dialog's state to its parts. Returns the renderer: the default arrangement with
   * host parts in place (`preset`), or only the host's children.
   */
  function provideParts(state: State) {
    provide(stateKey, state);
    const composition: Composition = { defaults: shallowRef(new Map()), overrides: new Map() };
    provide(compositionKey, composition);
    return (
      defaults: VNode,
      children: VNodeArrayChildren | undefined,
      preset: boolean
    ): VNodeArrayChildren => {
      composition.overrides.clear();
      const extras: VNode[] = [];
      for (const child of flatten(children)) {
        const key = partOf(child);
        if (!key) {
          extras.push(child);
          // Every level down, as the React twin.
          const scan = (node: VNode): void => {
            if (!Array.isArray(node.children)) return;
            for (const inner of flatten(node.children as VNodeArrayChildren)) {
              const nested = partOf(inner);
              if (nested) warnDialogPart(label(nested), 'nested');
              scan(inner);
            }
          };
          if (preset) scan(child);
        } else {
          if (composition.overrides.has(key)) warnDialogPart(label(key), 'duplicate');
          composition.overrides.set(key, child);
        }
      }
      const map = new Map<string, VNode>();
      // Text children stay as they are: a row such as "Widescreen <span>13.33 × 7.50 in</span>"
      // keeps its text beside its elements.
      const visitChildren = (children: VNodeArrayChildren): VNodeArrayChildren =>
        children.map((child) =>
          Array.isArray(child) ? visitChildren(child) : isVNode(child) ? visit(child) : child
        );
      const visit = (node: VNode): VNode => {
        const copy = cloneVNode(node);
        if (Array.isArray(node.children))
          copy.children = visitChildren(node.children as VNodeArrayChildren);
        const key = markerOf(node);
        if (key) {
          map.set(key, copy);
          if (preset && composition.overrides.has(key)) return composition.overrides.get(key)!;
        }
        return copy;
      };
      const rendered = visit(defaults);
      // A part with no default in place is never rendered, so it says so here.
      if (preset)
        for (const [key, node] of composition.overrides)
          if (!map.has(key) && [undefined, false].includes(node.props?.hidden))
            warnDialogPart(label(key), 'unknown');
      // A part whose default sits inside another part that brings its own children is dropped
      // with that default, as in the React twin.
      if (preset)
        for (const [key, node] of composition.overrides) {
          if (node.children == null) continue;
          // As the React twin: a misused `asChild` falls back to the packaged part.
          const asChild = node.props?.asChild ?? node.props?.['as-child'];
          if (asChild !== undefined && asChild !== false) {
            // The part's own test: exactly one node, and it is an element.
            const slot = (node.children as { default?: () => VNodeArrayChildren }).default;
            const nodes = flatten(slot?.() ?? []);
            if (nodes.length !== 1 || typeof nodes[0]!.type === 'symbol') continue;
          }
          const scan = (inner: VNode): void => {
            if (!Array.isArray(inner.children)) return;
            for (const child of flatten(inner.children as VNodeArrayChildren)) {
              // The copy already holds the override in place of the nested default.
              const nested = markerOf(child) ?? partOf(child);
              if (nested && nested !== key && composition.overrides.has(nested))
                warnDialogPart(label(nested), 'shadowed', label(key));
              scan(child);
            }
          };
          const owner = map.get(key);
          if (owner) scan(owner);
        }
      composition.defaults.value = map;
      return preset ? [rendered, ...extras] : (children ?? []);
    };
  }
  return {
    Header: makePart('header'),
    Title: makePart('title'),
    Body: makePart('body'),
    Footer: makePart('footer'),
    Apply: makePart('apply'),
    Cancel: makePart('cancel'),
    Reset: makePart('reset'),
    Error: makePart('error'),
    Field: makePart('field') as unknown as DefineComponent<
      DialogPartProps & { readonly name: Name }
    >,
    useState,
    part: (name) => makePart(name),
    provideParts,
  };
}

/**
 * The native modal every packaged dialog renders in, the twin of React's `DialogFrame`: it
 * opens as a modal on mount and closes on unmount, focuses its first control, keeps Tab inside,
 * closes on Escape and on a click on the backdrop, and gives focus back. Outside an editor
 * scope it carries the scope itself, so the tokens apply. @internal
 */
export const DialogFrame = defineComponent({
  name: 'PptxDialogFrame',
  props: {
    kind: { type: String, required: true },
    label: { type: String, required: true },
    onClose: { type: Function as PropType<() => void>, required: true },
    onKeydown: { type: Function as PropType<(event: KeyboardEvent) => void>, default: undefined },
  },
  setup(props, { slots }) {
    const panel = ref<HTMLDialogElement | null>(null);
    const ownScope = ref(false);
    const dark = useChromeDark();
    let opener: HTMLElement | null = null;
    // A press inside the panel that is released over the backdrop is a drag, not a dismissal.
    let pressedOutside = false;
    onMounted(() => {
      const el = panel.value;
      if (!el) return;
      ownScope.value = !el.parentElement?.closest('.pptx-editor');
      opener = el.ownerDocument.activeElement as HTMLElement | null;
      if (typeof el.showModal === 'function' && !el.open) el.showModal();
      else el.setAttribute('open', '');
      dialogInitialFocus(el)?.focus({ preventScroll: true });
    });
    onBeforeUnmount(() => {
      const el = panel.value;
      if (el && typeof el.close === 'function' && el.open) el.close();
      if (opener?.isConnected) opener.focus({ preventScroll: true });
    });
    return () =>
      h(
        'dialog',
        {
          ref: panel,
          'aria-modal': 'true',
          'aria-label': props.label,
          class: [
            'pptx-dialog',
            ownScope.value ? 'pptx-editor' : null,
            ownScope.value && dark.value ? 'dark' : null,
          ],
          'data-pptx-dialog': props.kind,
          onCancel: (event: Event) => {
            // A nested dialog's own Escape closes only that dialog, as in the React twin.
            if (fromNestedDialog(event.currentTarget as Element, event.target)) return;
            event.preventDefault();
            props.onClose();
          },
          onMousedown: (event: MouseEvent) => {
            pressedOutside = isDialogBackdropPoint(event.currentTarget as HTMLElement, event);
            event.stopPropagation();
          },
          onClick: (event: MouseEvent) => {
            // As DOCX Editor and the React twin: a press and a release outside the dialog's box.
            if (pressedOutside && isDialogBackdropPoint(event.currentTarget as HTMLElement, event))
              props.onClose();
            pressedOutside = false;
            event.stopPropagation();
          },
          onKeydown: (event: KeyboardEvent) => {
            // A dialog opened inside this one handles its own keys; they still stop here.
            if (fromNestedDialog(event.currentTarget as Element, event.target)) {
              event.stopPropagation();
              return;
            }
            if (event.key === 'Escape' && !event.isComposing) {
              event.preventDefault();
              props.onClose();
            } else if (trapTabWithin(event.currentTarget as HTMLElement, event))
              event.preventDefault();
            else if (!event.isComposing && !event.defaultPrevented) props.onKeydown?.(event);
            event.stopPropagation();
          },
        },
        slots.default?.()
      );
  },
});

/**
 * Return in a field applies the draft, as the React twin: from any input but a file input,
 * and from a radio of a radio group. @internal
 */
export function applyOnEnter(event: KeyboardEvent, enabled: boolean, apply: () => void): void {
  if (event.key !== 'Enter' || !enabled || !returnAppliesFrom(event.target)) return;
  event.preventDefault();
  apply();
}

/**
 * Close a dialog when its editor opens another document, the twin of React's
 * `useCloseOnNewDocument` (DOCX Editor's `useDialogGeneration`). @internal
 */
export function useCloseOnNewDocument(onClose: () => void): void {
  const editor = usePptxEditor();
  watch(
    editor,
    (current, _old, onCleanup) => {
      if (!current) return;
      let session = current.session;
      // Once per new document; `destroy()` leaves no session and closes nothing.
      onCleanup(
        current.subscribe(() => {
          if (current.session === session) return;
          session = current.session;
          if (session) onClose();
        })
      );
    },
    { immediate: true }
  );
}
