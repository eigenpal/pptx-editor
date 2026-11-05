/** @jsxImportSource react */
// Ported from DOCX Editor packages/react/src/editor/dialog-parts.tsx at 27720ad5
// (`createDialogParts`, `DialogFrame`, `UseDialogReturn`).
// PPTX changes: part markers are `data-pptx-*`, the draft lives in core, `createDialogParts`
// takes the dialog's name for warnings, and the frame rules in core differ where noted there.
//
// Dialog named parts over a draft controller. A packaged dialog renders its default arrangement with each element
// marked `data-pptx-part` (and `data-pptx-field` for a field). A host child with the same part
// name replaces that element in place: it keeps the part's wiring and adds its own class,
// style, and attributes, or takes the wiring onto its own element with `asChild`.
// `preset={false}` renders only the host's children. The draft itself lives in core
// (`createDialogDraft`), so the Vue twin behaves the same.

import {
  Children,
  Fragment,
  cloneElement,
  createContext,
  isValidElement,
  useContext,
  useEffect,
  useLayoutEffect,
  useRef,
  useState,
  useSyncExternalStore,
  type CSSProperties,
  type KeyboardEvent,
  type ReactElement,
  type ReactNode,
} from 'react';
import {
  DIALOG_FIELD_ATTRIBUTE,
  DIALOG_PART_ATTRIBUTE,
  createDialogDraft,
  dialogInitialFocus as initialFocus,
  dialogPartKey,
  dialogPartLabel,
  fromNestedDialog,
  isDialogBackdropPoint,
  returnAppliesFrom,
  trapTabWithin,
  warnDialogPart,
  type DialogDraft,
  type DialogDraftOptions,
  type DialogErrors,
} from '@pptx-editor.dev/core/editor';
import { renderAsChild, type HostElementProps } from './slot-override.ts';
import { useChromeScopeClassName } from './color-mode.ts';
import { usePptxEditor } from './context.ts';

/**
 * Presentation overrides for a dialog part. @public
 *
 * Anything beyond these members reaches the rendered element, so a part takes an `id`, an
 * `aria-*` label, or a `data-*` test hook. The part's own wiring wins every collision: a host
 * cannot replace Apply's handler or its `data-pptx-part` marker by passing one. To own the
 * element and its handler outright, pass `asChild` and supply your own.
 */
export interface DialogPartProps {
  /** A class name added to the part's element. */
  readonly className?: string;
  /** Inline styles for the part's element. */
  readonly style?: CSSProperties;
  /** Remove the part from the arrangement. */
  readonly hidden?: boolean;
  /** Apply the part's wiring to the single child element. */
  readonly asChild?: boolean;
  /** Content in place of the part's default content. */
  readonly children?: ReactNode;
  /** The id of the part's element. */
  readonly id?: string;
  /** The tooltip of the part's element. */
  readonly title?: string;
  /** The accessible name of the part's element. */
  readonly 'aria-label'?: string;
  /** `data-*` attributes, passed to the part's element. */
  readonly [attribute: `data-${string}`]: unknown;
}

/** Layout customization for a packaged dialog. @public */
export interface DialogCustomizationProps {
  /** A class name added to the dialog surface. */
  readonly className?: string;
  /** Styles added to the dialog surface. */
  readonly style?: CSSProperties;
  /**
   * `true` (default): the default arrangement, with named children replacing their parts in
   * place and other children appended. `false`: only the children.
   */
  readonly preset?: boolean;
  /** Named parts that replace the defaults, or the whole arrangement with `preset={false}`. */
  readonly children?: ReactNode;
}

/** A dialog's draft and actions, as its parts and a host's own fields read them. @public */
export interface UseDialogReturn<Fields extends object> {
  /** The draft values. */
  readonly values: Fields;
  /** Change one field of the draft. */
  setValue<K extends keyof Fields>(name: K, value: Fields[K]): void;
  /** Translated messages per field, and `form` for the whole draft. */
  readonly errors: Readonly<Partial<Record<Extract<keyof Fields, string> | 'form', string>>>;
  /** False when the dialog cannot apply now (no document, viewing mode). */
  readonly isEnabled: boolean;
  /** Validate and commit the draft. The dialog closes when the document took it. */
  apply(): void;
  /** Discard the draft and close. */
  cancel(): void;
}

type NodeProps = Record<string, unknown> & { children?: ReactNode };

/** The part key of a default element, read from its markers. */
const markerOf = (props: NodeProps): string | null =>
  dialogPartKey(props[DIALOG_PART_ATTRIBUTE], props[DIALOG_FIELD_ATTRIBUTE]);

/** Every element in `nodes`, depth first, through fragments. */
function visit(nodes: ReactNode, callback: (node: ReactElement<NodeProps>) => void): void {
  Children.forEach(nodes, (node) => {
    if (!isValidElement<NodeProps>(node)) return;
    callback(node);
    visit(node.props.children, callback);
  });
}

/** The top-level nodes of `nodes`, with fragments opened. */
const flatten = (nodes: ReactNode): ReactNode[] =>
  Children.toArray(nodes).flatMap((node) =>
    isValidElement<NodeProps>(node) && node.type === Fragment
      ? flatten(node.props.children)
      : [node]
  );

/** The parts a `createDialogParts` call returns. @internal */
export interface DialogParts<Name extends string, State> {
  readonly Header: (props: DialogPartProps) => ReactNode;
  readonly Title: (props: DialogPartProps) => ReactNode;
  readonly Body: (props: DialogPartProps) => ReactNode;
  readonly Footer: (props: DialogPartProps) => ReactNode;
  readonly Apply: (props: DialogPartProps) => ReactNode;
  readonly Cancel: (props: DialogPartProps) => ReactNode;
  readonly Reset: (props: DialogPartProps) => ReactNode;
  readonly Error: (props: DialogPartProps) => ReactNode;
  /** One field of the dialog, by its field name. */
  readonly Field: (props: DialogPartProps & { readonly name: Name }) => ReactNode;
  /** Renders the default arrangement with the host's parts in place. Internal to dialogs. */
  readonly Composition: (props: {
    readonly defaults: ReactNode;
    readonly children?: ReactNode;
    readonly preset?: boolean;
    readonly state: State;
  }) => ReactNode;
  /** The dialog's draft, for a child of the dialog. Throws outside it. */
  readonly useState: () => State;
  /** A dialog-specific part, matched to the default element with this `data-pptx-part`. */
  readonly part: (name: string) => (props: DialogPartProps) => ReactNode;
}

/**
 * Build the named parts of one dialog. `Name` is the dialog's field names; `State` is what
 * `useState()` returns to a host's own controls. `dialogName` names the dialog in errors and
 * development warnings. The document state stays in the dialog's draft controller; the parts
 * only render. Every packaged dialog is built with it. @internal
 */
export function createDialogParts<Name extends string, State>(
  dialogName = 'Dialog'
): DialogParts<Name, State> {
  const Context = createContext<{
    defaults: Map<string, ReactElement<NodeProps>>;
    overrides: Map<string, ReactElement<DialogPartProps>>;
    state: State;
  } | null>(null);
  const identities = new Map<unknown, string>();
  const keyOf = (node: ReactElement<NodeProps>): string | null => {
    const part = identities.get(node.type);
    return part === undefined ? null : dialogPartKey(part, node.props.name);
  };
  const label = (key: string): string => dialogPartLabel(dialogName, key);
  function useDialogState(): State {
    const context = useContext(Context);
    if (!context)
      throw new globalThis.Error(`${dialogName} parts must be rendered inside ${dialogName}.`);
    return context.state;
  }
  function replace(
    nodes: ReactNode,
    overrides: Map<string, ReactElement<DialogPartProps>>
  ): ReactNode {
    return Children.map(nodes, (node) => {
      if (!isValidElement<NodeProps>(node)) return node;
      const key = markerOf(node.props);
      if (key && overrides.has(key)) return overrides.get(key);
      return node.props.children === undefined
        ? node
        : cloneElement(node, { children: replace(node.props.children, overrides) });
    });
  }
  function makePart(part: string) {
    function Part(props: DialogPartProps & { name?: Name }): ReactNode {
      const context = useContext(Context);
      const key = dialogPartKey(part, props.name);
      if (!context) {
        if (key) warnDialogPart(label(key), 'outside');
        return null;
      }
      if (props.hidden) return null;
      const original = key ? context.defaults.get(key) : undefined;
      if (!original) {
        if (key)
          warnDialogPart(
            label(key),
            props.children != null && props.children !== false ? 'childrenOnly' : 'unknown'
          );
        return props.children ?? null;
      }
      // `hidden` and `name` steer this component; they are not element attributes.
      const { children, asChild, className, style, hidden: _hidden, name: _name, ...rest } = props;
      // The part's own wiring outranks anything the host passed, so a stray `onClick` or
      // `data-pptx-part` cannot silently detach the control from the dialog.
      // A default element's `undefined` (a `title` it sets only when disabled) leaves the
      // host's value in place.
      const wiring = Object.fromEntries(
        Object.entries(original.props).filter(([, value]) => value !== undefined)
      );
      const shared: NodeProps = {
        ...rest,
        ...wiring,
        className: [original.props.className, className].filter(Boolean).join(' '),
        style: { ...(original.props.style as CSSProperties | undefined), ...style },
      };
      // A field control (`NumberField`, `SelectField`) has no single element for `asChild`:
      // the host's content replaces the whole field and reads the draft with the dialog's hook.
      const control = (original.type as { dialogControl?: boolean }).dialogControl === true;
      if (control && (asChild || (children !== undefined && children !== null))) {
        if (asChild && key) warnDialogPart(label(key), 'controlAsChild');
        return <>{children}</>;
      }
      const only = asChild ? Children.toArray(children) : [];
      const child = only.length === 1 ? only[0] : undefined;
      if (asChild && !isValidElement<HostElementProps>(child) && key)
        warnDialogPart(label(key), 'asChild');
      if (asChild && isValidElement<HostElementProps>(child)) {
        // The host owns the element's look: forward the wiring and the host's own class and
        // style, never the packaged `pptx-dialog__*` classes.
        const {
          children: _defaultChildren,
          className: _presetClassName,
          style: _presetStyle,
          ...wiring
        } = shared;
        return renderAsChild(child, {
          ...wiring,
          ...(className ? { className } : {}),
          ...(style ? { style } : {}),
        });
      }
      // A misused `asChild` falls back to the packaged part with its packaged content.
      const content = asChild ? undefined : children;
      return cloneElement(
        original,
        shared,
        content === undefined ? replace(original.props.children, context.overrides) : content
      );
    }
    identities.set(Part, part);
    return Part;
  }
  function Composition({
    defaults,
    children,
    preset = true,
    state,
  }: {
    readonly defaults: ReactNode;
    readonly children?: ReactNode;
    readonly preset?: boolean;
    readonly state: State;
  }): ReactNode {
    const defaultMap = new Map<string, ReactElement<NodeProps>>();
    visit(defaults, (node) => {
      const key = markerOf(node.props);
      if (key) defaultMap.set(key, node);
    });
    // Only top-level parts replace a default in place, as in the Vue twin. A part inside a
    // host element renders where the host put it, and the default stays (it says so).
    const overrides = new Map<string, ReactElement<DialogPartProps>>();
    for (const node of flatten(children)) {
      if (!isValidElement<NodeProps>(node)) continue;
      const key = keyOf(node);
      if (key) {
        if (overrides.has(key)) warnDialogPart(label(key), 'duplicate');
        overrides.set(key, node as ReactElement<DialogPartProps>);
      } else if (preset)
        visit(node.props.children, (inner) => {
          const nested = keyOf(inner);
          if (nested) warnDialogPart(label(nested), 'nested');
        });
    }
    // A part with no default in place is never rendered by `replace`, so it says so here.
    if (preset)
      for (const [key, node] of overrides)
        if (!defaultMap.has(key) && !node.props.hidden) warnDialogPart(label(key), 'unknown');
    // A part whose default sits inside another part that brings its own children is dropped
    // with that default.
    if (preset)
      for (const [key, node] of overrides) {
        // Only an owner that really replaces its content drops what it held: a misused
        // `asChild` falls back to the packaged part, defaults and all.
        const only = Children.toArray(node.props.children);
        const replaces = node.props.asChild
          ? only.length === 1 && isValidElement(only[0])
          : node.props.children !== undefined;
        if (!replaces) continue;
        visit(defaultMap.get(key)?.props.children, (inner) => {
          const nested = markerOf(inner.props);
          if (nested && nested !== key && overrides.has(nested))
            warnDialogPart(label(nested), 'shadowed', label(key));
        });
      }
    const extras = flatten(children).filter(
      (node) => !isValidElement<NodeProps>(node) || !keyOf(node)
    );
    return (
      <Context.Provider value={{ defaults: defaultMap, overrides, state }}>
        {preset ? (
          <>
            {replace(defaults, overrides)}
            {extras}
          </>
        ) : (
          children
        )}
      </Context.Provider>
    );
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
    Field: makePart('field') as (props: DialogPartProps & { readonly name: Name }) => ReactNode,
    Composition,
    useState: useDialogState,
    part: (name) => makePart(name) as (props: DialogPartProps) => ReactNode,
  };
}

/**
 * A draft for a dialog's lifetime: created when the component mounts, re-read on every change.
 * The dialog unmounts when it closes (the popups host renders it only while open), so the next
 * opening starts from the document again. @internal
 */
export function useDialogDraft<Fields extends object>(
  options: () => DialogDraftOptions<Fields>
): DialogDraft<Fields> {
  // The draft is made once; its callbacks read the latest render's options, so a commit that
  // closes over host props (`onThemeChange`) never runs a stale copy.
  const latest = useRef<DialogDraftOptions<Fields> | null>(null);
  latest.current = options();
  const [draft] = useState(() => {
    const first = latest.current!;
    return createDialogDraft<Fields>({
      initial: first.initial,
      validate: (values) => latest.current!.validate?.(values) ?? ({} as DialogErrors<Fields>),
      commit: (values) => latest.current!.commit(values),
      resetValues: () => latest.current!.resetValues?.() ?? first.initial,
    });
  });
  useSyncExternalStore(
    draft.subscribe,
    () => draft.values,
    () => draft.values
  );
  useSyncExternalStore(
    draft.subscribe,
    () => draft.errors,
    () => draft.errors
  );
  return draft;
}

/** Props of `DialogFrame`. @internal */
export interface DialogFrameProps {
  /** The dialog's registry key, as `data-pptx-dialog`. */
  readonly kind: string;
  /** The accessible name. */
  readonly label: string;
  readonly className?: string;
  readonly style?: CSSProperties;
  /** Escape, a click on the backdrop, or the host closing it. */
  readonly onClose: () => void;
  /** Keys the frame does not handle itself (Enter to apply). */
  readonly onKeyDown?: (event: KeyboardEvent<HTMLDialogElement>) => void;
  readonly children: ReactNode;
}

/**
 * The native modal every packaged dialog renders in, after DOCX Editor's `DialogFrame`: it
 * opens as a modal on mount and closes on unmount, focuses its first control, keeps Tab inside,
 * closes on Escape and on a click on the backdrop, and gives focus back to the element that
 * had it. Outside an editor scope it carries the scope itself, so the tokens apply. @internal
 */
export function DialogFrame({
  kind,
  label,
  className,
  style,
  onClose,
  onKeyDown,
  children,
}: DialogFrameProps): ReactNode {
  const ref = useRef<HTMLDialogElement>(null);
  const closeRef = useRef(onClose);
  closeRef.current = onClose;
  // Where the last press started: a press inside the panel that is released over the backdrop
  // is a text selection or a drag, not a dismissal.
  const pressedOutside = useRef(false);
  const [ownScope, setOwnScope] = useState(false);
  const scope = useChromeScopeClassName();
  useLayoutEffect(() => {
    const panel = ref.current;
    if (!panel) return;
    setOwnScope(!panel.parentElement?.closest('.pptx-editor'));
    const opener = panel.ownerDocument.activeElement as HTMLElement | null;
    if (typeof panel.showModal === 'function' && !panel.open) panel.showModal();
    else panel.setAttribute('open', ''); // DOM test environments; browsers use native modality.
    initialFocus(panel)?.focus({ preventScroll: true });
    return () => {
      // Unmounted while open: leave the modal state, so the page is not left inert.
      if (typeof panel.close === 'function' && panel.open) panel.close();
      if (opener?.isConnected) opener.focus({ preventScroll: true });
    };
  }, []);
  return (
    <dialog
      ref={ref}
      aria-modal="true"
      aria-label={label}
      className={['pptx-dialog', ownScope ? scope : '', className].filter(Boolean).join(' ')}
      style={style}
      data-pptx-dialog={kind}
      onCancel={(event) => {
        // React dispatches `cancel` to every ancestor's handler: a nested dialog's own Escape
        // must close only that dialog.
        if (fromNestedDialog(event.currentTarget, event.target)) return;
        event.preventDefault();
        closeRef.current();
      }}
      onMouseDown={(event) => {
        pressedOutside.current = isDialogBackdropPoint(event.currentTarget, event);
        event.stopPropagation();
      }}
      onClick={(event) => {
        // As DOCX Editor: a click on the backdrop lands on the dialog element, outside its box.
        // The press must start there too (see `pressedOutside`).
        if (
          event.target === event.currentTarget &&
          pressedOutside.current &&
          isDialogBackdropPoint(event.currentTarget, event)
        )
          closeRef.current();
        pressedOutside.current = false;
        event.stopPropagation();
      }}
      onKeyDown={(event) => {
        // A dialog opened inside this one (the custom gradient editor) handles its own keys;
        // they still stop here, so they never reach the editor's shortcuts.
        if (fromNestedDialog(event.currentTarget, event.target)) {
          event.stopPropagation();
          return;
        }
        if (event.key === 'Escape' && !event.nativeEvent.isComposing) {
          event.preventDefault();
          closeRef.current();
        } else if (trapTabWithin(event.currentTarget, event.nativeEvent)) event.preventDefault();
        else if (!event.nativeEvent.isComposing && !event.defaultPrevented) onKeyDown?.(event);
        event.stopPropagation();
      }}
    >
      {children}
    </dialog>
  );
}

/**
 * Return in a field applies the draft, as a press on OK does: from any input (as
 * DOCX Editor), and from a radio of a radio group. A button, a link, and a text area keep
 * their own Return. @internal
 */
export function applyOnEnter(
  event: KeyboardEvent<HTMLDialogElement>,
  enabled: boolean,
  apply: () => void
): void {
  if (event.key !== 'Enter' || !enabled || !returnAppliesFrom(event.target)) return;
  event.preventDefault();
  apply();
}

/**
 * Close a dialog when its editor opens another document, after DOCX Editor's
 * `useDialogDocument`: a retained draft never edits the replacement. The registry closes its
 * dialogs too; this covers a dialog the host renders itself. @internal
 */
export function useCloseOnNewDocument(onClose: () => void): void {
  const editor = usePptxEditor();
  const closeRef = useRef(onClose);
  closeRef.current = onClose;
  useEffect(() => {
    if (!editor) return;
    let session = editor.session;
    // Once per new document; `destroy()` leaves no session and closes nothing.
    return editor.subscribe(() => {
      if (editor.session === session) return;
      session = editor.session;
      if (session) closeRef.current();
    });
  }, [editor]);
}
