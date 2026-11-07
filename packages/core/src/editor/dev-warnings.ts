// Development warnings for customizations that cannot do what they say. Shared by the React
// and Vue adapters (moved from packages/react/src/dev-warnings.ts). After DOCX Editor's
// `packages/react/src/editor/toolbar/toolbar-warnings.ts` and `lib/is-development.ts`.
//
// A misplaced part still renders, because breaking the host's page is worse than a misplaced
// control. But it must not fail silently: each warning names the part, the value, and what the
// toolbar did instead. Each message prints once, and never in a production build. Every
// warning takes the adapter's own part names, so a Vue host reads `ToolbarSlot` and a React
// host reads `Toolbar.Slot`.

/** The names an adapter gives its toolbar parts, for warning text. @internal */
export interface ToolbarPartNames {
  /** The slot part: `Toolbar.Slot` in React, `ToolbarSlot` in Vue. */
  readonly slot: string;
  /** The group part: `Toolbar.Group` in React, `ToolbarGroup` in Vue. */
  readonly group: string;
}

const REACT_PARTS: ToolbarPartNames = { slot: 'Toolbar.Slot', group: 'Toolbar.Group' };

/** Vite's flag when it is a boolean, else the Node convention; production when neither says. */
function isDevelopment(): boolean {
  let dev: unknown;
  try {
    dev = (import.meta as { readonly env?: { readonly DEV?: unknown } }).env?.DEV;
  } catch {
    dev = undefined;
  }
  if (typeof dev === 'boolean') return dev;
  const hasProcess = typeof process !== 'undefined';
  return hasProcess && process.env.NODE_ENV !== 'production';
}

const printed = new Set<string>();

/** Print `message` once per page, in development builds only. */
function devWarning(message: string): void {
  if (!isDevelopment() || printed.has(message)) return;
  printed.add(message);
  console.warn(`[pptx-editor] ${message}`);
}

/** Forget printed messages. Tests only. @internal */
export function resetDevWarnings(): void {
  printed.clear();
}

/** The names an adapter gives its menu bar and root, for warning text. @internal */
export interface MenuPartNames {
  /** The menu bar: `PptxEditor.Menu` in React, `PptxEditorMenu` in Vue. */
  readonly bar: string;
  /** The root: `PptxEditor.Root` in React, `PptxEditorRoot` in Vue. */
  readonly root: string;
}

const REACT_MENU_PARTS: MenuPartNames = { bar: 'PptxEditor.Menu', root: 'PptxEditor.Root' };

/**
 * A menu part is misused, after DOCX Editor's menu warnings: a row for an id that is not a
 * registered control, a menu outside the menu bar, or a context menu outside the root. Each renders nothing. `part` and `names` are the adapter's own
 * part names. @internal
 */
export function warnMenuPart(
  part: string,
  problem:
    | 'unknownSlot'
    | 'outsideBar'
    | 'outsideContextMenu'
    | 'unmatchedHidden'
    | 'unknownAnchor'
    | 'valueRow',
  slot?: string,
  names: MenuPartNames = REACT_MENU_PARTS
): void {
  devWarning(
    problem === 'unknownSlot'
      ? `${part} slotId="${slot}" is not a registered control, so the row renders nothing.`
      : problem === 'unmatchedHidden'
        ? `${part} slotId="${slot}" hidden matches no row of this menu, so it removes nothing.`
        : problem === 'unknownAnchor'
          ? `${part} has after="${slot}", but no menu has that id. The menu goes after every other menu.`
          : problem === 'valueRow'
            ? `${part} slotId="${slot}" picks a value (a color, a font, a size) and has no menu ` +
              `form yet, so the row shows as disabled. Give it a value, or use the toolbar control.`
            : problem === 'outsideBar'
              ? `${part} is outside ${names.bar}, so it renders nothing. Render it as a direct ` +
                `child of ${names.bar}.`
              : `${part} is outside ${names.root}, so it renders nothing. Place it in ` +
                `${names.root} or its viewport, or render it from the popups.contextMenu entry.`
  );
}

/** A second pane of one kind in the same root: it keeps its own state. @internal */
export function warnDuplicatePane(part: string): void {
  devWarning(
    `${part} is rendered twice in one PptxEditor.Root. The first one shares its open state ` +
      `with the root's hook; each other one keeps its own (or follows its own open prop).`
  );
}

/** A hook or part that needs `PptxEditor.Root` is used outside one: it does nothing. @internal */
export function warnOutsideRoot(part: string): void {
  devWarning(
    `${part} is used outside PptxEditor.Root, so it does nothing. Render it inside the root.`
  );
}

/**
 * A compound part (`Toolbar.FontFamily.Item`) that is not inside its compound root, so it has no
 * state to read and renders nothing. @internal
 */
export function warnPartOutsideCompound(part: string, owner: string): void {
  devWarning(`${part} is used outside ${owner}, so it renders nothing. Render it inside ${owner}.`);
}

/**
 * A font picker `Item` whose `value` is neither a valid family name nor a theme font reference
 * (`+mj-lt`, `+mn-lt`), so it renders nothing. @internal
 */
export function warnInvalidFontItem(part: string, value: string): void {
  devWarning(
    `${part} value "${value}" is not a font family name or a theme font reference (+mj-lt, ` +
      `+mn-lt), so it renders nothing. A name has 1 to 64 letters, digits, spaces, or - . + _.`
  );
}

/** An icon name that is not in `ICON_PATHS`: the icon draws empty. @internal */
export function warnUnknownIcon(name: string): void {
  devWarning(
    `The icon "${name}" is not a packaged Material Symbol, so it draws nothing. Use a name from ` +
      `ICON_PATHS, or pass icon content (an element).`
  );
}

/** `useEditorCommand('text.link').execute()` with no URL, where no link popup can open. @internal */
export function warnNoLinkPopup(): void {
  devWarning(
    'useEditorCommand("text.link").execute() opens the link popup, but none can open here ' +
      '(popups.hyperlink is false, or there is no PptxEditor.Root). Pass the URL instead: ' +
      'execute(url) sets the link, and execute(null) removes it.'
  );
}

/** `useEditorCommand('image.altText').execute()` with no description, where no dialog can open. @internal */
export function warnNoAltTextPopup(): void {
  devWarning(
    'useEditorCommand("image.altText").execute() opens the Alt Text dialog, but none can open ' +
      'here (popups.imageAltText is false, or there is no PptxEditor.Root). Pass the ' +
      'description instead: execute(text) sets the alt text of the selected element.'
  );
}

/** A slot part names an id that is not a registered control. @internal */
export function warnUnknownSlot(slot: string, parts: ToolbarPartNames = REACT_PARTS): void {
  devWarning(
    `"${slot}" is not a toolbar slot id, so ${parts.slot} renders nothing. ` +
      `Use a ChromeSlotId such as "text.bold".`
  );
}

/** Two slot parts name the same slot. @internal */
export function warnDuplicateSlot(slot: string, parts: ToolbarPartNames = REACT_PARTS): void {
  devWarning(`${parts.slot} "${slot}" appears more than once. The last one is used.`);
}

/**
 * A slot part names a registered control the preset bar does not draw; the toolbar puts it
 * after the groups. DOCX Editor's `warnSlotOutsideArrangement`. @internal
 */
export function warnSlotOutsideBar(slot: string, parts: ToolbarPartNames = REACT_PARTS): void {
  devWarning(
    `${parts.slot} "${slot}" is not part of the preset toolbar, so it is appended after the ` +
      `groups. Put it in a ${parts.group} to give it a place.`
  );
}

/** `asChild` without exactly one element child. @internal */
export function warnAsChildChild(slot: string, parts: ToolbarPartNames = REACT_PARTS): void {
  devWarning(
    `${parts.slot} "${slot}" has asChild but not exactly one element child, so the ` +
      `packaged control renders instead.`
  );
}

/** An override that the slot's control cannot apply: it is a picker, not one button. @internal */
export function warnUnsupportedOverride(
  slot: string,
  setting: 'icon' | 'asChild',
  parts: ToolbarPartNames = REACT_PARTS
): void {
  devWarning(
    `${parts.slot} "${slot}" renders a picker, not a single button, so its ${setting} is ` +
      `ignored. To replace the control, pass children to ${parts.slot} slotId="${slot}".`
  );
}

/**
 * A slot part has its own content, which replaces the control, and also a setting for the
 * control (a class, an icon): the control does not render, so the setting does nothing. `setting`
 * is the adapter's prop name (`className` in React, `class` in Vue). @internal
 */
export function warnReplacedControlSetting(
  slot: string,
  setting: string,
  parts: ToolbarPartNames = REACT_PARTS
): void {
  devWarning(
    `${parts.slot} "${slot}" has its own content, which replaces the control, so its ` +
      `${setting} is ignored. Put the ${setting} on your content instead.`
  );
}

/**
 * A slot part has attributes that it does not pass to its control (`style`, `id`, `data-*`,
 * `aria-*`, a listener): the control keeps its own state attributes and handlers, as in DOCX
 * Editor. `classProp` is the adapter's class prop (`className` in React, `class` in Vue).
 * @internal
 */
export function warnIgnoredSlotAttributes(
  slot: string,
  names: readonly string[],
  classProp: string,
  parts: ToolbarPartNames = REACT_PARTS
): void {
  if (names.length === 0) return;
  const list = names.join(', ');
  devWarning(
    `${parts.slot} "${slot}" does not pass ${list} to its control, so ` +
      `${names.length === 1 ? 'it is' : 'they are'} ignored. Style the control with ${classProp} ` +
      `and its data-active and data-disabled attributes, or replace it with your own content.`
  );
}

/** A host group's `after` names no group. @internal */
export function warnUnknownGroupAnchor(
  id: string,
  after: string,
  known: readonly string[],
  parts: ToolbarPartNames = REACT_PARTS
): void {
  devWarning(
    `${parts.group} "${id}" has after="${after}", but no group has that id. ` +
      `The group is placed after every other group. Known ids: ${known.join(', ')}.`
  );
}

/** A built-in group override sets a prop that only a host group uses. @internal */
export function warnBuiltInGroupSetting(
  id: string,
  setting: string,
  parts: ToolbarPartNames = REACT_PARTS
): void {
  devWarning(
    `${parts.group} "${id}" is a built-in group, so its ${setting} is ignored. ` +
      `A built-in group keeps its own label and place in the bar.`
  );
}

/**
 * A dialog part that the dialog cannot use. `part` is the adapter's part name, for example
 * `PageSetupDialog.Field name="theme"`. `unknown`: the dialog renders no default for it, so the
 * part renders nothing; `duplicate`: the last of several parts with the same name is used;
 * `asChild`: not exactly one element child, so the packaged part renders; `nested`: inside a
 * host element, so it renders there and the default stays; `outside`: not inside its dialog;
 * `childrenOnly`: no default, so only the part's children render; `shadowed`: its default sits
 * inside `owner`, which brings its own children, so the part renders nothing; `controlAsChild`:
 * a field control has no one element for `asChild`, so its children replace the field. @internal
 */
export function warnDialogPart(
  part: string,
  problem:
    | 'unknown'
    | 'duplicate'
    | 'asChild'
    | 'nested'
    | 'outside'
    | 'childrenOnly'
    | 'shadowed'
    | 'controlAsChild',
  /** For 'shadowed': the part whose own children replace the default that holds `part`. */
  owner?: string
): void {
  devWarning(
    problem === 'shadowed'
      ? `${part} replaces a default inside ${owner}, but ${owner} has its own children, so ` +
          `${part} renders nothing. Put ${part} inside ${owner}'s children instead.`
      : problem === 'outside'
        ? `${part} is outside its dialog, so it renders nothing. Render it as a child of its dialog.`
        : problem === 'childrenOnly'
          ? `${part} matches no part the dialog renders now, so it renders only its children. ` +
            `Check the part's name, and that the dialog shows that part (some show only with a prop).`
          : problem === 'nested'
            ? `${part} is inside another element, so it does not replace the default part: both ` +
              `render. Put the part at the top level of the dialog's children to replace it in place.`
            : problem === 'unknown'
              ? `${part} matches no part the dialog renders now, so it renders nothing. ` +
                `Check the part's name, and that the dialog shows that part (some show only with a prop).`
              : problem === 'duplicate'
                ? `${part} appears more than once. The last one is used.`
                : problem === 'controlAsChild'
                  ? `${part} is one field, which has no single element for asChild, so its ` +
                    `children replace the whole field. Drop asChild, and read and write the ` +
                    `draft from your control with the dialog's hook (for example useParagraphDialog()).`
                  : `${part} has asChild but not exactly one element child, so the packaged part renders instead.`
  );
}

/** What writes each dialog's fields, in both adapters: a format hook, or the command. */
const DIALOG_HOOKS: Readonly<Record<string, string>> = {
  paragraph: 'useParagraphFormat().apply',
  font: 'useFontFormat().apply',
  bulletsAndNumbering: 'useBulletFormat().apply',
  columns: "editor.exec({ type: 'setColumns', count, gapEmu })",
};

/**
 * An adapter was asked to open a registry dialog that it does not package, and the host gave no
 * renderer for it, so nothing opens. @internal
 */
export function warnDialogNeedsRenderer(adapter: string, dialog: string): void {
  const write = DIALOG_HOOKS[dialog];
  devWarning(
    `The ${adapter} adapter has no packaged \`${dialog}\` dialog, so it opens only with a ` +
      `renderer: pass popups.${dialog}` +
      (write ? `, or write its fields with ${write}.` : '.')
  );
}

/**
 * A dialog slot (`paragraph.dialog`) was run as a command, or a menu row's value that opens a
 * dialog ("Line spacing options…") was run through its slot. Neither runs a command. @internal
 */
export function warnDialogSlotRun(slot: string, dialog: string, rowValue?: string): void {
  const write = DIALOG_HOOKS[dialog];
  const how =
    `In React and Vue, open the dialog with useEditorDialog('${dialog}').open()` +
    (write ? `, or write its fields with ${write}.` : '.');
  devWarning(
    rowValue === undefined
      ? `\`${slot}\` opens the \`${dialog}\` dialog and runs no command by itself. ${how}`
      : `\`${slot}\`'s row "${rowValue}" opens the \`${dialog}\` dialog; it is not a value of the slot. ${how}`
  );
}

/** A value slot got a value it cannot run (`shape.effects` with `'shadow.bogus'`). @internal */
export function warnInvalidSlotValue(slot: string, value: unknown): void {
  if (!isDevelopment()) return;
  devWarning(
    `The value ${describeValue(value)} is not a choice of "${slot}"; the control did nothing. ` +
      'ToolbarValueMap lists the value types of typed slots.'
  );
}

/** A value for a warning. Never throws: a host may pass an event, which is cyclic. */
function describeValue(value: unknown): string {
  if (value === null || typeof value !== 'object') {
    return typeof value === 'string' ? JSON.stringify(value) : String(value);
  }
  try {
    return JSON.stringify(value);
  } catch {
    return Object.prototype.toString.call(value);
  }
}

/** `renderLayoutThumbnail` names no layout of the deck. @internal */
export function warnUnknownLayout(layoutId: string): void {
  devWarning(
    `renderLayoutThumbnail("${layoutId}") names no slide layout of the open deck, so the ` +
      `container is left empty. Use an id from snapshot().layouts.`
  );
}

/** A host ran a slot that adapter chrome runs, not one engine command. @internal */
export function warnChromeSlotRun(slot: string, reason: string): void {
  devWarning(`runToolbarCommand("${slot}") ran nothing: ${reason}.`);
}

/** `requestSectionRename` found no slide rail or slide sorter of the editor. @internal */
export function warnSectionRenameUnheard(): void {
  devWarning(
    'requestSectionRename found no slide rail or slide sorter of this editor, so no name field ' +
      'opens. Mount PptxEditor.SlideRail or PptxEditor.SlideSorter, or run the renameSection ' +
      'command with a name.'
  );
}

/** The view became Slide Sorter while no slide sorter of the editor is mounted. @internal */
export function warnSlideSorterUnmounted(): void {
  devWarning(
    'The view is Slide Sorter, but no PptxEditor.SlideSorter of this editor is mounted, so no ' +
      'slide shows. Render PptxEditor.SlideSorter when snapshot().presentationView is ' +
      '"slideSorter", or keep the Normal view.'
  );
}

/**
 * A ribbon tab override that cannot do what it says. `tab` is the adapter's part name
 * (`Ribbon.Tab` in React, `PptxEditorRibbon.Tab` in Vue). @internal
 */
export function warnRibbonTab(
  tab: string,
  id: string,
  problem: 'untitled' | 'hiddenUnknown' | 'builtInAfter' | 'unknownAfter',
  known: readonly string[],
  after?: string
): void {
  const ids = `Built-in tab ids: ${known.join(', ')}.`;
  const text = {
    untitled: `${tab} id="${id}" is not a built-in tab and has no label or labelKey, so the ribbon shows its id as the title. Check the id, or give the tab a label. ${ids}`,
    hiddenUnknown: `${tab} id="${id}" is hidden, but no tab has that id, so nothing is removed. ${ids}`,
    builtInAfter: `${tab} id="${id}" is a built-in tab, so its after is ignored. A built-in tab keeps its own place.`,
    unknownAfter: `${tab} id="${id}" has after="${after ?? ''}", but no tab has that id. The tab is placed after every other tab.`,
  }[problem];
  devWarning(text);
}

/** A ribbon child that is not a tab is not drawn. `ribbon` is the adapter's part name. @internal */
export function warnRibbonChild(ribbon: string, tab: string): void {
  devWarning(
    `${ribbon} draws only ${tab} children; another child is not drawn. Put toolbar parts inside a ${tab}.`
  );
}

/**
 * A collaboration session has an empty `sessionId`. New ids need a per-replica actor, so the
 * editor mints dense numbers instead, and two peers can mint one id twice.
 * @internal
 */
export function warnEmptySessionId(): void {
  devWarning(
    'The collaboration session has an empty sessionId. New ids are not kept apart from the ids of other replicas. Give each replica a unique sessionId.'
  );
}

/** The names an adapter gives its open-failure actions, for warning text. @internal */
export interface OpenFailureActionNames {
  /** The actions input: `renderOpenFailureActions` in React, `#openFailureActions` in Vue. */
  readonly actions: string;
  /** The value that removes the notice: `openFailure={false}` or `:open-failure="false"`. */
  readonly off: string;
}

/**
 * The editor's open-failure actions cannot do what they say, after DOCX Editor's
 * `warnOpenFailureActions`: they replace `openFailure.actions`, and with the notice turned off
 * they render nothing. Each case prints once. @internal
 */
export function warnOpenFailureActions(
  openFailure: boolean | { readonly actions?: unknown },
  hasActions: boolean,
  names: OpenFailureActionNames
): void {
  if (!hasActions) return;
  if (openFailure === false) {
    devWarning(`${names.actions} is ignored with ${names.off}, which renders no notice.`);
  } else if (typeof openFailure === 'object' && openFailure.actions !== undefined) {
    devWarning(`${names.actions} replaces openFailure.actions. Pass only one.`);
  }
}

/**
 * Whether two `modules` lists hold the same modules in the same order, by the identity of each
 * module, so an inline array of the same module objects counts as no change. @internal
 */
export function sameModules(
  a: readonly unknown[] | undefined,
  b: readonly unknown[] | undefined
): boolean {
  const left = a ?? [];
  const right = b ?? [];
  return left.length === right.length && left.every((module, i) => module === right[i]);
}

/**
 * The root's `modules` changed after it created the editor. Modules register when the editor is
 * created, as in DOCX Editor, where `modules` is sampled at mount only, so the change has no
 * effect. `root` is the adapter's root name (`PptxEditor.Root` in React, `PptxEditorRoot` in
 * Vue). Prints once. @internal
 */
export function warnModulesChanged(
  created: readonly unknown[] | undefined,
  next: readonly unknown[] | undefined,
  root: string
): void {
  if (sameModules(created, next)) return;
  devWarning(
    `${root} ignores a modules change after mount: modules register when the editor is created. ` +
      'Create the modules once (outside the component), or change the root key to remount it.'
  );
}

/** How many changes in {@link CHURN_WINDOW_MS} count as a prop that changes on every render. */
const CHURN_LIMIT = 10;
const CHURN_WINDOW_MS = 1000;

/**
 * A counter for one root prop whose identity reloads or replaces something (`document` loads
 * the deck, `fonts` replaces the fonts). A prop made anew on each render, with a host that
 * renders on `onChange`, loads forever. Call the result on each change; past ten changes in
 * one second it warns once, in development. `root` is the adapter's root name. @internal
 */
export function churnWarning(root: string, prop: 'document' | 'fonts'): () => void {
  let times: number[] = [];
  let warned = false;
  return () => {
    if (warned) return;
    const now = Date.now();
    times = times.filter((t) => now - t < CHURN_WINDOW_MS);
    times.push(now);
    if (times.length <= CHURN_LIMIT) return;
    warned = true;
    devWarning(
      prop === 'document'
        ? `${root} got new \`document\` bytes ${times.length} times in one second, and each one ` +
            'loads the deck again. Keep the bytes in state, and pass the same Uint8Array until ' +
            'the deck changes.'
        : `${root} got a new \`fonts\` value ${times.length} times in one second, and each one ` +
            'replaces the fonts. Make it once with `useFonts(...)` or outside the component.'
    );
  };
}

/** A slide rail `orientation` that is not `vertical` or `horizontal`: the rail stays vertical. @internal */
export function warnInvalidRailOrientation(part: string, value: unknown): void {
  devWarning(
    `${part} orientation "${String(value)}" is not "vertical" or "horizontal", so the rail ` +
      `stays vertical.`
  );
}
