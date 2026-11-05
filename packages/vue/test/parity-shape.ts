// The Vue adapter's runtime shape for the parity check: each paired component's declared props
// and emits (the runtime truth; the public prop types derive from them), and mount probes that
// prove a host `class` and the `className` prop land on an element, every mapped slot renders,
// and, for the panes and menus, that a paired prop given a value shows in the part. Static reads
// cannot prove these: a root that renders only its slot drops `class` with a Vue warning.
//
// Registers happy-dom, so it runs in its own process (`bun run check:parity-contract`) or in
// the DOM test run (`bun run test:dom`), never in the engine test run.

import { GlobalRegistrator } from '@happy-dom/global-registrator';

export interface VueComponentShape {
  readonly props: readonly string[];
  readonly emits: readonly string[];
  /** A host `class` on the component lands on a rendered element. */
  readonly classLands: boolean;
  /** A `className` prop value lands on a rendered element (false when there is no such prop). */
  readonly classNameLands: boolean;
  /** The slots whose content renders: `default` and each named slot the contract maps. */
  readonly slots: readonly string[];
  /** Value probes: a paired prop, given a value, shows in the rendered part (name → passed). */
  readonly values: Readonly<Record<string, boolean>>;
}

/** One value probe: mount the part with `props` and check that `shows` finds their effect. */
interface ValueProbe {
  readonly props: Record<string, unknown>;
  readonly shows: (scope: ParentNode) => boolean;
}

export type VueAdapterShape = Readonly<Record<string, VueComponentShape>>;

const PROBE_CLASS = 'pptx-parity-probe';
const PROBE_CLASS_NAME = 'pptx-parity-class-name';
const slotId = (slot: string): string => `pptx-parity-slot-${slot}`;

/** The part of `scripts/parity/parity.contract.json` the probes read. */
interface ParityContract {
  readonly components: Record<
    string,
    { readonly reactRenderPropsAsVueSlots?: Record<string, string> } | string
  >;
}

interface ComponentDefinition {
  readonly props?: Record<string, unknown> | readonly string[];
  readonly emits?: Record<string, unknown> | readonly string[];
}

/** An element under `scope` has an inline width of `px` pixels. */
const hasWidth = (scope: ParentNode, px: number): boolean =>
  [...scope.querySelectorAll<HTMLElement>('[style]')].some((el) => el.style.width === `${px}px`);

const keys = (value: ComponentDefinition['props']): string[] =>
  Array.isArray(value) ? [...value] : Object.keys(value ?? {});

/** The slots a probe fills: `default`, and every Vue slot the contract maps for the part. */
function probedSlots(contract: ParityContract | undefined, name: string): string[] {
  const rules = contract?.components[name];
  const mapped =
    typeof rules === 'object' ? Object.values(rules.reactRenderPropsAsVueSlots ?? {}) : [];
  return [...new Set(['default', ...mapped])];
}

/**
 * Mount the paired Vue components and read their shape. With the contract, each probe also
 * fills every named slot the contract maps for the part.
 */
export async function vueAdapterShape(contract?: ParityContract): Promise<VueAdapterShape> {
  if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();
  const { createApp, h, nextTick, provide, shallowRef } = await import('vue');
  const vue = await import('../src/index.ts');
  const { providePptxEditor } = await import('../src/composables.ts');

  // Extra props a probe needs so the component renders its slot.
  const probes: Record<
    string,
    {
      component: unknown;
      props?: Record<string, unknown>;
      tab?: boolean;
      /** Renders the probed vnode inside another part (a menu entry inside the bar). */
      wrap?: (node: ReturnType<typeof h>) => ReturnType<typeof h>;
      /** A selector to click after mount, to open what renders the class and the slot. */
      click?: string;
      /** The part renders in a teleport to the page body, so the probe reads the document. */
      inDocument?: boolean;
      /** Paired props that must show in the rendered part, by the React prop name. */
      values?: Record<string, ValueProbe>;
      /** Milliseconds to wait after mount: a part that shows after the editor refuses a file. */
      settle?: number;
    }
  > = {
    // A file the editor refuses, so the open-failure notice renders its `openFailureActions` slot.
    PptxEditor: {
      component: vue.PptxEditor,
      props: { document: new Uint8Array([1, 2, 3]) },
      settle: 50,
    },
    PptxEditorRoot: { component: vue.PptxEditorRoot },
    PptxEditorToolbar: { component: vue.PptxEditorToolbar, props: { preset: false } },
    ToolbarAction: { component: vue.ToolbarAction, props: { label: 'Probe' } },
    // Outside a preset bar, a group renders its element and its slot.
    ToolbarGroup: { component: vue.ToolbarGroup, props: { id: 'probe', label: 'Probe' } },
    // The ribbon's default slot is its tabs: the probe content goes in a host tab it shows.
    PptxEditorRibbon: {
      component: vue.PptxEditorRibbon,
      props: { preset: false, tab: 'probe' },
      tab: true,
    },
    PptxEditorStatusBar: { component: vue.PptxEditorStatusBar, props: { preset: false } },
    // Outside a root there is no deck, so the screen shows and renders its default slot.
    PptxEditorLoading: { component: vue.PptxEditorLoading },
    // The panes render outside a root too: the class lands on the rail, the notes pane, the
    // panel, the zoom pill, and the sorter.
    PptxEditorSlideRail: {
      component: vue.PptxEditorSlideRail,
      values: {
        paneWidth: { props: { paneWidth: 321 }, shows: (s) => hasWidth(s, 321) },
        orientation: {
          props: { orientation: 'horizontal' },
          shows: (s) => s.querySelector('.pptx-rail[data-orientation="horizontal"]') !== null,
        },
      },
    },
    PptxEditorNotesPane: { component: vue.PptxEditorNotesPane },
    PptxEditorSidePanel: {
      component: vue.PptxEditorSidePanel,
      values: { paneWidth: { props: { paneWidth: 321 }, shows: (s) => hasWidth(s, 321) } },
    },
    PptxEditorZoomPill: { component: vue.PptxEditorZoomPill },
    PptxEditorSlideSorter: { component: vue.PptxEditorSlideSorter },
    GradientDialog: { component: vue.GradientDialog, props: { open: false } },
    // The form dialogs render only while open; the host's content renders beside the default
    // parts, and the host class lands on the dialog surface.
    BackgroundDialog: {
      component: vue.BackgroundDialog,
      props: { open: true, onClose: () => {} },
    },
    BulletsAndNumberingDialog: {
      component: vue.BulletsAndNumberingDialog,
      props: { open: true, onClose: () => {} },
    },
    ColumnsDialog: { component: vue.ColumnsDialog, props: { open: true, onClose: () => {} } },
    FontDialog: { component: vue.FontDialog, props: { open: true, onClose: () => {} } },
    ParagraphDialog: { component: vue.ParagraphDialog, props: { open: true, onClose: () => {} } },
    ShapeStyleGallery: { component: vue.ShapeStyleGallery, props: { onPick: () => {} } },
    MenuTableGrid: { component: vue.MenuTableGrid },
    // Outside a root there are no layouts: the probe gives one, so the gallery renders.
    LayoutGallery: {
      component: vue.LayoutGallery,
      props: {
        layouts: [
          {
            id: 'probe',
            name: 'Probe',
            type: 'blank',
            masterId: 'm',
            masterName: 'Master',
          },
        ],
      },
    },
    // With no anchor the popup panel opens in a teleport to the page body.
    HyperlinkPopup: {
      component: vue.HyperlinkPopup,
      props: { anchor: null, url: '', onClose: () => undefined },
      inDocument: true,
    },
  };

  // The link popup compound renders while a popup state is open: the probe provides one.
  const { hyperlinkPopupKey } = await import('../src/HyperLink.tsx');
  const { hyperlinkPopupFromRequest } = await import('@pptx-editor.dev/core/editor');
  const openLinkPopup = (node: ReturnType<typeof h>) =>
    h({
      setup() {
        const controller = hyperlinkPopupFromRequest(null, {
          anchor: null,
          url: 'https://example.com/',
          onClose: () => undefined,
        });
        provide(hyperlinkPopupKey, {
          controller,
          state: shallowRef(controller.getState()),
          errorId: 'probe-error',
        });
        return () => node;
      },
    });
  probes.PptxEditorHyperLink = {
    component: vue.PptxEditorHyperLink,
    wrap: openLinkPopup,
    inDocument: true,
  };

  // The menus: a menu entry renders in the bar's open panel; the context menus render in the
  // page body, and the canvas menu only inside an open request.
  const menuState = { at: { x: 0, y: 0 }, keyboard: false, owner: null };
  const { ContextMenuTargetProvider } = await import('../src/context-menu-parts.tsx');
  Object.assign(probes, {
    PptxEditorMenu: { component: vue.PptxEditorMenu, props: { preset: false } },
    MenuEntry: {
      component: vue.MenuEntry,
      props: { id: 'probe', label: 'Probe' },
      wrap: (node: ReturnType<typeof h>) => h(vue.PptxEditorMenu, { preset: false }, () => [node]),
      click: '.pptx-menubar__trigger',
      inDocument: true,
    },
    MenuGroup: {
      component: vue.MenuGroup,
      values: {
        label: {
          props: { label: 'Probe group' },
          shows: (s: ParentNode) => (s as Element).textContent?.includes('Probe group') ?? false,
        },
      },
    },
    MenuSlotItem: { component: vue.MenuSlotItem, props: { slotId: 'text.bold' } },
    PptxEditorContextMenu: {
      component: vue.PptxEditorContextMenu,
      wrap: (node: ReturnType<typeof h>) =>
        h(
          ContextMenuTargetProvider,
          {
            value: { ...menuState, kind: 'slide', owner: document.body, close: () => undefined },
          },
          () => [node]
        ),
      inDocument: true,
    },
    ContextMenu: {
      component: vue.ContextMenu,
      props: { ...menuState, label: 'Probe', onClose: () => undefined },
      inDocument: true,
      values: {
        label: {
          props: { label: 'Probe menu' },
          shows: (s: ParentNode) => s.querySelector('[aria-label="Probe menu"]') !== null,
        },
      },
    },
    RailContextMenu: {
      component: vue.RailContextMenu,
      props: { menu: { ...menuState, index: 0 }, onClose: () => undefined },
      inDocument: true,
    },
    SectionContextMenu: {
      component: vue.SectionContextMenu,
      props: {
        section: { id: 'probe', name: 'Probe', firstSlide: 0, slideCount: 0 },
        menu: menuState,
        onClose: () => undefined,
        onRename: () => undefined,
        onRemoved: () => undefined,
      },
      inDocument: true,
    },
  });

  // The parts that render inside a bar, a ribbon, or a root, and the registry controls.
  const { chromeControlForSlot, EMPTY_EDITOR_SNAPSHOT } =
    await import('@pptx-editor.dev/core/editor');
  const control = (slot: string) => {
    const found = chromeControlForSlot(slot as Parameters<typeof chromeControlForSlot>[0]);
    if (!found) throw new Error(`parity probe: no registry control for ${slot}`);
    return { slot, control: found };
  };
  type Wrap = (node: ReturnType<typeof h>) => ReturnType<typeof h>;
  const inRoot: Wrap = (node) => h(vue.PptxEditorRoot, null, () => [node]);
  const inToolbar: Wrap = (node) =>
    inRoot(h(vue.PptxEditorToolbar, { preset: false }, () => [node]));
  // A bar under an editor that takes every command, so a picker opens without a deck.
  const inReadyToolbar: Wrap = (node) =>
    h({
      setup() {
        providePptxEditor(
          shallowRef({
            snapshot: () => ({ ...EMPTY_EDITOR_SNAPSHOT, status: 'ready' }),
            subscribe: () => () => undefined,
            can: () => ({ ok: true }),
            session: null,
            fontFamilies: () => [],
          } as unknown as Parameters<typeof providePptxEditor>[0]['value'])
        );
        return () => h(vue.PptxEditorToolbar, { preset: false }, () => [node]);
      },
    });
  const inStatusBar: Wrap = (node) =>
    inRoot(h(vue.PptxEditorStatusBar, { preset: false }, () => [node]));
  const dialog = { open: true, onClose: () => undefined };
  Object.assign(probes, {
    ImageAltTextDialog: { component: vue.ImageAltTextDialog, props: dialog },
    ImagePropertiesDialog: { component: vue.ImagePropertiesDialog, props: dialog },
    PageSetupDialog: { component: vue.PageSetupDialog, props: dialog, wrap: inRoot },
    RibbonTab: {
      component: vue.RibbonTab,
      props: { id: 'probe', label: 'Probe' },
      wrap: (node: ReturnType<typeof h>) =>
        h(vue.PptxEditorRibbon, { preset: false, tab: 'probe' }, () => [node]),
    },
    // Its content renders as the host element of the button (`asChild`), as in React.
    StatusBarNotes: { component: vue.StatusBarNotes, props: { asChild: true }, wrap: inStatusBar },
    StatusBarSlideNumber: { component: vue.StatusBarSlideNumber, wrap: inStatusBar },
    StatusBarViews: { component: vue.StatusBarViews, wrap: inStatusBar },
    StatusBarZoom: { component: vue.StatusBarZoom, wrap: inStatusBar },
    ToolbarSeparator: { component: vue.ToolbarSeparator, wrap: inToolbar },
    ToolbarSlot: { component: vue.ToolbarSlot, props: { slotId: 'text.bold' }, wrap: inToolbar },
    // The font family parts render inside `PptxEditorToolbar.FontFamily`; the list and its items
    // open from the trigger into the page body. The list opens only while the editor takes a
    // font, so those probes run under an editor that takes every command.
    FontFamilyTrigger: {
      component: vue.PptxEditorToolbar.FontFamily.Trigger,
      wrap: (node: ReturnType<typeof h>) =>
        inToolbar(h(vue.PptxEditorToolbar.FontFamily, null, () => [node])),
    },
    FontFamilyContent: {
      component: vue.PptxEditorToolbar.FontFamily.Content,
      wrap: (node: ReturnType<typeof h>) =>
        inReadyToolbar(
          h(vue.PptxEditorToolbar.FontFamily, null, () => [
            h(vue.PptxEditorToolbar.FontFamily.Trigger),
            node,
          ])
        ),
      click: '[data-slot="font.family"]',
      inDocument: true,
    },
    FontFamilyItem: {
      component: vue.PptxEditorToolbar.FontFamily.Item,
      props: { value: 'Arial' },
      wrap: (node: ReturnType<typeof h>) =>
        inReadyToolbar(
          h(vue.PptxEditorToolbar.FontFamily, null, () => [
            h(vue.PptxEditorToolbar.FontFamily.Trigger),
            h(vue.PptxEditorToolbar.FontFamily.Content, null, () => [node]),
          ])
        ),
      click: '[data-slot="font.family"]',
      inDocument: true,
    },
    ToolbarControl: {
      component: vue.ToolbarControl,
      props: { group: 'text', control: control('text.bold').control },
      wrap: inToolbar,
    },
    PptxEditorLoadingSpinner: { component: vue.PptxEditorLoadingSpinner },
    PptxEditorViewport: { component: vue.PptxEditorViewport },
    PptxEditorContent: { component: vue.PptxEditorContent, wrap: inRoot },
    PptxEditorFindBar: { component: vue.PptxEditorFindBar, wrap: inRoot },
    PptxEditorEditNotice: { component: vue.PptxEditorEditNotice, wrap: inRoot },
    PptxEditorOpenFailure: {
      component: vue.PptxEditorOpenFailure,
      // An editor that refused a file: the notice renders, with its class and its slots.
      wrap: (node: ReturnType<typeof h>) =>
        h({
          setup() {
            providePptxEditor(
              shallowRef({
                snapshot: () => ({
                  ...EMPTY_EDITOR_SNAPSHOT,
                  status: 'ready',
                  openFailure: { code: 'zip', reason: 'damaged' },
                }),
                subscribe: () => () => undefined,
              } as unknown as Parameters<typeof providePptxEditor>[0]['value'])
            );
            return () => node;
          },
        }),
    },
    PptxEditorFontNotice: { component: vue.PptxEditorFontNotice, wrap: inRoot },
    MenuRow: { component: vue.MenuRow, props: { label: 'Probe' } },
    // The submenu's panel, with its class and content, opens from its row into the page body.
    MenuSubmenu: {
      component: vue.MenuSubmenu,
      props: { label: 'Probe' },
      click: '.pptx-menu__row',
      inDocument: true,
    },
    MenuItem: { component: vue.MenuItem, props: { onSelect: () => undefined } },
    Popover: {
      component: vue.Popover,
      props: { anchor: null, onClose: () => undefined },
      inDocument: true,
    },
    ToolButton: {
      component: vue.ToolButton,
      props: { label: 'Probe', onClick: () => undefined },
      wrap: inRoot,
    },
    Icon: { component: vue.Icon, props: { name: 'bold' } },
    ShapeIcon: { component: vue.ShapeIcon, props: { preset: 'rect' } },
    ShapeGrid: {
      component: vue.ShapeGrid,
      props: { presets: ['rect'], onPick: () => {} },
      wrap: inRoot,
    },
    LineGlyph: { component: vue.LineGlyph, props: { value: 'solid' } },
    FillControl: { component: vue.FillControl, props: control('shape.fill'), wrap: inToolbar },
    LineMenuControl: {
      component: vue.LineMenuControl,
      props: control('shape.lineStart'),
      wrap: inToolbar,
    },
    ShapeGalleryControl: {
      component: vue.ShapeGalleryControl,
      props: control('insert.shape'),
      wrap: inToolbar,
    },
    NewSlideControl: {
      component: vue.NewSlideControl,
      props: control('slide.add'),
      wrap: inToolbar,
    },
    LayoutControl: {
      component: vue.LayoutControl,
      props: control('slide.layout'),
      wrap: inToolbar,
    },
  });

  type Probe = (typeof probes)[string];
  /** Mount one probe with the given slots filled, and read what renders. */
  const mount = async (
    probe: Probe,
    filled: readonly string[],
    value?: ValueProbe
  ): Promise<{
    classLands: boolean;
    classNameLands: boolean;
    slots: string[];
    shows: boolean;
  }> => {
    const props = keys((probe.component as ComponentDefinition).props);
    const host = document.createElement('div');
    document.body.append(host);
    const warn = console.warn;
    // Vue warns about the dropped `class` this probe is looking for; the probe reports it.
    console.warn = () => undefined;
    const slotContent = (slot: string) => () =>
      slot === 'default' && probe.tab
        ? h(vue.RibbonTab, { id: 'probe', label: 'Probe' }, () => [h('i', { id: slotId(slot) })])
        : h('i', { id: slotId(slot) });
    const node = () =>
      h(
        probe.component as Parameters<typeof h>[0],
        {
          class: PROBE_CLASS,
          ...(props.includes('className') ? { className: PROBE_CLASS_NAME } : {}),
          ...probe.props,
          ...value?.props,
        },
        Object.fromEntries(filled.map((slot) => [slot, slotContent(slot)]))
      );
    const app = createApp({ render: () => (probe.wrap ? probe.wrap(node()) : node()) });
    try {
      app.mount(host);
      await nextTick();
      if (probe.settle) {
        await new Promise((resolve) => setTimeout(resolve, probe.settle));
        await nextTick();
      }
      if (probe.click) {
        host.querySelector<HTMLElement>(probe.click)?.click();
        await nextTick();
      }
      const scope: ParentNode = probe.inDocument ? document : host;
      return {
        classLands: scope.querySelector(`.${PROBE_CLASS}`) !== null,
        classNameLands: scope.querySelector(`.${PROBE_CLASS_NAME}`) !== null,
        slots: filled.filter((slot) => scope.querySelector(`#${slotId(slot)}`) !== null),
        shows: value ? value.shows(probe.inDocument ? document.body : host) : false,
      };
    } finally {
      app.unmount();
      host.remove();
      console.warn = warn;
    }
  };

  const out: Record<string, VueComponentShape> = {};
  for (const [name, probe] of Object.entries(probes)) {
    const definition = probe.component as ComponentDefinition;
    const slotNames = probedSlots(contract, name);
    // Every slot filled first. A part can render its class only without host content (a slot
    // that replaces the default control), and a named slot only without the default slot (a
    // loading indicator inside the default screen), so those get a mount of their own.
    const all = await mount(probe, slotNames);
    let { classLands, classNameLands } = all;
    if (!classLands || !classNameLands) {
      const bare = await mount(probe, []);
      classLands ||= bare.classLands;
      classNameLands ||= bare.classNameLands;
    }
    const slots = new Set(all.slots);
    for (const slot of slotNames)
      if (!slots.has(slot) && (await mount(probe, [slot])).slots.includes(slot)) slots.add(slot);
    const values: Record<string, boolean> = {};
    for (const [prop, value] of Object.entries(probe.values ?? {}))
      values[prop] = (await mount(probe, [], value)).shows;
    out[name] = {
      props: keys(definition.props),
      emits: keys(definition.emits),
      classLands,
      classNameLands,
      slots: slotNames.filter((slot) => slots.has(slot)),
      values,
    };
  }
  return out;
}
