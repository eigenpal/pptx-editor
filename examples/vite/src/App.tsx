// The PPTX Editor demo: the family shell (top bar, menus, mode pill, Open/New/Export, panel
// toggle) around the engine's toolbar, slide rail, canvas, notes, panel, and zoom pill.
// Everything loads and saves in the browser; nothing is uploaded.

import {
  useCallback,
  useEffect,
  useMemo,
  useRef,
  useState,
  type CSSProperties,
  type DragEvent,
  type ReactNode,
} from 'react';
import {
  translateKeyOrText,
  type EditorModule,
  type PptxEditorInstance,
  type SidePanelTab,
} from '@pptx-editor.dev/core/editor';
import { drawingModule } from '@pptx-editor.dev/pro/drawing';
import { PptxEditorReview, reviewModule } from '@pptx-editor.dev/pro/react';
import {
  PptxEditor,
  Icon,
  keepFocus,
  useEditorCommand,
  useEditorState,
  type PptxEditorPopups,
  usePptxEditor,
  useSlideshow,
  useTranslation,
  listenPaneRequests,
} from '@pptx-editor.dev/react';
import { BrandLogo } from './BrandLogo.tsx';
import { ThemeToggle } from './ThemeToggle.tsx';
import { AdapterSwitcher, SHOW_ADAPTER_SWITCHER } from './AdapterSwitcher.tsx';
import { MenuBar } from './menus.tsx';
import { ExportMenu, ModePill } from './TopBarMenus.tsx';
import { DiscardDialog } from './DiscardDialog.tsx';
import { CATALOG_FAMILIES, useDemoFonts } from './fonts.ts';
// The file handling that the React and Vue demos share.
import {
  BLANK,
  EXAMPLES,
  dismissPanelOverlay,
  dismissRailOverlay,
  PPTX_MIME,
  exportDeck as saveDeck,
  exportSlides as saveSlides,
  fetchDeck,
  fetchFirstDeck,
  refuseDeckFetch,
  fileTitle,
  followPanelWidth,
  initialExample,
  initialTheme,
  isPptxFile,
  carriesFiles,
  leftDropZone,
  openDeck,
  panelStartsOpen,
  refuseDroppedFile,
  rememberTheme,
  ribbonChrome,
  showDeckTitle,
  trackUnsavedEdits,
  type SlideExport,
  type UnsavedEdits,
} from '../../shared/demo-files.ts';

// Glow, inner shadow, reflection, and soft edge, and the comments pane (EigenPal Pro License).
const MODULES: readonly EditorModule[] = [
  drawingModule({ licenseKey: import.meta.env.VITE_PPTX_LICENSE }),
  // Initials badges in each author's color, as on the comment cards.
  reviewModule({
    licenseKey: import.meta.env.VITE_PPTX_LICENSE,
    pane: { commentMarkers: 'initials' },
  }),
];
/** The Comments tab of the side panel: the Pro comments pane. */
const REVIEW_PANES = { comments: <PptxEditorReview /> };

/** The author of the comments that the demo adds. */
const DEMO_AUTHOR = 'Demo author';

const base = import.meta.env.BASE_URL;
/** A phone opens with the notes closed, so the slide gets the room; one tap opens them. */
const PHONE_SCREEN =
  typeof matchMedia === 'function' && matchMedia('(max-height: 500px), (max-width: 600px)').matches;
/** The notes pane stays mounted in the slide sorter, so its open state and draft stay. */
const HIDDEN: CSSProperties = { display: 'none' };

function Shell({
  onOpenFile,
  onNew,
  onExample,
  onSaved,
  title,
  setTitle,
  mode,
  setMode,
  theme,
  setTheme,
  deckFailed,
  children,
}: {
  onOpenFile: (f: File) => void;
  onNew: () => void;
  onExample: () => void;
  /** The deck was exported: its edits are saved. */
  onSaved: () => void;
  title: string;
  setTitle: (t: string) => void;
  mode: 'edit' | 'view';
  setMode: (m: 'edit' | 'view') => void;
  theme: 'light' | 'dark';
  setTheme: (t: 'light' | 'dark') => void;
  /** The first deck did not load: no spinner while no deck is open. */
  deckFailed: boolean;
  /** Dialogs of the demo, inside the editor scope so the editor's tokens apply. */
  children?: ReactNode;
}) {
  const { t } = useTranslation();
  const editor = usePptxEditor();
  const sorter = useEditorState((s) => s.presentationView === 'slideSorter');
  // `?chrome=ribbon` shows the ribbon and the status bar in place of the flat toolbar.
  const ribbon = ribbonChrome(location.search);
  // Present follows the slot's state: off with a reason in the slide master view, for one.
  const present = useEditorCommand('view.playFromStart');
  const presentOff = present.disabledReason ? translateKeyOrText(t, present.disabledReason) : null;
  // A phone sizes the slide area to the slide's shape (shell.css).
  const slideRatio = useEditorState((s) => (s.slideWidth ? s.slideHeight / s.slideWidth : 0.5625));
  const show = useSlideshow();
  // With no deck to open (its fetch failed), the notice says why in place of the spinner.
  const waiting = useEditorState((s) => s.status === 'empty');
  // Closed on a narrow window, so the slide keeps a usable zoom (demo-files.ts).
  const [panel, setPanel] = useState(panelStartsOpen);
  const [panelTab, setPanelTab] = useState<SidePanelTab>('format');
  useEffect(() => followPanelWidth(setPanel), []);
  // The toolbar's Selection Pane and Format Pane controls open the panel on their tab.
  useEffect(
    () =>
      editor
        ? listenPaneRequests(editor, (pane) => {
            setPanelTab(pane);
            setPanel(true);
          })
        : undefined,
    [editor]
  );
  const [dragging, setDragging] = useState(false);
  // The slide thumbnails over the slide, on a narrow window only.
  const [rail, setRail] = useState(false);
  const railToggle = useRef<HTMLButtonElement>(null);
  // The overlay rail closes on Escape, on a press outside it, and after a slide is picked.
  useEffect(() => {
    if (!rail) return;
    return dismissRailOverlay({
      root: () => railToggle.current?.closest<HTMLElement>('.demo-root') ?? null,
      editor: () => editor,
      close: () => setRail(false),
    });
  }, [rail, editor]);
  // The overlay side panel (a narrow window) closes on Escape, as the overlay rail does.
  useEffect(() => {
    if (!panel) return;
    return dismissPanelOverlay({
      root: () => railToggle.current?.closest<HTMLElement>('.demo-root') ?? null,
      close: () => setPanel(false),
    });
  }, [panel]);
  // The deck name field, which File > Rename presentation shows and focuses.
  const [renaming, setRenaming] = useState(false);
  const titleInput = useRef<HTMLInputElement>(null);
  const fileInput = useRef<HTMLInputElement>(null);
  const exportDeck = useCallback(() => {
    saveDeck(editor, title || t('app.untitled'));
    onSaved();
  }, [editor, title, t, onSaved]);
  // Slide pictures come from the editor's own paint. Only the PDF goes to the demo's server.
  const exportSlides = useCallback(
    (kind: SlideExport) => saveSlides(editor, kind, title || t('app.untitled')),
    [editor, title, t]
  );
  // A file dropped on the top bar or the workspace opens; the hint hides when the drag leaves
  // either of them. A drag that carries no files (a slide in the rail) is not ours.
  const fileDrop = {
    onDragOver: (e: DragEvent<HTMLElement>) => {
      if (!carriesFiles(e.dataTransfer)) return;
      e.preventDefault();
      setDragging(true);
    },
    onDragLeave: (e: DragEvent<HTMLElement>) => {
      // A move onto a child of the zone is not a leave: the hint does not flicker.
      if (leftDropZone(e.currentTarget, e.relatedTarget)) setDragging(false);
    },
    onDrop: (e: DragEvent<HTMLElement>) => {
      if (!carriesFiles(e.dataTransfer)) return;
      e.preventDefault();
      setDragging(false);
      const f = e.dataTransfer.files[0];
      if (isPptxFile(f)) onOpenFile(f);
      else if (f) refuseDroppedFile(editor, f);
    },
  };
  return (
    // The page has its own landmarks, such as the top bar's banner, so the viewport renders no
    // region of its own.
    <PptxEditor.Viewport className="demo-root" label={false}>
      <header className="demo-topbar" {...fileDrop}>
        <BrandLogo name={t('app.productName')} />
        {SHOW_ADAPTER_SWITCHER ? <AdapterSwitcher current="react" /> : null}
        <span className="demo-topbar__divider" />
        {/* File > Rename presentation opens the name over the menu bar on a narrow window. */}
        <div className={renaming ? 'demo-file demo-file--open' : 'demo-file'}>
          <input
            ref={titleInput}
            className="demo-file__name"
            aria-label={t('app.untitled')}
            value={title}
            onChange={(e) => setTitle(e.target.value)}
            onBlur={() => setRenaming(false)}
            onKeyDown={(e) => {
              if (e.key === 'Enter' || e.key === 'Escape') e.currentTarget.blur();
            }}
          />
          <span className="demo-file__ext">{t('app.extension')}</span>
        </div>
        <MenuBar
          onRename={() => {
            setRenaming(true);
            requestAnimationFrame(
              () => (titleInput.current?.focus(), titleInput.current?.select())
            );
          }}
          onOpen={() => fileInput.current?.click()}
          onSelectionPane={() => {
            setPanelTab('selection');
            setPanel(true);
          }}
          onNew={onNew}
          onExample={onExample}
          onExport={exportDeck}
          onExportSlides={(kind) => void exportSlides(kind)}
        />
        <div className="demo-topbar__actions">
          {/* On a narrow window the slide thumbnails open over the slide (shell.css). */}
          <button
            type="button"
            ref={railToggle}
            className="demo-icon-btn demo-rail-toggle"
            aria-label={t('app.slidesToggle')}
            data-pptx-tooltip={t('app.slidesToggle')}
            aria-pressed={rail}
            onMouseDown={keepFocus}
            onClick={() => setRail((r) => !r)}
          >
            <Icon name="view_agenda" className="" />
          </button>
          <button
            type="button"
            className="demo-btn demo-btn--sm"
            aria-label={t('slideshow.present')}
            data-pptx-tooltip={presentOff ?? t('slideshow.presentHint')}
            disabled={!present.isEnabled}
            onMouseDown={keepFocus}
            onClick={() => show.start('start')}
          >
            <Icon name="slideshow" className="" />
            <span className="demo-btn__label">{t('slideshow.present')}</span>
          </button>
          <ModePill mode={mode} onMode={setMode} />
          <ThemeToggle value={theme} onChange={setTheme} />
          <span className="demo-topbar__divider" />
          <button
            type="button"
            className="demo-btn demo-btn--quiet demo-btn--sm demo-open"
            // The label can hide on a narrow bar: the name and the tooltip stay.
            aria-label={t('menu.open').replace('…', '')}
            data-pptx-tooltip={t('menu.open').replace('…', '')}
            onMouseDown={keepFocus}
            onClick={() => fileInput.current?.click()}
          >
            <Icon name="folder_open" className="" />
            <span className="demo-btn__label">{t('menu.open').replace('…', '')}</span>
          </button>
          <button
            type="button"
            className="demo-btn demo-btn--quiet demo-btn--sm demo-new"
            aria-label={t('menu.new')}
            data-pptx-tooltip={t('menu.new')}
            onMouseDown={keepFocus}
            onClick={onNew}
          >
            <Icon name="note_add" className="" />
            <span className="demo-btn__label">{t('menu.new').split(' ')[0]}</span>
          </button>
          <ExportMenu onExport={exportDeck} onExportSlides={(kind) => void exportSlides(kind)} />
          <button
            type="button"
            className="demo-icon-btn demo-panel-toggle"
            aria-label={t('panel.toggle')}
            data-pptx-tooltip={t('panel.toggle')}
            aria-pressed={panel}
            onMouseDown={keepFocus}
            onClick={() => setPanel((p) => !p)}
          >
            <Icon name="view_sidebar" className="" />
          </button>
        </div>
        <input
          ref={fileInput}
          type="file"
          accept={`.pptx,${PPTX_MIME}`}
          hidden
          onChange={(e) => {
            const f = e.target.files?.[0];
            if (f) onOpenFile(f);
            e.target.value = '';
          }}
        />
      </header>
      <main className="demo-main">
        <h1 className="demo-sr-only">{t('app.productName')}</h1>
        <div className="demo-dock">
          {ribbon ? (
            <PptxEditor.Ribbon />
          ) : (
            <PptxEditor.Toolbar>
              <PptxEditor.Toolbar.ChangeCase hidden />
              <PptxEditor.Toolbar.Link hidden />
            </PptxEditor.Toolbar>
          )}
        </div>
        <div
          className={rail ? 'demo-workspace demo-workspace--rail' : 'demo-workspace'}
          {...fileDrop}
        >
          {sorter ? null : <PptxEditor.SlideRail />}
          {/* The slide sorter takes the canvas's place, so the status bar and the notices stay. */}
          <div className="demo-docpane">
            {sorter ? <PptxEditor.SlideSorter /> : null}
            <div
              className="demo-canvas"
              hidden={sorter}
              style={{ '--demo-slide-ratio': slideRatio } as CSSProperties}
            >
              <PptxEditor.Content />
              <PptxEditor.FindBar />
              {deckFailed && waiting ? null : <PptxEditor.Loading overlay />}
              <PptxEditor.FontNotice />
            </div>
            {/* The view switch and the zoom stay in view in both views, under the slide. */}
            {ribbon ? null : (
              <div className="demo-status">
                <PptxEditor.StatusBar.Views />
                <PptxEditor.ZoomPill />
              </div>
            )}
            <PptxEditor.NotesPane
              defaultOpen={!PHONE_SCREEN}
              {...(sorter ? { style: HIDDEN } : {})}
            />
            <PptxEditor.EditNotice />
            <PptxEditor.OpenFailure />
            {ribbon ? <PptxEditor.StatusBar /> : null}
          </div>
          {panel ? (
            <PptxEditor.SidePanel
              panes={REVIEW_PANES}
              tab={panelTab}
              onTabChange={setPanelTab}
              onClose={() => setPanel(false)}
            />
          ) : null}
          {dragging ? <div className="demo-drop">{t('app.dropHint')}</div> : null}
        </div>
      </main>
      <PptxEditor.Slideshow />
      {children}
    </PptxEditor.Viewport>
  );
}

export function App() {
  const initial = initialExample(location.search);
  const [bytes, setBytes] = useState<Uint8Array | undefined>();
  const [title, setTitle] = useState<string>(EXAMPLES[initial].title);
  const [mode, setMode] = useState<'edit' | 'view'>('edit');
  const [theme, setThemeState] = useState<'light' | 'dark'>(() => initialTheme(location.search));
  // A theme picked here is kept; until then the demo starts light.
  const setTheme = useCallback((next: 'light' | 'dark') => {
    rememberTheme(next);
    setThemeState(next);
  }, []);
  const [editor, setEditor] = useState<PptxEditorInstance | null>(null);
  // Unsaved edits: New, Open, and a drop ask before they replace them, and so does leaving.
  const unsaved = useRef<UnsavedEdits | null>(null);
  useEffect(() => {
    if (!editor) return;
    const tracker = trackUnsavedEdits(editor);
    unsaved.current = tracker;
    return () => tracker.dispose();
  }, [editor]);
  const [discard, setDiscard] = useState<(() => void) | null>(null);
  const replaceDeck = useCallback((run: () => void) => {
    if (unsaved.current?.dirty()) setDiscard(() => run);
    else run();
  }, []);
  // The first deck's download: a deck the user opens before it arrives cancels it.
  const cancelFirst = useRef<() => void>(() => {});
  // The name changes only once the deck opened: a refused file keeps the open deck's name.
  const openBytes = useCallback(
    async (next: Uint8Array, label: string) => {
      cancelFirst.current();
      if (await openDeck(editor, next)) setTitle(label);
    },
    [editor]
  );
  const loadUrl = useCallback(
    async (name: string, label: string) => {
      cancelFirst.current();
      try {
        await openBytes(await fetchDeck(base, name), label);
      } catch (cause) {
        refuseDeckFetch(editor, cause);
      }
    },
    [openBytes, editor]
  );
  const [deckFailed, setDeckFailed] = useState<unknown>(null);
  useEffect(() => {
    // The first deck: `initial` comes from the URL at mount.
    const cancel = fetchFirstDeck(base, EXAMPLES[initial].file, {
      loaded: setBytes,
      failed: setDeckFailed,
    });
    cancelFirst.current = cancel;
    return cancel;
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);
  useEffect(() => {
    if (deckFailed && editor) refuseDeckFetch(editor, deckFailed);
  }, [deckFailed, editor]);
  useEffect(() => {
    document.documentElement.style.colorScheme = theme;
  }, [theme]);
  // The deck name in the browser tab.
  useEffect(() => showDeckTitle(title), [title]);
  const fonts = useDemoFonts();
  // The registry mounts Page setup; this render callback adds the demo's theme switch to it.
  const popups = useMemo<PptxEditorPopups>(
    () => ({
      pageSetup: (props) => (
        <PptxEditor.PageSetupDialog {...props} theme={theme} onThemeChange={setTheme} />
      ),
    }),
    [theme, setTheme]
  );
  return (
    <PptxEditor.Root
      {...(bytes ? { document: bytes } : {})}
      mode={mode}
      colorMode={theme}
      popups={popups}
      modules={MODULES}
      author={DEMO_AUTHOR}
      fonts={fonts}
      fontFamilies={CATALOG_FAMILIES}
      onReady={(instance) => {
        // The e2e suite and the verification scripts drive the editor through this handle.
        (window as unknown as { __pptxEditor?: unknown }).__pptxEditor = instance;
        setEditor(instance);
      }}
    >
      <Shell
        title={title}
        setTitle={setTitle}
        mode={mode}
        setMode={setMode}
        theme={theme}
        setTheme={setTheme}
        onNew={() => replaceDeck(() => void loadUrl(BLANK.file, BLANK.title))}
        onExample={() =>
          replaceDeck(() => void loadUrl(EXAMPLES.showcase.file, EXAMPLES.showcase.title))
        }
        onOpenFile={(f) =>
          replaceDeck(
            () => void f.arrayBuffer().then((b) => openBytes(new Uint8Array(b), fileTitle(f)))
          )
        }
        onSaved={() => unsaved.current?.markSaved()}
        deckFailed={deckFailed !== null}
      >
        {discard ? (
          <DiscardDialog
            onCancel={() => setDiscard(null)}
            onDiscard={() => {
              setDiscard(null);
              discard();
            }}
          />
        ) : null}
      </Shell>
    </PptxEditor.Root>
  );
}
