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
import type { PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import {
  PptxEditor,
  Icon,
  keepFocus,
  useEditorState,
  type PptxEditorPopups,
  usePptxEditor,
  useSlideshow,
  useTranslation,
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
  exportPicture as savePicture,
  fetchDeck,
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
  type PictureExport,
  type UnsavedEdits,
} from '../../shared/demo-files.ts';

const base = import.meta.env.BASE_URL;
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
  /** Dialogs of the demo, inside the editor scope so the editor's tokens apply. */
  children?: ReactNode;
}) {
  const { t } = useTranslation();
  const editor = usePptxEditor();
  const sorter = useEditorState((s) => s.presentationView === 'slideSorter');
  // `?chrome=ribbon` shows the ribbon and the status bar in place of the flat toolbar.
  const ribbon = ribbonChrome(location.search);
  const canPresent = useEditorState((s) => s.status === 'ready' && s.slides.length > 0);
  // A phone sizes the slide area to the slide's shape (shell.css).
  const slideRatio = useEditorState((s) => (s.slideWidth ? s.slideHeight / s.slideWidth : 0.5625));
  const show = useSlideshow();
  // Closed on a narrow window, so the slide keeps a usable zoom (demo-files.ts).
  const [panel, setPanel] = useState(panelStartsOpen);
  useEffect(() => followPanelWidth(setPanel), []);
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
  // Slide pictures come from the editor's own paint; nothing leaves the browser.
  const exportPicture = useCallback(
    (kind: PictureExport) => savePicture(editor, kind, title || t('app.untitled')),
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
    <PptxEditor.Viewport className="demo-root">
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
          onNew={onNew}
          onExample={onExample}
          onExport={exportDeck}
          onExportPicture={(kind) => void exportPicture(kind)}
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
            data-pptx-tooltip={t('slideshow.presentHint')}
            disabled={!canPresent}
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
            className="demo-btn demo-btn--quiet demo-btn--sm"
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
          <ExportMenu onExport={exportDeck} onExportPicture={(kind) => void exportPicture(kind)} />
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
            <PptxEditor.Loading overlay />
            <PptxEditor.FontNotice />
          </div>
          {/* The view switch and the zoom stay in view in both views, under the slide. */}
          {ribbon ? null : (
            <div className="demo-status">
              <PptxEditor.StatusBar.Views />
              <PptxEditor.ZoomPill />
            </div>
          )}
          <PptxEditor.NotesPane {...(sorter ? { style: HIDDEN } : {})} />
          <PptxEditor.EditNotice />
          <PptxEditor.OpenFailure />
          {ribbon ? <PptxEditor.StatusBar /> : null}
        </div>
        {panel ? <PptxEditor.SidePanel onClose={() => setPanel(false)} /> : null}
        {dragging ? <div className="demo-drop">{t('app.dropHint')}</div> : null}
      </div>
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
  // A theme picked here is kept; until then the demo follows the system color scheme.
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
  // The name changes only once the deck opened: a refused file keeps the open deck's name.
  const openBytes = useCallback(
    async (next: Uint8Array, label: string) => {
      if (await openDeck(editor, next)) setTitle(label);
    },
    [editor]
  );
  const loadUrl = useCallback(
    async (name: string, label: string) => openBytes(await fetchDeck(base, name), label),
    [openBytes]
  );
  useEffect(() => {
    // The first deck: `initial` comes from the URL at mount.
    void fetchDeck(base, EXAMPLES[initial].file).then(setBytes);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);
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
