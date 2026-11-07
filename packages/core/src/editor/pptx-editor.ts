// `createPptxEditor`: the full Editor contract over the document session and the surface.
//
// Snapshot contract (same as docx-editor's facade): `snapshot()` returns the SAME object until
// state moves, so `useSyncExternalStore` selectors only re-render when their slice changes.

import { pickImageInto } from './chrome-actions.ts';
import { runLinkOf } from './link-read.ts';
import { documentOpenFailure, OPEN_FAILED, openRejectionCode } from './loading.ts';
import { prepareDeck } from './open-steps.ts';
import { writePptx, type PptxOp, type PptxPackage } from '../store/index.ts';
import { registerFontFaces, type FontFaceHandle } from '../output/index.ts';
import type { SlideLayout } from '../layout/index.ts';
import type {
  CanResult,
  CommandResult,
  EditorCommand,
  EditorExecOptions,
  EditorMode,
  EditorRejection,
  EditorSelection,
  EditorSnapshot,
  TextPoint,
} from '../contracts/index.ts';
import { tableStyleList } from '../cascade/index.ts';
import { createClipboardHost } from './clipboard-host.ts';
import { createOpeningLane } from './pptx-editor-opening.ts';
import { commandCacheKey, compile, type CommandContext } from './commands.ts';
import { clampSlide, type DocumentSession } from './document-session.ts';
import { changedSlide, isHistoryMark, validSelection, type HistoryMark } from './history-marks.ts';
import { markOptions } from './history-marks.ts';
import { EditorHistoryGroups } from './history-groups.ts';
import {
  CommandRefusal,
  commitRefusalDetail,
  isNothingToDo,
} from '../automation/commands/op-batch.ts';
import {
  createFontFamilySource,
  createNotesFormattedSource,
  createNotesSource,
  createSubstitutionSource,
  createVersionSource,
} from './snapshot-sources.ts';
import { createShapeListSource, elementIn } from './shape-list.ts';
import { activeOf, createBackgroundSource, formatAt, geometryOf } from './selection-state.ts';
import { FormatPainter } from './format-painter.ts';
import { CropMode } from './crop-mode.ts';
import { pictureStateAt, withAltText } from './picture-state.ts';
import { isViewCommand, ViewSettings, type ViewCommand } from './view-settings.ts';
import { PlaceholderPrompts } from './placeholder-prompts.ts';
import { RemovalNotice } from './removal-notice.ts';
import { CanCache, checkCommand, createStructureKey, typingBox } from './can-cache.ts';
import { itemById, SlideSurface } from './surface.ts';
import { createShortcut, createSurfaceHost, type SurfaceHostBase } from './surface-host.ts';
import { setSurfaceRoot } from './surface-registry.ts';
import { refit } from '../automation/commands/autofit.ts';
import { withAutoCorrect } from './autocorrect.ts';
import { selectAllOf } from './shape-keys.ts';
import { clampZoom } from './zoom.ts';
import { snapTextSelection, type TextInputKind } from './text-input.ts';
import { comparePoints } from '../automation/commands/text-model.ts';
import { CaretArm, armedCaretOf, caretSelection } from './caret-arm.ts';
import { fontFamilyOf } from '../automation/commands/font-format.ts';
import { createHostFonts, openOptionsOf } from './host-fonts.ts';
import { selectionBody, snapCellAddress, textItemOf } from './table-cells.ts';
import { repairTablesForSave } from '../automation/commands/table-repair.ts';
import { withoutStaleCopies } from '../automation/commands/alternate-copies.ts';
import { firstCellCaret } from './table-keys.ts';
import { createTableStateSource } from './table-state.ts';
import { createThemeColorSource } from './theme-palette.ts';
import { createSlideSummaries, SlideSelection } from './slide-selection.ts';
import { createLayoutSummaries } from '../automation/commands/slide-layouts.ts';
import { createSectionSummaries } from '../automation/commands/slide-sections.ts';
import { createCustomShowSummaries } from '../automation/commands/custom-show-links.ts';
import { hasSlideNumber } from '../automation/commands/slide-numbers.ts';
import { dateShown, footerText } from '../automation/commands/header-footer.ts';
import { createAnimationsSource } from './animation-command-table.ts';
import { frameLabelOf, ThumbnailPainter } from './thumbnails.ts';
import { uniquePrefix } from './paint-ids.ts';
import { createFind } from './find-session.ts';
import * as seam from './modules.ts';
import type { PptxEditorInstance, PptxEditorOptions } from './pptx-editor-types.ts';
import { PresenceLink } from './collaboration-presence.ts';
import { INSERT_COMMAND_TYPES as INSERTS } from './command-types.ts';

export type { PptxEditorInstance, PptxEditorOptions } from './pptx-editor-types.ts';

/** Create an editor. Load a document with `options.document` or `load(bytes)`. @public */
export function createPptxEditor(options: PptxEditorOptions = {}): PptxEditorInstance {
  let session: DocumentSession | null = null;
  let fonts: FontFaceHandle | null = null;
  let status: EditorSnapshot['status'] = 'empty';
  let failure: ReturnType<typeof documentOpenFailure> | undefined;
  /** Why the open deck is read-only, after edits, because a slide of it did not parse. */
  let readOnly: EditorSnapshot['readOnlyReason'] = null;
  /** The mode the host asked for (`setMode`, `options.mode`). */
  let hostMode: EditorMode = options.mode ?? 'edit';
  /** The mode in effect: view while the deck is read-only (`readOnlyReason`). */
  let mode: EditorMode = hostMode;
  // The caret and the store revision just after AutoCorrect started a list.
  let autoListed: { readonly sel: EditorSelection; readonly rev: number } | null = null;
  const typeText = withAutoCorrect(options.autoCorrect !== false);
  let slide = 0;
  let selection: EditorSelection = { kind: 'none' };
  let zoom = typeof options.zoom === 'number' ? clampZoom(options.zoom) : 1;
  let fitZoom = options.zoom === undefined || options.zoom === 'fit';
  let fitValue = 1;
  const views = new ViewSettings(options.smartGuides ?? true);
  let surface: SlideSurface | null = null;
  let cached: EditorSnapshot | null = null;
  // The typing burst, the drag, and the host handles that make one undo step each.
  const groups = new EditorHistoryGroups();
  // Repeated pastes of one payload cascade; a parsed payload is cached by its clipboard text.
  let pasteSeries: { payload: unknown; count: number } = { payload: null, count: 0 };
  let lastRejection: EditorRejection | null = null;
  const removal = new RemovalNotice(); // `snapshot.lastRemoval`
  const painter = new FormatPainter();
  const cropper = new CropMode();
  /** The picture in crop mode, while it is the selection. */
  const cropping = (): string | null => cropper.nodeFor(partAt(slide), selection);
  const prompts = new PlaceholderPrompts(options.translate);
  // `load` calls in flight: only the newest may install its session, and none after `destroy`.
  let generation = 0;
  let destroyed = false;
  const listeners = new Set<() => void>();
  const modules = seam.resolveEditorModules(options.modules);
  const presence = new PresenceLink(modules);
  let detachModules = (): void => {};
  const idPrefix = uniquePrefix();
  const thumbnails = new ThumbnailPainter({
    session: () => session,
    idPrefix,
    frameLabel: (kind) => frameLabelOf(options.translate, kind),
    layoutPreview: (owner, id) => prompts.preview(owner, id),
    currentSlide: () => slide,
  });

  const effectiveZoom = (): number => (fitZoom ? fitValue : zoom);
  const layout = (): SlideLayout | undefined => (session ? session.layout(slide) : undefined);
  // `can` compiles the command; toolbars ask for every slot on every change, so answers are
  // cached until the state a command compiles against moves: the package, the slide, the
  // selection, or the mode. Zoom, hover, and fit changes keep them. New Slide and Add Section
  // read no slide content and no selection: typing keeps their answers (`STRUCTURE_COMMANDS`).
  const structureOf = createStructureKey();
  const canCache = new CanCache<CommandContext['active']>(
    () => [
      session?.store.package,
      slide,
      slideSel.version,
      selection,
      mode,
      painter.state,
      cropper.state,
      arm.version,
      views.presentationView,
    ],
    // The slide is part of each answer's key instead (`structureSlideKey`).
    () => [...structureOf(session), mode, views.presentationView]
  );
  const emit = (): void => {
    arm.sync();
    // The slide sorter shows no slide canvas: no shape or text selection there, whatever set it
    // (an undo, Find, Select All, or `setSelection`), so no command edits a slide out of view.
    if (views.presentationView === 'slideSorter') selection = { kind: 'none' };
    cached = null;
    shownPart = partAt(slide);
    cropping(); // Crop mode ends as soon as the selection is no longer its picture.
    presence.publish(session?.store.package, shownPart, selection);
    surface?.render();
    for (const l of [...listeners]) l();
  };
  // Commits the editor makes itself emit once, after it has set the selection and the slide.
  let committing = 0;
  // The slide part on screen at the last emit: a change made on the store directly keeps it.
  let shownPart: string | undefined;

  const snapshot = (): EditorSnapshot => {
    if (cached) return cached;
    const info = session?.info;
    const lay = layout();
    const revision = session?.store.revision ?? 0;
    const pkg = session?.store.package;
    const slides = slideSummaries(session, pkg);
    cached = {
      status,
      ...(failure ? { error: failure.code, openFailure: failure } : {}),
      revision,
      mode,
      slides,
      pendingSlides: opening.pending(),
      isOpening: status === 'loading',
      readOnlyReason: readOnly,
      currentSlide: slide,
      selectedSlides: slideSel.indices(info?.slides ?? [], slide),
      slideNumbers: pkg ? hasSlideNumber(pkg, slide) : false,
      footer: pkg ? footerText(pkg, slide, (i) => !session!.slidePending(i)) : null,
      dateTime: pkg ? dateShown(pkg, slide, (i) => !session!.slidePending(i)) : null,
      slideBackground: stableBackground(lay),
      slideWidth: info?.width ?? 12_192_000,
      slideHeight: info?.height ?? 6_858_000,
      zoom: effectiveZoom(),
      fitZoom,
      insertions,
      ...views.snapshot(sectionSummaries(pkg).map((x) => x.id)),
      selection,
      format: formatAt(arm.layout() ?? lay, selection, (rId, action) =>
        runLinkOf(pkg, partAt(slide), info?.slides.map((s) => s.partName) ?? [], rId, action)
      ),
      geometry: withAltText(
        geometryOf(lay, selection, (id) => elementIn(pkg, partAt(slide), id)),
        pkg,
        partAt(slide)
      ),
      canUndo:
        arm.current() !== null ||
        (seam.canSessionHistory(modules, true) ?? session?.store.canUndo ?? false),
      canRedo: seam.canSessionHistory(modules, false) ?? session?.store.canRedo ?? false,
      notes: notesOf(pkg),
      notesFormatted: notesFormatted(session, pkg, slide),
      embeddedFontFaces: session?.fontReport.embeddedFaces ?? 0,
      // Fonts still on the way are not substitutions yet, as in DOCX Editor.
      substitutedFonts: hostFonts.resolving() ? NO_SUBSTITUTIONS : stableSubstitutions(pkg),
      themeColors: themeColors(session ? pkg : undefined, slide),
      lastRejection,
      lastRemoval: removal.current,
      formatPainter: painter.mode,
      image: pictureStateAt(pkg, partAt(slide), selection, cropping(), lay),
      layouts: layoutSummaries(pkg),
      sections: sectionSummaries(pkg),
      customShows: customShowSummaries(pkg),
      table: tableState(pkg, partAt(slide), lay, selection),
      shapes: shapeList(pkg, partAt(slide)),
      animations: animationsOf(pkg, partAt(slide)),
    };
    return cached;
  };
  const [slideSummaries, shapeList] = [createSlideSummaries(), createShapeListSource()];
  const animationsOf = createAnimationsSource();
  const [layoutSummaries, sectionSummaries] = [createLayoutSummaries(), createSectionSummaries()];
  const customShowSummaries = createCustomShowSummaries();
  const slideSel = new SlideSelection();
  const notesSource = createNotesSource();
  const notesFormatted = createNotesFormattedSource();
  const notesOf = (pkg: PptxPackage | undefined): string => notesSource(session, pkg, slide);
  const themeColors = createThemeColorSource();
  const stableBackground = createBackgroundSource();
  const tableState = createTableStateSource();
  const substitutions = createSubstitutionSource();
  const stableSubstitutions = (pkg: PptxPackage | undefined) => substitutions(session, pkg);
  const NO_SUBSTITUTIONS: EditorSnapshot['substitutedFonts'] = Object.freeze([]);
  const partVersion = createVersionSource();

  // A format armed at a word edge for the text typed next (caret-arm.ts).
  const arm = new CaretArm({
    package: () => session?.store.package,
    caret: () => armedCaretOf(selection, partAt(slide)),
    compileOn: (command, base, lay, at) => {
      const active = activeOf(formatAt(lay, caretSelection(at)));
      const sel = caretSelection(at);
      return compile({ ...context(command), selection: sel, base, active, caretArm: true }, command)
        .batch;
    },
    formatAt: (lay, at) => formatAt(lay, caretSelection(at)),
    layoutOf: (pkg) => (pkg === session?.store.package ? layout() : session?.layoutOf(pkg, slide)),
  });

  /** Formatting state for toggles, read once per state. */
  const activeFormat = (): CommandContext['active'] =>
    canCache.active(() => activeOf(formatAt(arm.layout() ?? layout(), selection)));

  const context = (command?: EditorCommand, probe = false): CommandContext => ({
    session: session!,
    slide,
    selectedSlides: slideSel.indices(session!.info.slides, slide),
    ...(options.translate ? { translate: options.translate } : {}),
    selection,
    readOnly: mode === 'view',
    ...(views.presentationView === 'slideSorter' ? { slideSorter: true } : {}),
    ...(probe ? { probe } : {}),
    ...(modules.collaboration ? { collaborating: true } : {}),
    ...(command?.type === 'paste'
      ? {
          pasteRepeat: (command.payload === pasteSeries.payload ? pasteSeries.count : 0) + 1,
        }
      : {}),
    active: activeFormat(),
  });

  const can = (command: EditorCommand, options?: EditorExecOptions): CanResult => {
    if (!session) return { ok: false, reason: 'disabled.noEditor' };
    if (groups.token(options?.historyGroup) === null)
      return { ok: false, reason: 'errors.unsupported' };
    const key =
      command.type === 'insertImage' || command.type === 'replaceImage' || command.type === 'paste'
        ? // Commands carrying bytes are cached by object identity, not serialized.
          `${command.type}:${partVersion(command)}`
        : commandCacheKey(command);
    const at = { pkg: session.store.package, slide, mode, selection };
    return canCache.ask(command, key, at, () => canUncached(command));
  };

  const sectionIds = () => sectionSummaries(session?.store.package).map((x) => x.id);
  const viewRefusal = (command: ViewCommand) => views.refused(command, sectionIds());
  const canUncached = (command: EditorCommand): CanResult => {
    if (isViewCommand(command)) return viewRefusal(command) ?? { ok: true };
    if (!session) return { ok: false, reason: 'disabled.noEditor' };
    if (command.type === 'undo')
      return (arm.current() || (seam.canSessionHistory(modules, true) ?? session.store.canUndo)) &&
        mode === 'edit'
        ? { ok: true }
        : { ok: false, reason: mode === 'view' ? 'disabled.viewing' : 'disabled.nothingToUndo' };
    if (command.type === 'redo')
      return (seam.canSessionHistory(modules, false) ?? session.store.canRedo) && mode === 'edit'
        ? { ok: true }
        : { ok: false, reason: mode === 'view' ? 'disabled.viewing' : 'disabled.nothingToRedo' };
    return checkCommand(command, selection, mode === 'view', session, slide, (only) => {
      try {
        // A probe: the same compiler without autofit (it adds ops, never a refusal); ops drop.
        const ctx = { ...context(command, true), ...(only ? { selection: only } : {}) };
        if ((painter.compile(ctx, command) ?? cropper.compile(ctx, command)) === null)
          compile(ctx, command);
        return { ok: true };
      } catch (e) {
        return { ok: false, reason: e instanceof CommandRefusal ? e.reason : 'errors.unsupported' };
      }
    });
  };

  /** Publish a refusal (`lastRejection`) and return it; "nothing to do" carries `noop`. */
  const refuse = (reason: string, detail?: string): CommandResult => {
    const rejection = {
      reason,
      ...(detail ? { detail } : {}),
      ...(isNothingToDo(reason) ? { noop: true } : {}),
    };
    lastRejection = rejection;
    emit();
    return { ok: false, ...rejection };
  };

  const partAt = (index: number): string | undefined => session?.info.slides[index]?.partName;

  /** The history mark for an edit of `part`: the selection, when it is on that slide. */
  const markFor = (part: string | undefined): HistoryMark => ({
    part,
    selection: part === partAt(slide) ? selection : { kind: 'none' },
  });

  /**
   * Commit ops as one history entry (or a frame of the open one, for the same `historyGroup`),
   * remembering the slide and selection before it, and after it when the caller knows it.
   */
  const commit = (
    ops: readonly PptxOp[],
    options: { label: string; historyGroup?: symbol },
    before: HistoryMark,
    after?: HistoryMark
  ) => {
    committing += 1;
    try {
      const r = session!.store.transact(withoutStaleCopies(session!.store.package, ops), {
        ...options,
        ...markOptions(before, after),
      });
      if (r.ok) ((lastRejection = null), removal.clear());
      return r;
    } finally {
      committing -= 1;
    }
  };
  /** Record the slide and selection after the newest entry: redo restores them. */
  const markAfter = (part: string | undefined): void => {
    session?.store.setSelectionAfter(markFor(part));
  };

  /** Go to the slide of a history mark and restore its selection, if it is still valid. */
  const restore = (mark: HistoryMark | undefined): void => {
    if (!session) return;
    const index = mark?.part ? session.info.slides.findIndex((s) => s.partName === mark.part) : -1;
    slide = index >= 0 ? index : Math.min(slide, Math.max(0, session.info.slides.length - 1));
    selection = mark && index >= 0 ? validSelection(layout(), mark.selection) : { kind: 'none' };
  };

  // Objects inserted so far (`snapshot.insertions`): a ribbon opens their Format tab.
  let insertions = 0;
  const exec = (command: EditorCommand, options?: EditorExecOptions): CommandResult => {
    const hostGroup = groups.token(options?.historyGroup);
    if (hostGroup === null) return refuse('errors.unsupported', 'historyGroup');
    if (isViewCommand(command)) {
      const no = viewRefusal(command);
      if (no || !views.apply(command, sectionIds())) return refuse(no?.reason ?? 'errors.view');
      // The slide sorter edits slides only: no shape or text stays selected behind it.
      if (views.presentationView === 'slideSorter') selection = { kind: 'none' };
      return (emit(), { ok: true });
    }
    if (!session) return { ok: false, reason: 'disabled.noEditor' };
    if (command.type === 'undo' || command.type === 'redo') {
      if (mode === 'view') return refuse('disabled.viewing');
      surface?.cancelDrag();
      const undo = command.type === 'undo';
      // An armed format is the last step, as in DOCX Editor: the first undo ends it alone.
      if (undo && arm.current()) return (arm.clear(), emit(), { ok: true });
      arm.clear();
      // With collaboration history, undo reverts only this replica's changes, in the room.
      const pkgBefore = session.store.package;
      const shared = seam.runSessionHistory(modules, undo);
      if (shared === 'conflict') return refuse('disabled.historyConflict');
      if (shared === 'empty')
        return refuse(undo ? 'disabled.nothingToUndo' : 'disabled.nothingToRedo');
      if (shared === 'ok') {
        const slides = session.info.slides.map((sl) => sl.partName);
        const to = presence.historyTarget(pkgBefore, session.store.package, slides, slide);
        slide = to.slide;
        selection = validSelection(layout(), to.selection ?? selection);
        slideSel.reset(partAt(slide));
        session.store.closeHistoryGroup(); // a boundary in the room too
        removal.clear();
        return ((lastRejection = null), groups.boundary(), emit(), { ok: true });
      }
      committing += 1;
      let r: { readonly ok: boolean; readonly selection?: unknown };
      try {
        r = undo ? session.store.undo() : session.store.redo();
      } finally {
        committing -= 1;
      }
      if (!r.ok) return refuse(undo ? 'disabled.nothingToUndo' : 'disabled.nothingToRedo');
      // Undo goes to the changed slide and its `selectionBefore`; redo to `selectionAfter`.
      const mark = isHistoryMark(r.selection) ? r.selection : undefined;
      restore(mark);
      if (undo) arm.rearm(mark); // Typing undone: the format it took is armed again.
      slideSel.reset(partAt(slide));
      lastRejection = null;
      removal.clear();
      groups.boundary();
      emit();
      return { ok: true };
    }
    let compiled;
    try {
      const ctx = context(command);
      const own = painter.compile(ctx, command) ?? cropper.compile(ctx, command);
      if (own === 'state') return (emit(), { ok: true });
      compiled = own ?? compile(ctx, command);
    } catch (e) {
      if (e instanceof CommandRefusal) return refuse(e.reason, e.detail);
      throw e;
    }
    const { batch, effect } = compiled;
    const removed = removal.pending(command, selection);
    // A family loads only for a format that is kept: armed, or committed.
    const family = fontFamilyOf(command);
    // At a word edge the format waits for the text typed next: nothing is committed.
    if (batch.armedCaret) {
      if (!arm.add(command)) return refuse('errors.unsupported', 'caret');
      if (family) void ensureFamilies([family]);
      return ((lastRejection = null), groups.boundary(), emit(), { ok: true });
    }
    // Consecutive nudges of one selection are one undo step, unless the host names a group.
    const nudge = command.type === 'nudge';
    groups.boundary(!nudge);
    const historyGroup = hostGroup ?? (nudge ? groups.nudge() : undefined);
    const target = partAt(changedSlide(command, slide));
    const revision = session.store.revision;
    const r = commit(
      batch.ops,
      { label: effect.label, ...(historyGroup ? { historyGroup } : {}) },
      markFor(target)
    );
    if (!r.ok) return refuse('errors.unsupported', commitRefusalDetail(r.rejection));
    if (family) void ensureFamilies([family]);
    if (effect.slide !== undefined)
      slide = Math.max(0, Math.min(effect.slide, session.info.slides.length - 1));
    if (effect.slides) slideSel.set(session.info.slides, effect.slides);
    else if (effect.slide !== undefined) slideSel.reset(partAt(slide));
    if (effect.selection) selection = effect.selection;
    if (INSERTS.has(command.type)) insertions += 1;
    cropper.after(command, partAt(slide), selection);
    // A command that changed nothing keeps the newest entry's selection after.
    if (session.store.revision !== revision) {
      markAfter(effect.slide !== undefined ? partAt(slide) : target);
      removal.record(removed);
    }
    if (command.type === 'paste')
      pasteSeries = {
        payload: command.payload,
        count: command.payload === pasteSeries.payload ? pasteSeries.count + 1 : 1,
      };
    emit();
    // Deleting a table removes its contextual chrome, and the focus with it: it goes to the
    // slide, not to the page.
    if (effect.selection?.kind === 'text' || effect.label === 'deleteTable') surface?.focus();
    return { ok: true };
  };

  /**
   * Apply a text edit at the current text selection as one transaction; typing at a caret
   * merges into one undo step. A refusal is returned and published, never dropped.
   */
  const textInput = (kind: TextInputKind, data?: string): CommandResult => {
    if (!session) return { ok: false, reason: 'disabled.noEditor' };
    if (mode === 'view') return refuse('disabled.viewing');
    if (selection.kind !== 'text')
      return refuse(selection.kind === 'cells' ? 'disabled.typeInOneCell' : 'disabled.noText');
    const sel = selection;
    const part = partAt(slide);
    const item = textItemOf(layout(), sel);
    if (!part || !item) return refuse('disabled.noText');
    // Typed text takes an armed format in its transaction; a break carries it (caret-arm.ts).
    const typed = kind === 'text' && /[^\n\v]/u.test(data ?? '');
    const deletes = kind === 'deleteBackward' || kind === 'deleteForward';
    const carried =
      deletes || (!typed && (kind === 'text' || kind === 'paragraph' || kind === 'lineBreak'));
    const armed = typed ? arm.current() : null;
    const batch = session.batch(armed?.pkg);
    let at: TextPoint;
    let listed = false;
    try {
      const txBody = selectionBody(batch, part, sel);
      if (!txBody) throw new CommandRefusal('disabled.noText');
      ({ at, list: listed } = typeText(batch, part, slide, txBody.id, item, sel, kind, data));
      const caret = { part, nodeId: sel.nodeId, at, ...(sel.cell ? { cell: sel.cell } : {}) };
      if (carried) for (const op of arm.carry(batch.pkg, caret)) batch.push(op);
      refit(batch, slide, part, sel.nodeId, session.measurer);
    } catch (e) {
      if (e instanceof CommandRefusal) return refuse(e.reason, e.detail);
      throw e;
    }
    // Typed text is one burst, one undo step; text typed over a range starts a new burst that
    // holds the replace, so one undo removes both. A break or a delete is a step of its own.
    const burst = kind === 'text' && !listed;
    if (!burst || comparePoints(sel.anchor, sel.focus) !== 0) groups.boundary();
    const caret: EditorSelection = { ...sel, anchor: at, focus: at };
    // The entry's `selectionBefore`: an undo returns this object, so the arm finds its commands.
    const mark = { part, selection: sel };
    const r = commit(
      armed ? [...armed.ops, ...batch.ops] : batch.ops,
      { label: 'typing', ...(burst ? { historyGroup: groups.typing(sel.nodeId) } : {}) },
      mark,
      { part, selection: caret }
    );
    if (!r.ok) return refuse('errors.unsupported', commitRefusalDetail(r.rejection));
    if (armed) arm.typed(mark, armed.commands);
    if (!carried) arm.clear();
    selection = caret;
    autoListed = listed ? { sel: caret, rev: session.store.revision } : null;
    canCache.typed(`${mode}:${slide}${typingBox(caret)}`, emit);
    emit();
    return { ok: true };
  };

  const shortcut = createShortcut({
    editor: () => instance,
    layout: () => layout(),
    selection: () => selection,
    setSelection: (next) => ((selection = next), emit()),
    mode: () => mode,
    exec: (command) => exec(command),
    select: (next) => select(next),
    selectAll: () => selectAll(),
    textInput: (kind, data) => textInput(kind, data),
    enterText: (id) => enterText(id, null),
    surface: () => surface,
    fitZoom: () => instance.setZoom('fit'),
    cancelTools: () => (painter.cancel() || cropper.cancel() ? (emit(), true) : false),
    boundary: () => groups.boundary(),
    // Backspace right after AutoCorrect started a list gives the typed marker back (undo).
    autoListed: () => autoListed?.sel === selection && autoListed.rev === session?.store.revision,
  });

  /** Enter text editing at the end of a shape's text. */
  const enterText = (id: string, point: TextPoint | null): boolean => {
    const cell = !point && mode !== 'view' ? firstCellCaret(layout(), id) : undefined;
    if (cell) return (select(cell), surface?.focus(), true);
    const item = itemById(layout(), id);
    if (!item || item.kind !== 'shape' || !item.element.text || mode === 'view') return false;
    const lens = item.text?.paragraphLengths ?? [0];
    const end = point ?? {
      paragraph: Math.max(0, lens.length - 1),
      offset: lens[lens.length - 1] ?? 0,
    };
    selection = { kind: 'text', nodeId: id, anchor: end, focus: end };
    emit();
    surface?.focus();
    return true;
  };

  const select = (asked: EditorSelection): void => {
    // A caret never stops inside an emoji or a field (or on a covered table cell): snap it.
    const requested = snapCellAddress(layout(), asked);
    const item = textItemOf(layout(), requested);
    const next =
      requested.kind === 'text' && item?.element.text
        ? snapTextSelection(item, requested)
        : requested;
    if (JSON.stringify(next) === JSON.stringify(selection)) return;
    selection = next;
    lastRejection = null;
    groups.boundary(); // Ends typing and nudges; a host handle's step stays open (no write).
    emit();
  };

  /** The session's faces changed: register them for paint (in a browser) and publish. */
  const refreshFaces = async (owner: DocumentSession): Promise<void> => {
    if (typeof document !== 'undefined') {
      fonts?.dispose();
      const faces = registerFontFaces(document, owner.faces);
      fonts = faces;
      await faces.ready;
      // A load or destroy in the meantime owns the faces now.
      if (destroyed || session !== owner) return;
    }
    cached = null;
    emit();
  };
  const hostFonts = createHostFonts(options, {
    session: () => session,
    installed: refreshFaces,
    changed: () => {
      cached = null;
      emit();
    },
  });
  const ensureFamilies = (families: readonly string[]) => hostFonts.request(families);

  const familySource = createFontFamilySource(options.fontFamilies ?? [], () =>
    hostFonts.supportedFamilies()
  );
  const fontFamilies = (): readonly string[] => familySource.list(session);

  /** Select every shape in the current scope (the slide, or the entered group). */
  const selectAll = (): void => {
    // In the slide sorter, Select All selects every slide.
    const count = session?.info.slides.length ?? 0;
    if (views.presentationView === 'slideSorter') {
      if (count) (instance.selectSlide(0), instance.selectSlide(count - 1, 'range'));
      return;
    }
    const lay = layout();
    if (!lay) return;
    const all = selectAllOf(lay, selection);
    // A text selection widens in place; shapes go through `select`.
    if (all.kind === 'text') return ((selection = all), emit());
    select(all);
  };

  const clipboard = createClipboardHost({
    session: () => session,
    slide: () => slide,
    selection: () => selection,
    layout,
    readOnly: () => mode === 'view',
    emit,
    deleteText: () => void textInput('deleteBackward'),
    cutShapes: () => removal.cut(() => exec({ type: 'delete' })),
    insertText: (text) => void textInput('text', text),
    refuse: (reason) => void refuse(reason),
    exec: (command) => void exec(command),
    selectedSlides: () => slideSel.indices(session?.info.slides ?? [], slide),
    ...(options.readClipboardDrawing ? { readDrawing: options.readClipboardDrawing } : {}),
  });

  /**
   * Put the mode in effect: the host's, or view while the deck is read-only. The host's mode is
   * kept, so the next deck opens in it.
   */
  const applyMode = (): void => {
    mode = readOnly || session?.damage ? 'view' : hostMode;
    if (mode !== 'view') return;
    painter.reset();
    cropper.cancel();
    // View mode edits no text and no cells: the element stays selected.
    if (selection.kind === 'text' || selection.kind === 'cells')
      selection = { kind: 'shapes', nodeIds: [selection.nodeId] };
  };
  /**
   * Refuse a file that did not open, or the deck `drop` opened when one of its slides did not
   * parse. A deck with no edits goes. A deck with edits stays open for viewing, as DOCX Editor's
   * write guard keeps a document readable: its store refuses every commit, and the notice says
   * why, so the user can still read and copy the work.
   */
  const refuseOpen = (code: string, drop?: DocumentSession): CommandResult => {
    const refused: CommandResult = { ok: false, reason: 'errors.malformed', detail: code };
    if (drop && drop !== session) return refused;
    if (drop && drop.store.edited) {
      // The deck opened and holds edits: it stays on screen, read-only. It is no failed open.
      readOnly = 'damaged-slide';
      applyMode();
      opening.keepReading(drop);
      emit();
      return refused;
    }
    if (drop) {
      (detachModules(), thumbnails.clear(), drop.dispose());
      // The deck goes with its fonts, its slide, and any gesture still open on it.
      fonts?.dispose();
      fonts = null;
      session = null;
      selection = { kind: 'none' };
      slide = 0;
      slideSel.reset(undefined);
      groups.closeAll();
    }
    // A deck that was open stays open. A load that runs keeps its own status.
    if (!drop || status !== 'loading') status = session ? 'ready' : 'error';
    failure = documentOpenFailure(code);
    emit();
    return { ok: false, reason: 'errors.malformed', detail: code };
  };
  const opening = createOpeningLane({
    session: () => session,
    emit,
    picturesChanged: () => (thumbnails.picturesChanged(), emit()),
    refuse: (owner, rejection) => void refuseOpen(openRejectionCode(rejection), owner),
    facesChanged: refreshFaces,
    requestFamilies: (families) => hostFonts.request(families),
  });

  const openFile: PptxEditorInstance['load'] = async (bytes, loadOptions) => {
    if (destroyed) return { ok: false, reason: 'disabled.destroyed' };
    // A later `load` (or `destroy`) supersedes this one: its result is disposed, not shown.
    const token = ++generation;
    const stale = (): boolean => destroyed || token !== generation;
    const superseded = (): CommandResult => ({
      ok: false,
      reason: destroyed ? 'disabled.destroyed' : 'errors.superseded',
    });
    // Any throw ends the open with a refusal to report: the loading screen never stays up.
    let phase = 'paint';
    // The new deck's session once it mounts: a throw after that drops it.
    let adopted: DocumentSession | null = null;
    try {
      status = 'loading';
      // The open deck's background parse waits: the new file gets the thread.
      opening.pause(session);
      emit();
      // A collaboration room needs every part, so a deck that joins one opens whole.
      const startSlide = clampSlide(loadOptions?.startSlide, Number.MAX_SAFE_INTEGER);
      const progressive = modules.collaboration ? {} : { progressive: { startSlide } };
      const openOptions = { ...openOptionsOf(options), ...progressive };
      const r = await prepareDeck(bytes, openOptions, hostFonts, stale, (s) => void (phase = s));
      if (r.kind === 'stale') return superseded();
      if (r.kind === 'refused') {
        opening.resume(session);
        return refuseOpen(openRejectionCode(r.error));
      }
      phase = 'mount';
      adopted = r.session;
      session?.dispose();
      fonts?.dispose();
      arm.clear();
      canCache.dispose();
      session = r.session;
      fonts = r.faces;
      const opened = session;
      detachModules();
      const detachSession = seam.attachEditorModules(modules, opened.store);
      const offPresence = presence.subscribe(() => surface?.render());
      presence.reset();
      detachModules = () => (offPresence(), detachSession());
      opened.store.subscribe(() => {
        cached = null;
        // A change made on the store directly (a host, a script): the editor did not set the
        // slide and the selection for it, so keep them valid and tell the chrome.
        if (committing > 0 || session !== opened) return;
        const at = opened.info.slides.findIndex((s) => s.partName === shownPart);
        if (at < 0) selection = { kind: 'none' };
        slide = at >= 0 ? at : Math.min(slide, Math.max(0, opened.info.slides.length - 1));
        if (at < 0) slideSel.reset(partAt(slide));
        arm.follow(() => {
          selection = presence.followRemote(opened.store.package, partAt(slide), selection);
          selection = validSelection(layout(), selection);
        });
        emit();
      });
      thumbnails.clear();
      status = 'ready';
      failure = undefined;
      readOnly = null;
      mode = hostMode;
      slide = clampSlide(startSlide, session.info.slides.length);
      slideSel.reset(session.info.slides[slide]?.partName);
      selection = { kind: 'none' };
      lastRejection = null;
      removal.clear();
      views.reset();
      painter.reset(true);
      groups.closeAll();
      emit();
      opening.start(opened);
      // Register faces that arrived before the mount. Shaped measurement can still be late.
      if (opened.fontGeneration !== r.generation) void refreshFaces(opened);
      hostFonts.awaitShaping(opened);
      return { ok: true };
    } catch (cause) {
      console.error(`[pptx-editor] The open failed (${phase})`, cause);
      if (adopted && session === adopted) {
        // The new deck was on screen, half set up: it goes, as a deck with a damaged part does.
        // The load is over, so the refusal sets the status the open leaves.
        status = 'ready';
        return refuseOpen(OPEN_FAILED, adopted);
      }
      if (adopted) (hostFonts.abandon(adopted), adopted.dispose());
      if (stale()) return superseded();
      opening.resume(session);
      return refuseOpen(OPEN_FAILED);
    }
  };

  const instance: PptxEditorInstance = {
    get session() {
      return session;
    },
    get fonts() {
      return hostFonts.current();
    },
    fontMeasurement: () => ({
      measurer: session && !session.fixedMeasurer ? 'shaped' : 'fixed',
      resolving: hostFonts.resolving(),
    }),
    get lastCopied() {
      return clipboard.lastCopied;
    },
    fontFamilies,
    tableStyles: () => (session ? tableStyleList(session.store.package) : []),
    setSelection: (next) => select(next),
    copy: clipboard.copy,
    selectAll,
    focus: () => surface?.focus(),
    input: (kind, data) => textInput(kind, data),
    find: createFind({
      session: () => session,
      slide: () => slide,
      selection: () => selection,
      show: (index, next) => (instance.goToSlide(index), select(next)),
      exec,
    }),
    snapshot,
    subscribe(listener) {
      listeners.add(listener);
      return () => listeners.delete(listener);
    },
    can,
    exec,
    beginHistoryGroup: () => groups.begin(),
    load: async (bytes, loadOptions) => {
      const before = session;
      try {
        return await openFile(bytes, loadOptions);
      } catch (error) {
        // The open threw before it replaced the deck: the old deck's background parse goes on.
        if (session === before) opening.resume(before);
        throw error;
      }
    },
    // A deck still opening parses the rest first; a slide that does not parse refuses the deck.
    save: () =>
      session && !session.damage && opening.finish() && session && !session.damage
        ? writePptx(repairTablesForSave(session.store.package))
        : null,
    setZoom(value) {
      if (value === 'fit') fitZoom = true;
      else {
        fitZoom = false;
        zoom = clampZoom(value);
      }
      emit();
    },
    setMode(next) {
      hostMode = next;
      applyMode();
      emit();
    },
    goToSlide(index) {
      instance.selectSlide(index);
    },
    selectSlide(index, mode = 'replace') {
      if (!session) return;
      const before = slide;
      const next = Math.max(0, Math.min(index, session.info.slides.length - 1));
      slide = slideSel.click(session.info.slides, slide, next, mode);
      if (slide !== before) {
        selection = { kind: 'none' };
        lastRejection = null;
        painter.reset();
        session.store.closeHistoryGroup();
      }
      emit();
    },
    attach(element) {
      surface?.destroy();
      const focus = (): void => surface?.focus();
      setSurfaceRoot(instance, element, focus, () => surface?.render());
      const base: SurfaceHostBase = {
        idPrefix,
        layout,
        label: (key, params) => (options.translate ? options.translate(key, params) : key),
        select,
        copy: clipboard.copyTo,
        paste: clipboard.pasteFrom,
        warmCopy: clipboard.warm,
        cropping,
        textInput,
        shortcut,
        run: (command, options) => exec(command, options),
        beginHistoryGroup: () => groups.begin(),
        refuse: (reason, detail) => void refuse(reason, detail),
        fitChanged: (value) => {
          if (Math.abs(value - fitValue) < 0.001) return;
          fitValue = value;
          if (fitZoom) {
            cached = null;
            surface?.render();
            for (const l of [...listeners]) l();
          }
        },
        zoom: effectiveZoom,
        setZoom: (value) => instance.setZoom(value),
        pickImage: () => pickImageInto(instance, element.ownerDocument),
      };
      surface = new SlideSurface(
        element,
        createSurfaceHost(base, {
          session: () => session,
          selection: () => selection,
          viewing: () => mode === 'view',
          part: () => partAt(slide),
          surface: () => surface,
          ...{ prompts, presence, painter, views, groups, gestureEdit },
        })
      );
      surface.render();
    },
    detach() {
      surface?.destroy();
      surface = null;
      setSurfaceRoot(instance, null);
    },
    setFonts: (next) => hostFonts.setFonts(next),
    setFontFamilies: (next) => void (familySource.set(next) && emit()),
    renderThumbnail: (index, container, opts) =>
      thumbnails.render(index, container, opts?.defer === true),
    renderLayoutThumbnail: (layoutId, container) => thumbnails.renderLayout(layoutId, container),
    finishOpening: () => !session || (!session.damage && opening.finish() && !session?.damage),
    reject: refuse,
    destroy() {
      // A `load` still in flight sees the flag and disposes what it opened.
      destroyed = true;
      generation += 1;
      surface?.destroy();
      surface = null;
      setSurfaceRoot(instance, null);
      thumbnails.clear();
      canCache.dispose();
      groups.closeAll();
      detachModules();
      hostFonts.dispose();
      session?.dispose();
      session = null;
      fonts?.dispose();
      fonts = null;
      status = 'empty';
      cached = null;
      listeners.clear();
    },
  };

  /** One drag step; the steps of one gesture are one undo step. A refusal says why. */
  function gestureEdit(command: EditorCommand, gesture: string, gestureBase?: PptxPackage): void {
    if (!session) return;
    const part = partAt(slide);
    const historyGroup = groups.drag(surface, gesture, session.store.package).token;
    const label = command.type === 'setTransforms' ? 'transform' : command.type;
    try {
      const { ops } = compile(
        { ...context(), ...(gestureBase ? { gestureBase } : {}) },
        command
      ).batch;
      const mark = markFor(part);
      const r = commit(ops, { label, historyGroup }, mark, mark);
      if (!r.ok) return void refuse('errors.unsupported', commitRefusalDetail(r.rejection));
    } catch (e) {
      const refusal = e instanceof CommandRefusal ? e : new CommandRefusal('errors.unsupported');
      // A drag step that changes nothing (a pointer that moved within one value) is no error.
      if (isNothingToDo(refusal.reason)) return;
      return void refuse(refusal.reason, refusal.detail);
    }
    emit();
  }

  if (options.document) void instance.load(options.document);
  return seam.withCollaborationFlag(removal.attach(instance), modules);
}
