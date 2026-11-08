// The slide master view (`presentationView: 'slideMaster'`): the canvas edits a slide layout or
// a slide master as a slide, and the rail lists each master with its layouts
// (`templateRefs`). Every slide that uses the layout shows the edit. The view's current "slide"
// is a template index (`TEMPLATE_BASE`), so the cascade, the session, and the commands read the
// template part through the same index lookup (`slideRefAt`). Commands that change slides, their
// order, notes, animations, transitions, comments, or sections are off here: they act on slides,
// which the view does not show.

import type {
  EditorCommand,
  EditorSelection,
  EditorSnapshot,
  PresentationView,
} from '../contracts/index.ts';
import type { DocumentSession } from './document-session.ts';
import { validSelection, type HistoryMark } from './history-marks.ts';
import { renderKey } from './slide-dependencies.ts';
import {
  isTemplateIndex,
  presentationInfo,
  templateIndexOf,
  templateRefs,
  TEMPLATE_BASE,
  type PptxPackage,
} from '../store/index.ts';

const NONE: EditorSelection = { kind: 'none' };

/** Commands that act on slides, not on the shapes of the template on the canvas. */
const SLIDE_ONLY: ReadonlySet<EditorCommand['type']> = new Set<EditorCommand['type']>([
  'addAnimation',
  'moveAnimation',
  'removeAnimation',
  'setAnimationTiming',
  'setTransition',
  'addComment',
  'deleteComment',
  'editComment',
  'reopenComment',
  'replyToComment',
  'resolveComment',
  'addSection',
  'moveSection',
  'removeAllSections',
  'removeSection',
  'renameSection',
  'applyLayout',
  'resetSlides',
  'setNotes',
  'formatNotes',
  'setFooter',
  'setSlideNumbers',
  'setDateTime',
  'setSlideSize',
  'setSlidesHidden',
  'replace',
  'replaceAll',
  'pasteSlides',
]);

/** Slide commands that act on the template on the canvas in the slide master view. */
const TEMPLATE_EDITS: ReadonlySet<EditorCommand['type']> = new Set<EditorCommand['type']>([
  'setSlideBackground',
]);

/**
 * Whether `command` cannot run in the slide master view: a command on slides, or any other
 * slide-level edit (`slideEdit`) that does not act on the template on the canvas.
 */
export const slideOnlyCommand = (command: EditorCommand, slideEdit = false): boolean =>
  !TEMPLATE_EDITS.has(command.type) && (slideEdit || SLIDE_ONLY.has(command.type));

/** The slide to return to, and the template to show, as the view opens and closes. */
export class MasterView {
  /** The slide shown when the view opened: its index, and its part, which an edit keeps. */
  #slide = 0;
  #part: string | undefined;

  /**
   * Open the view from slide `slide`: the view shows that slide's layout, or the first template
   * when the slide names none. Returns the template index to show, or null for a deck with no
   * master.
   */
  open(pkg: PptxPackage, slide: number): number | null {
    if (!templateRefs(pkg).length) return null;
    this.#slide = slide;
    this.#part = presentationInfo(pkg).slides[slide]?.partName;
    const layout = presentationInfo(pkg).slides[slide]?.layoutPart;
    const at = layout ? templateIndexOf(pkg, layout) : -1;
    return at >= 0 ? at : TEMPLATE_BASE;
  }

  /**
   * Close the view: the slide shown when it opened, wherever an edit moved it, else its index
   * kept in the deck's range.
   */
  close(pkg: PptxPackage): number {
    const at = this.#part ? this.indexOf(pkg, 'normal', this.#part) : -1;
    if (at >= 0) return at;
    const count = presentationInfo(pkg).slides.length;
    return Math.max(0, Math.min(this.#slide, count - 1));
  }

  /**
   * The template index a view change shows: the opened template when `command` opens the view,
   * the slide it opened from when it closes it, and undefined when it does neither. Null when
   * the deck has no master to show.
   */
  target(
    session: DocumentSession | null,
    view: PresentationView,
    command: EditorCommand,
    slide: number
  ): number | null | undefined {
    if (!session || command.type !== 'setPresentationView' || command.view === view) return;
    const pkg = session.store.package;
    if (command.view === 'slideMaster') return this.open(pkg, slide);
    return view === 'slideMaster' ? this.close(pkg) : undefined;
  }

  /**
   * The index of the slide or template `part` in `view`, or -1: the master view shows
   * templates only, and the other views show slides only. An undo of an edit made in the other
   * view keeps the current slide.
   */
  indexOf(pkg: PptxPackage, view: PresentationView, part: string | undefined): number {
    if (!part) return -1;
    if (view === 'slideMaster') return templateIndexOf(pkg, part);
    return presentationInfo(pkg).slides.findIndex((s) => s.partName === part);
  }

  /**
   * Why a view command cannot run here, or null: the master view needs a deck with a master.
   */
  refused(
    pkg: PptxPackage | undefined,
    command: EditorCommand
  ): { ok: false; reason: string } | null {
    if (command.type !== 'setPresentationView' || command.view !== 'slideMaster') return null;
    return pkg && templateRefs(pkg).length ? null : { ok: false, reason: 'disabled.noSlideMaster' };
  }

  /**
   * The slide and the selection that an undo or a redo of an edit marked `mark` shows: the
   * edited slide or template and its selection when the view shows it, else the current one
   * with no selection.
   */
  restore(
    session: DocumentSession,
    view: PresentationView,
    mark: HistoryMark | undefined,
    current: number
  ): { slide: number; selection: EditorSelection } {
    const pkg = session.store.package;
    const index = this.indexOf(pkg, view, mark?.part);
    if (!mark || index < 0) return { slide: this.kept(pkg, view, current), selection: NONE };
    return { slide: index, selection: validSelection(session.layout(index), mark.selection) };
  }

  /**
   * The parts that `view` lists, in order, and the index of the first: the templates from
   * `TEMPLATE_BASE` in the master view, else the slides from 0. A collaboration undo finds the
   * part it changed in this list.
   */
  parts(pkg: PptxPackage, view: PresentationView): { parts: readonly string[]; base: number } {
    if (view === 'slideMaster')
      return { parts: templateRefs(pkg).map((t) => t.partName), base: TEMPLATE_BASE };
    return { parts: presentationInfo(pkg).slides.map((s) => s.partName), base: 0 };
  }

  /** The current slide or template, kept in the range of `view` after an edit. */
  kept(pkg: PptxPackage, view: PresentationView, slide: number): number {
    if (view === 'slideMaster') return this.clamp(pkg, slide, TEMPLATE_BASE);
    return Math.min(slide, Math.max(0, presentationInfo(pkg).slides.length - 1));
  }

  /**
   * The template a request for `index` shows: a template index in the deck's range, or the
   * current one. A slide index does not leave the view.
   */
  clamp(pkg: PptxPackage, index: number, current: number): number {
    const count = templateRefs(pkg).length;
    if (!isTemplateIndex(index) || !count) return current;
    return Math.min(index, TEMPLATE_BASE + count - 1);
  }
}

type SlideSummary = EditorSnapshot['slides'][number];

/**
 * The rail of the slide master view: each master, then its layouts, as slide summaries with a
 * template index, their kind, and their name. A summary keeps its object while its template
 * renders the same, so the rail repaints only the thumbnails an edit changed.
 */
export function createTemplateSummaries(): (
  session: DocumentSession | null,
  pkg: PptxPackage | undefined
) => EditorSnapshot['slides'] {
  let last: { pkg: PptxPackage | undefined; fonts: string; out: EditorSnapshot['slides'] } = {
    pkg: undefined,
    fonts: '',
    out: [],
  };
  return (session, pkg) => {
    const fonts = session?.paintGeneration ?? '';
    if (last.pkg === pkg && last.fonts === fonts) return last.out;
    if (!session || !pkg) return (last = { pkg, fonts, out: [] }).out;
    const before = new Map(last.out.map((s) => [s.id, s]));
    const out: SlideSummary[] = [];
    // How many slides use each layout, and each master through its layouts.
    const used = new Map<string, number>();
    for (const s of presentationInfo(pkg).slides)
      for (const part of [s.layoutPart, s.masterPart])
        if (part) used.set(part, (used.get(part) ?? 0) + 1);
    templateRefs(pkg).forEach((t, k) => {
      const index = TEMPLATE_BASE + k;
      const key = renderKey(pkg, index, session.paintKey(t.partName), []) ?? '';
      const old = before.get(t.partName);
      const usedBy = used.get(t.partName) ?? 0;
      const same =
        old &&
        old.index === index &&
        old.key === key &&
        old.name === t.name &&
        old.usedBy === usedBy;
      out.push(
        same
          ? old
          : {
              index,
              id: t.partName,
              hidden: false,
              key,
              template: t.template,
              name: t.name,
              usedBy,
              ...(t.masterPart ? { masterId: t.masterPart } : {}),
            }
      );
    });
    last = { pkg, fonts, out };
    return last.out;
  };
}

type SlideSummaries = EditorSnapshot['slides'];

/**
 * The rail's summaries in `view`: the slides (`slides`), or the templates in the slide master
 * view.
 */
export function railSummaries(
  slides: (session: DocumentSession | null, pkg: PptxPackage | undefined) => SlideSummaries
): (
  view: PresentationView,
  session: DocumentSession | null,
  pkg: PptxPackage | undefined
) => SlideSummaries {
  const templates = createTemplateSummaries();
  return (view, session, pkg) =>
    view === 'slideMaster' ? templates(session, pkg) : slides(session, pkg);
}
