// Placeholder prompt text ("Click to add title"): what the Normal view shows in an empty
// placeholder (`parity-baseline` 4.8). It is editing chrome, so only the editor surface paints
// it: never the rail thumbnails, the slide show, the fidelity harness, or a saved file.
//
// The prompt goes into a copy of the package through the same text input as typing
// (`insertText`), and that copy is laid out with the slide's own cascade. So the prompt shows in
// the style typed text takes, and the copy is never committed.

import { createTranslator } from '@pptx-editor.dev/i18n';
import {
  attr,
  boolAttr,
  child,
  elements,
  path,
  presentationInfo,
  shapeTreeOf,
  textOf,
  type Element,
  type PptxPackage,
} from '../store/index.ts';
import { layoutPlaceholderOf, resolveSlide } from '../cascade/index.ts';
import {
  layoutSlide,
  type LayoutItem,
  type SlideLayout,
  type TextMeasurer,
} from '../layout/index.ts';
import type { EditorSelection } from '../contracts/index.ts';
import type { DocumentSession } from './document-session.ts';
import { CommandRefusal, OpBatch } from '../automation/commands/op-batch.ts';
import { insertText, paragraphsOf, splitParagraph } from '../automation/commands/text-model.ts';
import { addSlide, FOOTER_PLACEHOLDERS } from '../automation/commands/slide-commands.ts';
import { deckLayouts } from '../automation/commands/slide-layouts.ts';
import { emptyText } from '../automation/commands/placeholder-fill.ts';
import { warnUnknownLayout } from './dev-warnings.ts';

/** The prompt key of each placeholder type (§19.7.10 `ST_PlaceholderType`). */
const PROMPT_KEYS: Readonly<Record<string, string>> = {
  title: 'canvas.prompts.title',
  ctrTitle: 'canvas.prompts.title',
  subTitle: 'canvas.prompts.subtitle',
  body: 'canvas.prompts.text',
  obj: 'canvas.prompts.text',
  pic: 'canvas.prompts.picture',
  clipArt: 'canvas.prompts.picture',
  tbl: 'canvas.prompts.table',
  chart: 'canvas.prompts.chart',
  dgm: 'canvas.prompts.diagram',
  media: 'canvas.prompts.media',
};

const PROMPT_KEY_LIST: readonly string[] = [...new Set(Object.values(PROMPT_KEYS))];

/**
 * The prompt paragraphs of a slide placeholder: the layout placeholder's own text when it has a
 * custom prompt (`hasCustomPrompt`, §19.3.1.36), otherwise the default prompt for its type.
 */
function promptOf(
  layoutPh: Element | undefined,
  type: string,
  t: (key: string) => string
): string[] {
  const ph = layoutPh ? path(elements(layoutPh)[0], 'p:nvPr', 'p:ph') : undefined;
  const body = child(layoutPh, 'p:txBody');
  if (body && boolAttr(ph, 'hasCustomPrompt')) {
    const lines = paragraphsOf(body).map((p) =>
      elements(p)
        .map((r) => textOf(child(r, 'a:t')))
        .join('')
    );
    if (lines.some((l) => l)) return lines;
  }
  // `type` comes from the file: read only the table's own keys (`constructor` is not one).
  const key = Object.hasOwn(PROMPT_KEYS, type) ? PROMPT_KEYS[type] : undefined;
  return key ? [t(key)] : [];
}

interface Painted {
  readonly layout: SlideLayout;
  readonly editing: string | null;
  readonly fonts: number;
  /** The prompt strings the translator gave: a locale change repaints. */
  readonly strings: string;
  readonly out: SlideLayout;
}

/** Builds the layout the surface paints: the slide's, with prompts in empty placeholders. */
export class PlaceholderPrompts {
  readonly #t: (key: string) => string;
  #last: Painted | null = null;

  constructor(translate?: (key: string) => string) {
    const english = createTranslator();
    // A host translator that does not know a key returns the key: fall back to English.
    this.#t = (key) => {
      const text = translate?.(key);
      return text && text !== key ? text : english(key as Parameters<typeof english>[0]);
    };
  }

  /**
   * `layout` with each empty placeholder's prompt laid out as its text. The placeholder being
   * edited shows none (the caret shows instead), and viewing mode shows none.
   */
  paint(
    session: DocumentSession | null,
    layout: SlideLayout | undefined,
    selection: EditorSelection,
    readOnly: boolean
  ): SlideLayout | undefined {
    if (!session || !layout || readOnly) return layout;
    const editing = selection.kind === 'text' ? selection.nodeId : null;
    // The host translator can answer differently after a locale change, with the same identity.
    const strings = PROMPT_KEY_LIST.map(this.#t).join('\u0000');
    const last = this.#last;
    if (
      last &&
      last.layout === layout &&
      last.editing === editing &&
      last.fonts === session.fontGeneration &&
      last.strings === strings
    )
      return last.out;
    const out = this.#build(session, layout, editing) ?? layout;
    this.#last = { layout, editing, fonts: session.fontGeneration, strings, out };
    return out;
  }

  /**
   * How a new slide on `layoutPart` looks, with its prompts: what the Layout and New Slide
   * galleries show. The slide is added to a copy of the package that is never committed.
   */
  preview(session: DocumentSession, layoutPart: string): SlideLayout | undefined {
    if (!deckLayouts(session.store.package).some((l) => l.part === layoutPart)) {
      warnUnknownLayout(layoutPart);
      return undefined;
    }
    const batch = new OpBatch(session.store.package);
    let index: number;
    try {
      index = addSlide(batch, presentationInfo(batch.pkg).slides.length - 1, layoutPart);
    } catch (e) {
      if (e instanceof CommandRefusal) return undefined;
      throw e;
    }
    const resolved = resolveSlide(batch.pkg, index);
    if (!resolved) return undefined;
    const layout = layoutSlide(resolved, session.measurer);
    return withPrompts(batch.pkg, session.measurer, layout, this.#t, null) ?? layout;
  }

  #build(
    session: DocumentSession,
    layout: SlideLayout,
    editing: string | null
  ): SlideLayout | undefined {
    return withPrompts(session.store.package, session.measurer, layout, this.#t, editing);
  }
}

/**
 * `layout` (of a slide of `pkg`) with each empty placeholder's prompt laid out as its text,
 * except the placeholder `editing`. Undefined when no placeholder shows a prompt.
 */
export function withPrompts(
  pkg: PptxPackage,
  measurer: TextMeasurer,
  layout: SlideLayout,
  t: (key: string) => string,
  editing: string | null
): SlideLayout | undefined {
  const index = layout.slide.index;
  const part = presentationInfo(pkg).slides[index]?.partName;
  if (!part || part !== layout.slide.partName) return undefined;
  const batch = new OpBatch(pkg);
  const prompted = new Set<string>();
  try {
    for (const el of elements(shapeTreeOf(pkg, part))) {
      const ph = path(el, 'p:nvSpPr', 'p:nvPr', 'p:ph');
      const txBody = child(el, 'p:txBody');
      const type = attr(ph, 'type') ?? 'obj';
      if (el.kind !== 'sp' || !ph || !txBody || el.id === editing) continue;
      if (FOOTER_PLACEHOLDERS.has(type) || !emptyText(txBody)) continue;
      if (boolAttr(path(el, 'p:nvSpPr', 'p:cNvPr'), 'hidden')) continue;
      const lines = promptOf(layoutPlaceholderOf(pkg, index, el), type, t);
      if (!lines.length || !paragraphsOf(txBody).length) continue;
      let at = { paragraph: 0, offset: 0 };
      lines.forEach((line, i) => {
        if (i > 0) at = splitParagraph(batch, part, txBody.id, at);
        if (line) at = insertText(batch, part, txBody.id, at, line);
      });
      prompted.add(el.id);
    }
  } catch (e) {
    // A body the text model refuses shows no prompt; the slide paints as it is.
    if (e instanceof CommandRefusal) return undefined;
    throw e;
  }
  if (!prompted.size) return undefined;
  const resolved = resolveSlide(batch.pkg, index);
  if (!resolved) return undefined;
  const byId = new Map<string, LayoutItem>();
  for (const item of layoutSlide(resolved, measurer).items)
    if (prompted.has(item.element.nodeId)) byId.set(item.element.nodeId, item);
  const items = layout.items.map((item) => {
    const p = byId.get(item.element.nodeId);
    if (item.kind !== 'shape' || p?.kind !== 'shape' || !p.text) return item;
    return { ...item, text: p.text, ...(p.textMatrix ? { textMatrix: p.textMatrix } : {}) };
  });
  return { ...layout, items };
}
