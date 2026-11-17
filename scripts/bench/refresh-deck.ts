// The synthetic decks of the document refresh checks (`e2e/document-refresh.spec.ts`,
// `bun run bench:refresh`, and `bun run bench:refresh:browser`). Each call builds the same
// deck for the same arguments, so two calls differ only where an edit asks for it.

import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { buildPptx, para, run, shape } from '../fixtures/builder.ts';

const IN = 914_400;
const LOREM =
  'Quarterly results show steady growth across every region, with margins ahead of plan.';

/** The box of the `Body` shape, in EMU. */
export const BODY_BOX = { x: IN, y: 2 * IN, cx: 6 * IN, cy: 3 * IN } as const;
/** The slide size, in EMU. */
export const SLIDE_SIZE = { width: 12_192_000, height: 6_858_000 } as const;

/** The edits that a returned file can hold. */
export interface RefreshEdits {
  /** Slide indexes whose first body paragraph gets other text. */
  readonly text?: readonly number[];
  /**
   * Insert a new slide after this index. As a presentation processor does, the new slide is a
   * new part after the last one, and only its place in `p:sldIdLst` puts it after this index:
   * the existing slides keep their part names.
   */
  readonly insertAfter?: number;
  /** Slide indexes that change to the blank layout. */
  readonly layout?: readonly number[];
}

/** The first body paragraph of slide `index`, before and after a text edit. */
export const bodyText = (index: number, edited = false): string =>
  edited ? `Updated review date for slide ${index + 1}.` : `1. ${LOREM}`;

/** The title of the slide at original index `index`: the key of a slide across edits. */
export const titleText = (index: number): string => `Slide ${index + 1} title`;

function slide(index: number, edits: RefreshEdits, key = titleText(index)) {
  const edited = edits.text?.includes(index) ?? false;
  return {
    layout: edits.layout?.includes(index) ? ('blank' as const) : ('title' as const),
    shapes:
      shape({
        id: 2,
        name: 'Title',
        x: IN,
        y: 0.5 * IN,
        cx: 11 * IN,
        cy: IN,
        text: para(run(key, 'sz="3600"')),
      }) +
      shape({
        id: 3,
        name: 'Body',
        ...BODY_BOX,
        bodyPr: '<a:bodyPr wrap="square"/>',
        text: [bodyText(index, edited), `2. ${LOREM}`, `3. ${LOREM}`]
          .map((p) => para(run(p, 'sz="1800"')))
          .join(''),
      }),
  };
}

/** A deck of `n` text slides with `edits` applied. */
export function refreshDeck(n: number, edits: RefreshEdits = {}): Uint8Array {
  const slides = Array.from({ length: n }, (_, i) => slide(i, edits));
  const at = edits.insertAfter;
  if (at !== undefined) slides.push(slide(at, {}, 'Inserted slide'));
  const bytes = buildPptx({ slides, width: SLIDE_SIZE.width, height: SLIDE_SIZE.height });
  if (at === undefined || at >= n - 1) return bytes;
  // Move the last slide id to its place after slide `at`.
  const files = unzipSync(bytes);
  const name = 'ppt/presentation.xml';
  const xml = strFromU8(files[name]!).replace(
    /<p:sldIdLst>(.*?)<\/p:sldIdLst>/,
    (_, list: string) => {
      const ids = list.match(/<p:sldId [^>]*\/>/g) ?? [];
      const last = ids.pop()!;
      ids.splice(at + 1, 0, last);
      return `<p:sldIdLst>${ids.join('')}</p:sldIdLst>`;
    }
  );
  files[name] = strToU8(xml);
  return zipSync(files);
}
