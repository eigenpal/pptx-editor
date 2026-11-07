// Building blocks of the synthetic decks in `hang-decks.ts` and `ui-scale-decks.ts`.

import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { buildPptx, cell, shape } from '../fixtures/builder.ts';

export const IN = 914_400;
export const LOREM =
  'Quarterly results show steady growth across every region, with margins ahead of plan. ';

/** One deck of the set. */
export interface HangDeck {
  readonly id: string;
  /** What the deck tries to break. */
  readonly target: string;
  readonly bytes: () => Uint8Array;
  /** Open with a font origin that never answers, as a stalled font download does. */
  readonly stalledFonts?: boolean;
  /** This deck's slow mark in milliseconds, for a deck whose linear work is large. */
  readonly slowMs?: number;
}

export const one = (shapes: string, extra: Parameters<typeof buildPptx>[0] = { slides: [] }) =>
  buildPptx({ ...extra, slides: [{ shapes }] });

export const box = (id: number, text: string, o: Partial<Parameters<typeof shape>[0]> = {}) =>
  shape({ id, x: IN, y: IN, cx: 6 * IN, cy: 3 * IN, text, ...o });

/** The parts of a built deck, to edit: read a part's text, set it, or add a one-byte part. */
export interface DeckParts {
  get(name: string): string;
  set(name: string, xml: string): void;
  bin(name: string): void;
}

/** `bytes` with its parts edited by `edit`. */
export function patched(bytes: Uint8Array, edit: (parts: DeckParts) => void): Uint8Array {
  const files = unzipSync(bytes);
  edit({
    get: (name) => strFromU8(files[name]!),
    set: (name, xml) => void (files[name] = strToU8(xml)),
    bin: (name) => void (files[name] = new Uint8Array([0])),
  });
  return zipSync(files);
}

export const P_NS = 'http://schemas.openxmlformats.org/presentationml/2006/main';
export const MC_NS = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
export const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
export const range = (n: number): number[] => Array.from({ length: n }, (_, i) => i);

/** A placeholder shape of `type` with `idx`. */
export const placeholder = (id: number, type: string, idx: number): string =>
  `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="P${id}"/><p:cNvSpPr/><p:nvPr><p:ph type="${type}" idx="${idx}"/></p:nvPr></p:nvSpPr>` +
  '<p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>x</a:t></a:r></a:p></p:txBody></p:sp>';

export const SLIDE = 'ppt/slides/slide1.xml';

export const blipFill = (extra: string) =>
  `<a:blipFill><a:blip r:embed="rIdImg"/>${extra}</a:blipFill>`;

export const tableRows = (rows: number, cols: number, text = 'x') =>
  Array.from(
    { length: rows },
    () => `<a:tr h="${IN / 10}">${Array.from({ length: cols }, () => cell(text)).join('')}</a:tr>`
  ).join('');
