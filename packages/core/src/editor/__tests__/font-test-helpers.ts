// Font faces and decks for the font configuration tests. The faces are the packaged
// substitutes' own files (`@docx-editor.dev/fonts`), so every byte is a real sfnt.

import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { buildPptx, para, run, shape } from '../../../../../scripts/fixtures/builder.ts';
import type { FontSource } from '../../contracts/index.ts';
import { createFontSource } from '../load-fonts.ts';

export const FONT_ASSETS = join(
  dirname(Bun.resolveSync('@docx-editor.dev/fonts/package.json', import.meta.dir)),
  'assets'
);

export const fontBytes = (file: string): Uint8Array =>
  new Uint8Array(readFileSync(join(FONT_ASSETS, file)));

/** A host face for `family`, drawn with a packaged file. */
export function face(
  family: string,
  file = 'LiberationSans-Regular.ttf',
  weight = 400,
  style: 'normal' | 'italic' = 'normal'
): FontSource {
  const made = createFontSource(fontBytes(file), { family, weight, style });
  if (!('source' in made)) throw new Error(`fixture face ${file}: ${made.failure.reason}`);
  return made.source;
}

export const fixture = (name: string): Uint8Array =>
  new Uint8Array(readFileSync(join(import.meta.dir, '../../../../../fixtures', name)));

const IN = 914_400;

/**
 * A deck whose theme names `major` and `minor`, with one text box that names `latin` for its
 * Latin text and `ea` for its East Asian text.
 */
export function themedDeck(o: {
  major?: string;
  minor?: string;
  latin?: string;
  ea?: string;
}): Uint8Array {
  const children =
    (o.latin ? `<a:latin typeface="${o.latin}"/>` : '') +
    (o.ea ? `<a:ea typeface="${o.ea}"/>` : '');
  return buildPptx({
    ...(o.major ? { majorFont: o.major } : {}),
    ...(o.minor ? { minorFont: o.minor } : {}),
    slides: [
      {
        shapes: shape({
          id: 2,
          name: 'Text',
          x: IN,
          y: IN,
          cx: 6 * IN,
          cy: IN,
          text: para(run('Hello 世界', '', children)),
        }),
      },
    ],
  });
}

/** A promise and its resolve function. */
export function deferred<T>(): { promise: Promise<T>; resolve: (value: T) => void } {
  let resolve!: (value: T) => void;
  const promise = new Promise<T>((r) => (resolve = r));
  return { promise, resolve };
}

/** Wait for queued promise work to run. */
export const settle = async (turns = 5): Promise<void> => {
  for (let i = 0; i < turns; i += 1) await new Promise((r) => setTimeout(r, 0));
};
