// The faces that an open loads first, and the faces that the host's font origins hear: the
// faces that the text asks for (`face-styles.ts`), not all four faces of each family.

import { afterEach, expect, test } from 'bun:test';
import { hold, release, requestedFaces, resetGate, waitFor } from './substitute-gate.ts';

const { buildPptx, para, run, shape } = await import('../../../../../scripts/fixtures/builder.ts');
const { createPptxEditor, defineFontResolver } = await import('../index.ts');
const { openPresentation } = await import('../document-session.ts');
const { partStyles } = await import('../face-styles.ts');
const { parseXmlPart } = await import('../../store/index.ts');

afterEach(resetGate);

const IN = 914_400;

/** A deck with the theme font `theme` and one text box per slide, each with these runs. */
function deck(theme: string, slides: readonly string[]): Uint8Array {
  return buildPptx({
    majorFont: theme,
    minorFont: theme,
    slides: slides.map((runs, i) => ({
      shapes: shape({ id: 2 + i, x: IN, y: IN, cx: 6 * IN, cy: IN, text: para(runs) }),
    })),
  });
}

test('bold text on the opening slide loads its bold face before the open, and italic after', async () => {
  const bytes = deck('Courier New', [run('Title', 'b="1"'), run('Aside', 'i="1"')]);
  const opened = await openPresentation(bytes, { shapingWaitMs: 5000, firstSlideFonts: 0 });
  if (!opened.ok) throw new Error('open failed');
  // The master's bullet font (`a:buFont`) names Arial, so the slide uses its faces too.
  const opening = requestedFaces.slice(0, 4).sort();
  expect(opening).toEqual([
    'Arial 400 normal',
    'Arial 700 normal',
    'Courier New 400 normal',
    'Courier New 700 normal',
  ]);
  expect(await opened.session.restSubstitutes).toBe(true);
  expect(requestedFaces.indexOf('Courier New 400 italic')).toBeGreaterThan(3);
  expect(requestedFaces.indexOf('Courier New 700 italic')).toBeGreaterThan(3);
  opened.session.dispose();
});

test('the other faces of the families a slide has install late without holding the notice', async () => {
  hold('Century Gothic 700 normal');
  const { createHostFonts } = await import('../host-fonts.ts');
  const opened = await openPresentation(deck('Century Gothic', [run('Plain'), run('Plain')]), {
    shapingWaitMs: 60_000,
    firstSlideFonts: 0,
  });
  if (!opened.ok) throw new Error('open failed');
  const session = opened.session;
  expect(session.restAddsFamilies).toBe(false);
  let installs = 0;
  const fonts = createHostFonts(
    {},
    {
      session: () => session,
      installed: async () => void (installs += 1),
      changed: () => {},
    }
  );
  fonts.awaitShaping(session);
  // Only faces of families the deck already has are on the way, so the notice can show.
  expect(fonts.resolving()).toBe(false);
  release('Century Gothic 700 normal');
  expect(await session.restSubstitutes).toBe(true);
  await waitFor(() => installs === 1);
  expect(
    session.measurer.resolve({ family: 'Century Gothic', bold: true, italic: false })
  ).toMatchObject({ bold: true });
  fonts.dispose();
  session.dispose();
});

test('a host origin hears the faces that the slides naming each family ask for', async () => {
  const bytes = deck('Cambria', [
    run('Bold', 'b="1"', '<a:latin typeface="Roboto"/>'),
    run('Plain', '', '<a:latin typeface="Lato"/>'),
  ]);
  const heard: { family: string; weight: number; style: string }[][] = [];
  const editor = createPptxEditor({
    fonts: defineFontResolver((request) => {
      heard.push([...(request.faces ?? [])]);
      return undefined;
    }),
  });
  expect((await editor.load(bytes)).ok).toBe(true);
  const faces = heard[0]!;
  const of = (family: string) =>
    faces
      .filter((f) => f.family === family)
      .map((f) => `${f.weight} ${f.style}`)
      .sort();
  expect(of('Roboto')).toEqual(['400 normal', '700 normal']);
  expect(of('Lato')).toEqual(['400 normal']);
  // The theme font is the default family: it gets the faces of the whole deck.
  expect(of('Cambria')).toEqual(['400 normal', '700 normal']);
  editor.destroy();
});

test('a table or a table style that sets bold or italic counts as that face', () => {
  const styles = (xml: string) => {
    const parsed = parseXmlPart(
      '/ppt/slides/slide1.xml',
      `<a:root xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">${xml}</a:root>`
    );
    if (!parsed.ok) throw new Error(parsed.reason);
    return partStyles(parsed.part);
  };
  expect(styles('<a:rPr lang="en-US"/>')).toEqual({ bold: false, italic: false });
  expect(styles('<a:tbl/>')).toEqual({ bold: true, italic: false });
  expect(styles('<a:tcTxStyle i="on"/>')).toEqual({ bold: false, italic: true });
  expect(styles('<a:defRPr b="0" i="1"/>')).toEqual({ bold: false, italic: true });
});
