// Insert > Footer and Insert > Date and time.

import { describe, expect, test } from 'bun:test';
import { buildPptx, para, run, shape } from '../../../../../scripts/fixtures/builder.ts';
import { headerFooterDeck } from '../../../../../scripts/fixtures/round-trip.ts';
import { presentationInfo, serializeXmlPart, xmlPartOf } from '../../store/index.ts';
import { shortDate } from '../../automation/commands/header-footer.ts';
import { createPptxEditor, type PptxEditorInstance } from '../index.ts';

async function open(bytes: Uint8Array): Promise<PptxEditorInstance> {
  const e = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await e.load(bytes)).ok).toBe(true);
  return e;
}

const slideXml = (e: PptxEditorInstance, index: number): string => {
  const pkg = e.session!.store.package;
  return serializeXmlPart(xmlPartOf(pkg, presentationInfo(pkg).slides[index]!.partName)!);
};

/** The texts that the layout of slide `index` paints. */
const painted = (e: PptxEditorInstance, index: number): string[] =>
  e
    .session!.layout(index)!
    .items.flatMap((item) =>
      item.kind === 'shape' && item.text
        ? item.text.lines.map((l) => l.runs.map((r) => r.text).join(''))
        : []
    );

describe('Insert > Footer', () => {
  test('gives every slide the footer and keeps the format of an existing one', async () => {
    const e = await open(buildPptx(headerFooterDeck()));
    expect(e.snapshot().footer).toBe('Footer text & more');
    expect(e.exec({ type: 'setFooter', text: 'Quarterly <results>' }).ok).toBe(true);
    expect(e.snapshot().footer).toBe('Quarterly <results>');
    // Slide 1 keeps its own box, alignment, and size.
    const first = slideXml(e, 0);
    expect(first).toContain('<a:pPr algn="ctr"/>');
    expect(first).toContain('sz="1200"');
    expect(first).toContain('Quarterly &lt;results&gt;');
    expect(first).not.toContain('Footer text');
    // Slide 2 gets a placeholder from its layout, and paints the text there.
    expect(slideXml(e, 1)).toContain('type="ftr"');
    expect(painted(e, 1)).toContain('Quarterly <results>');
    // The save reopens with the same footer.
    const again = await open(e.save()!);
    again.selectSlide(1);
    expect(again.snapshot().footer).toBe('Quarterly <results>');
  });

  test('null removes the footer from every slide, and undo restores it', async () => {
    const e = await open(buildPptx(headerFooterDeck()));
    expect(e.exec({ type: 'setFooter', text: null }).ok).toBe(true);
    expect(e.snapshot().footer).toBeNull();
    expect(slideXml(e, 0)).not.toContain('type="ftr"');
    e.exec({ type: 'undo' });
    expect(e.snapshot().footer).toBe('Footer text & more');
  });

  test('refuses text that is too long or not valid XML', async () => {
    const e = await open(buildPptx(headerFooterDeck()));
    expect(e.exec({ type: 'setFooter', text: 'x'.repeat(1025) }).ok).toBe(false);
    expect(e.exec({ type: 'setFooter', text: 'bad \u0001' }).ok).toBe(false);
    expect(e.snapshot().footer).toBe('Footer text & more');
  });
});

describe('Insert > Date and time', () => {
  test('a field date on every slide, then a fixed date in its place', async () => {
    const e = await open(buildPptx(headerFooterDeck()));
    expect(e.snapshot().dateTime).toBe(false);
    expect(e.exec({ type: 'setDateTime', show: true, text: '11/7/2026' }).ok).toBe(true);
    expect(e.snapshot().dateTime).toBe(true);
    for (const i of [0, 1]) {
      expect(slideXml(e, i)).toMatch(/<a:fld id="\{[0-9A-F-]{36}\}" type="datetime1">/);
      expect(painted(e, i)).toContain('11/7/2026');
    }
    expect(e.exec({ type: 'setDateTime', show: true, text: '7 November', fixed: true }).ok).toBe(
      true
    );
    expect(slideXml(e, 0)).not.toContain('datetime1');
    expect(painted(e, 0)).toContain('7 November');
    expect(e.exec({ type: 'setDateTime', show: false }).ok).toBe(true);
    expect(e.snapshot().dateTime).toBe(false);
  });

  test('shortDate writes month/day/year', () => {
    expect(shortDate(new Date(2026, 10, 7))).toBe('11/7/2026');
  });
});

test('a deck without footer and date placeholders refuses with a reason', async () => {
  const e = await open(
    buildPptx({
      slides: [
        { shapes: shape({ id: 2, x: 0, y: 0, cx: 914400, cy: 914400, text: para(run('S')) }) },
      ],
    })
  );
  const footer = e.exec({ type: 'setFooter', text: 'Footer' });
  expect(footer.ok ? undefined : footer.reason).toBe('disabled.noFooterPlaceholder');
  const date = e.exec({ type: 'setDateTime', show: true });
  expect(date.ok ? undefined : date.reason).toBe('disabled.noDatePlaceholder');
  // Hiding what the deck cannot show is refused too, so `can` offers no empty step.
  expect(e.can({ type: 'setFooter', text: null }).ok).toBe(false);
  expect(e.can({ type: 'setDateTime', show: false }).ok).toBe(false);
});

/** A placeholder of `type` with `idx`, and `body` as its text body (none when empty). */
const placeholder = (id: number, type: string, idx: number, body: string): string =>
  `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${type} ${id}"/><p:cNvSpPr/>` +
  `<p:nvPr><p:ph type="${type}" idx="${idx}"/></p:nvPr></p:nvSpPr><p:spPr/>${body}</p:sp>`;

/** The header-footer deck, with `shapes` on slide 2. */
const withSlide2 = (shapes: string): Uint8Array => {
  const deck = headerFooterDeck();
  return buildPptx({ ...deck, slides: [deck.slides[0]!, { ...deck.slides[1]!, shapes }] });
};

describe('existing placeholders', () => {
  test('a footer without a text body gets one, after the shape properties', async () => {
    const e = await open(withSlide2(placeholder(4, 'ftr', 11, '')));
    expect(e.exec({ type: 'setFooter', text: 'Hello' }).ok).toBe(true);
    expect(slideXml(e, 1)).toMatch(/<p:spPr\/><p:txBody>.*Hello.*<\/p:txBody><\/p:sp>/);
    expect(painted(e, 1)).toContain('Hello');
  });

  test('a footer keeps its first paragraph only, with its format', async () => {
    const body =
      '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:pPr algn="r"/><a:r><a:rPr lang="de-DE" b="1"/>' +
      '<a:t>One</a:t></a:r><a:br/><a:r><a:rPr lang="de-DE"/><a:t>Two</a:t></a:r></a:p>' +
      '<a:p><a:r><a:t>Three</a:t></a:r></a:p></p:txBody>';
    const e = await open(withSlide2(placeholder(4, 'ftr', 11, body)));
    expect(e.exec({ type: 'setFooter', text: 'New' }).ok).toBe(true);
    const xml = slideXml(e, 1);
    expect(xml).toContain(
      '<a:p><a:pPr algn="r"/><a:r><a:rPr lang="de-DE" b="1"/><a:t>New</a:t></a:r></a:p>'
    );
    expect(xml).not.toMatch(/Two|Three|<a:br/);
  });

  test('a fixed date that becomes a field keeps its character format', async () => {
    const body =
      '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="de-DE" sz="900" b="1"/>' +
      '<a:t>7.11.2026</a:t></a:r></a:p></p:txBody>';
    const e = await open(withSlide2(placeholder(4, 'dt', 10, body)));
    expect(e.exec({ type: 'setDateTime', show: true, text: '11/7/2026' }).ok).toBe(true);
    expect(slideXml(e, 1)).toMatch(
      /<a:fld id="\{[0-9A-F-]{36}\}" type="datetime1"><a:rPr lang="de-DE" sz="900" b="1"\/><a:t>11\/7\/2026<\/a:t><\/a:fld>/
    );
    expect(slideXml(e, 1)).not.toContain('7.11.2026');
  });

  test('a date field of another format takes the format of its new text', async () => {
    const body =
      '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:fld id="{B6F15528-21DE-4FAA-801E-634DDDAF4B2C}" ' +
      'type="datetime4"><a:rPr lang="en-US"/><a:t>November 7, 2026</a:t></a:fld></a:p></p:txBody>';
    const e = await open(withSlide2(placeholder(4, 'dt', 10, body)));
    expect(e.exec({ type: 'setDateTime', show: true, text: '11/7/2026' }).ok).toBe(true);
    const xml = slideXml(e, 1);
    expect(xml).toContain('id="{B6F15528-21DE-4FAA-801E-634DDDAF4B2C}" type="datetime1"');
    expect(xml).toContain('<a:t>11/7/2026</a:t>');
  });
});

test('a layout without a footer placeholder uses the one of its master', async () => {
  const deck = headerFooterDeck();
  const e = await open(
    buildPptx({
      ...deck,
      layouts: [{ name: 'Bare', shapes: '' }],
      slides: [{ ...deck.slides[1]!, layoutName: 'Bare' }],
    })
  );
  expect(e.exec({ type: 'setFooter', text: 'From the master' }).ok).toBe(true);
  // The master's footer placeholder has idx 3.
  expect(slideXml(e, 0)).toContain('<p:ph type="ftr" sz="quarter" idx="3"/>');
  expect(painted(e, 0)).toContain('From the master');
});

test('the menu state is the deck state: a slide without a footer reads the others', async () => {
  const e = await open(buildPptx(headerFooterDeck()));
  e.selectSlide(1);
  expect(e.snapshot().footer).toBe('Footer text & more');
  expect(e.snapshot().dateTime).toBe(false);
});
