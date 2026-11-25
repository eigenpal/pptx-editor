// `readHeaderFooter` and `setHeaderFooter` through the automation host, over the editor's
// Insert > Slide numbers, Footer, and Date and time commands. The placeholders are `p:ph`
// (`CT_Placeholder`, ECMA-376 Part 1, 19.3.1.36) of type `sldNum`, `ftr`, and `dt`
// (`ST_PlaceholderType`, 19.7.10); a date is fixed text or an `a:fld` (`CT_TextField`,
// 21.1.2.2.4). `p:hf` (`CT_HeaderFooter`, 19.3.1.25) stays as it is.

import { describe, expect, test } from 'bun:test';
import { buildPptx } from '../../../../../scripts/fixtures/builder.ts';
import { headerFooterDeck } from '../../../../../scripts/fixtures/round-trip.ts';
import {
  PptxDocumentStore,
  presentationInfo,
  readPptx,
  serializeXmlPart,
  writePptx,
  xmlPartOf,
} from '../../store/index.ts';
import { shortDate } from '../commands/header-footer.ts';
import { createStoreAutomationHost, type AutomationOperation } from '../index.ts';

const DATE_PH = (body: string): string =>
  '<p:sp><p:nvSpPr><p:cNvPr id="7" name="Date Placeholder 6"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>' +
  '<p:nvPr><p:ph type="dt" sz="half" idx="10"/></p:nvPr></p:nvSpPr><p:spPr/>' +
  `<p:txBody><a:bodyPr/><a:lstStyle/><a:p>${body}</a:p></p:txBody></p:sp>`;

/** The header-footer deck, with `first` and `second` added to its two slides. */
function setup(first = '', second = '') {
  const deck = headerFooterDeck();
  const read = readPptx(
    buildPptx({
      ...deck,
      slides: [
        { ...deck.slides[0]!, shapes: deck.slides[0]!.shapes + first },
        { ...deck.slides[1]!, shapes: deck.slides[1]!.shapes + second },
      ],
    })
  );
  if (!read.ok) throw new Error('fixture');
  const store = new PptxDocumentStore(read.pkg);
  return { store, host: createStoreAutomationHost(store) };
}

type Host = ReturnType<typeof setup>['host'];

const slide = (index: number): AutomationOperation[] => [
  { op: 'getPresentation' },
  { op: 'getSlide', key: { index } },
];

function read(host: Host, index: number) {
  const r = host.execute({
    operations: [...slide(index), { op: 'readHeaderFooter', slide: { result: 1 } }],
  });
  if (!r.ok) throw new Error(JSON.stringify(r.results));
  const value = (r.results[2] as { value: { headerFooter: unknown } }).value;
  return value.headerFooter as Record<string, unknown>;
}

function write(host: Host, settings: unknown, index?: number) {
  const operations: AutomationOperation[] =
    index === undefined
      ? [{ op: 'setHeaderFooter', settings } as AutomationOperation]
      : [
          ...slide(index),
          { op: 'setHeaderFooter', slides: [{ result: 1 }], settings } as AutomationOperation,
        ];
  return host.execute({ operations });
}

const slideXml = (store: PptxDocumentStore, index: number): string => {
  const pkg = store.package;
  return serializeXmlPart(xmlPartOf(pkg, presentationInfo(pkg).slides[index]!.partName)!);
};

describe('readHeaderFooter', () => {
  test('reads the slide number, footer, and date of one slide', () => {
    const { host } = setup(
      DATE_PH(
        '<a:fld id="{0C7D2C8B-1A6B-4B7E-9C3D-5E6F7A8B9C0D}" type="datetime4"><a:t>May 1, 2026</a:t></a:fld>'
      )
    );
    expect(read(host, 0)).toEqual({
      slideNumberVisible: true,
      footerVisible: true,
      footerText: 'Footer text & more',
      dateVisible: true,
      dateText: 'May 1, 2026',
      dateFixed: false,
    });
    expect(read(host, 1)).toMatchObject({ slideNumberVisible: false, footerVisible: false });
  });

  test('a footer placeholder with no text reads as not visible', () => {
    const { host } = setup('', DATE_PH('<a:endParaRPr lang="en-US"/>'));
    expect(read(host, 1)).toMatchObject({ dateVisible: false, dateText: '', dateFixed: false });
  });
});

describe('setHeaderFooter: automatic date', () => {
  test('a date field keeps its own type, and fixed text becomes today as datetime1', () => {
    const { host, store } = setup(
      DATE_PH(
        '<a:fld id="{0C7D2C8B-1A6B-4B7E-9C3D-5E6F7A8B9C0D}" type="datetime4"><a:t>May 1, 2026</a:t></a:fld>'
      ),
      DATE_PH('<a:r><a:rPr lang="en-US" b="1"/><a:t>Draft</a:t></a:r>')
    );
    expect(write(host, { dateFixed: false }).ok).toBe(true);
    expect(read(host, 0)).toMatchObject({ dateText: 'May 1, 2026', dateFixed: false });
    expect(slideXml(store, 0)).toContain('type="datetime4"');
    expect(read(host, 1)).toMatchObject({ dateText: shortDate(), dateFixed: false });
    // The new field keeps the character format of the fixed text.
    expect(slideXml(store, 1)).toMatch(/type="datetime1"><a:rPr lang="en-US" b="1"\/>/);
  });

  test('a slide with no date takes a field, not the fixed text of another slide', () => {
    const { host } = setup(DATE_PH('<a:r><a:t>Draft</a:t></a:r>'));
    expect(write(host, { dateFixed: false }, 1).ok).toBe(true);
    expect(read(host, 1)).toMatchObject({ dateText: shortDate(), dateFixed: false });
    expect(read(host, 0)).toMatchObject({ dateText: 'Draft', dateFixed: true });
  });

  test('a read written back to all slides keeps the field type of a date field', () => {
    const { host, store } = setup(
      DATE_PH(
        '<a:fld id="{0C7D2C8B-1A6B-4B7E-9C3D-5E6F7A8B9C0D}" type="datetime4"><a:t>May 1, 2026</a:t></a:fld>'
      )
    );
    expect(write(host, read(host, 0)).ok).toBe(true);
    expect(slideXml(store, 0)).toContain('type="datetime4"');
    expect(read(host, 0)).toEqual(read(host, 1));
  });

  test('dateFixed true turns a field into its text', () => {
    const { host } = setup(
      DATE_PH(
        '<a:fld id="{0C7D2C8B-1A6B-4B7E-9C3D-5E6F7A8B9C0D}" type="datetime4"><a:t>May 1, 2026</a:t></a:fld>'
      )
    );
    expect(write(host, { dateFixed: true }, 0).ok).toBe(true);
    expect(read(host, 0)).toMatchObject({ dateText: 'May 1, 2026', dateFixed: true });
  });
});

describe('setHeaderFooter: refusals', () => {
  test('arguments that do not go together refuse and name the argument', () => {
    const { host, store } = setup();
    const before = slideXml(store, 0);
    for (const [settings, detail] of [
      [{}, 'settings'],
      [{ other: true }, 'settings'],
      ['all', 'settings'],
      [{ footerText: 7 }, 'footerText'],
      [{ footerVisible: false, footerText: 'X' }, 'footerText'],
      [{ dateText: '', slideNumberVisible: true }, 'dateText'],
      [{ dateVisible: false, dateText: 'X' }, 'dateText'],
      [{ hideOnTitleSlides: 1, slideNumberVisible: true }, 'hideOnTitleSlides'],
    ] as const) {
      const r = write(host, settings);
      expect(r.results[0]).toMatchObject({ error: { code: 'invalid-argument', detail } });
    }
    const empty = host.execute({
      operations: [{ op: 'setHeaderFooter', slides: [], settings: { slideNumberVisible: true } }],
    });
    expect(empty.results[0]).toMatchObject({
      error: { code: 'invalid-argument', detail: 'slides' },
    });
    expect(slideXml(store, 0)).toBe(before);
  });
});

test('a footer keeps the unknown content of its placeholder through a write, save, and reopen', () => {
  const ext =
    '<p:extLst><p:ext uri="{B1E1E1E1-0000-4000-8000-000000000001}"><x:keep xmlns:x="urn:example:keep"/></p:ext></p:extLst>';
  const footer =
    '<p:sp><p:nvSpPr><p:cNvPr id="7" name="Footer 6"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>' +
    `<p:nvPr><p:ph type="ftr" sz="quarter" idx="11"/>${ext}</p:nvPr></p:nvSpPr><p:spPr/>` +
    '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>Old</a:t></a:r></a:p></p:txBody></p:sp>';
  const { host, store } = setup('', footer);
  expect(write(host, { footerText: 'New' }, 1).ok).toBe(true);
  const again = readPptx(writePptx(store.package));
  if (!again.ok) throw new Error('reopen');
  const xml = serializeXmlPart(
    xmlPartOf(again.pkg, presentationInfo(again.pkg).slides[1]!.partName)!
  );
  expect(xml).toContain('urn:example:keep');
  expect(xml).toContain('<a:t>New</a:t>');
  expect(xml).not.toContain('Old');
});

test('a read written back keeps a footer of two paragraphs as it is', () => {
  const footer =
    '<p:sp><p:nvSpPr><p:cNvPr id="7" name="Footer 6"/><p:cNvSpPr/>' +
    '<p:nvPr><p:ph type="ftr" sz="quarter" idx="11"/></p:nvPr></p:nvSpPr><p:spPr/>' +
    '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Confidential</a:t></a:r></a:p>' +
    '<a:p><a:r><a:t>Acme</a:t></a:r></a:p></p:txBody></p:sp>';
  const { host, store } = setup('', footer);
  const before = slideXml(store, 1);
  const own = read(host, 1);
  expect(own).toMatchObject({ footerVisible: true, footerText: 'ConfidentialAcme' });
  expect(write(host, own, 1).ok).toBe(true);
  expect(slideXml(store, 1)).toBe(before);
});
