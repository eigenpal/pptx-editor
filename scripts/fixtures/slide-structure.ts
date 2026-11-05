// The `slide-structure` fixture: two sections (`p14:sectionLst`), two custom shows
// (`p:custShowLst`), legacy comments with a reply, a chart with its embedded workbook, an
// embedded OLE workbook, and pictures shared between slides. Slide add, duplicate, move, and delete are checked against it.

import { strToU8, zipSync } from 'fflate';
import { buildPptx, para, run, shape, type FixtureOptions } from './builder.ts';

const IN = 914_400;
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const CT_PML = 'application/vnd.openxmlformats-officedocument.presentationml';
const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';
const P = 'http://schemas.openxmlformats.org/presentationml/2006/main';

/** The section ids, as GUIDs (`p14:section/@id`). */
export const SECTION_IDS = ['{6E1B33A1-4C0F-4B53-9C2B-0D4A4B3E8A11}', '{0B7F5C2D-1E4A-4F3B-8C6D-2A9E7B1C5D22}'];

/** A minimal workbook (one sheet, two numbers), so the chart and the object have real data. */
export function workbook(): Uint8Array {
  const S = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const files: Record<string, Uint8Array> = {
    '[Content_Types].xml': strToU8(
      `${XML}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>' +
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
        '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>'
    ),
    '_rels/.rels': strToU8(
      `${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="xl/workbook.xml"/></Relationships>`
    ),
    'xl/workbook.xml': strToU8(
      `${XML}<workbook xmlns="${S}" xmlns:r="${R}"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>`
    ),
    'xl/_rels/workbook.xml.rels': strToU8(
      `${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`
    ),
    'xl/worksheets/sheet1.xml': strToU8(
      `${XML}<worksheet xmlns="${S}"><sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c></row></sheetData></worksheet>`
    ),
  };
  return zipSync(files, { mtime: new Date(Date.UTC(1980, 0, 2, 12)) });
}

const chart =
  `${XML}<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="${R}">` +
  '<c:chart><c:autoTitleDeleted val="1"/><c:plotArea><c:layout/><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>' +
  '<c:ser><c:idx val="0"/><c:order val="0"/><c:cat><c:strLit><c:ptCount val="2"/><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strLit></c:cat>' +
  '<c:val><c:numLit><c:formatCode>General</c:formatCode><c:ptCount val="2"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser>' +
  '<c:gapWidth val="150"/><c:axId val="10"/><c:axId val="20"/></c:barChart>' +
  '<c:catAx><c:axId val="10"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="20"/></c:catAx>' +
  '<c:valAx><c:axId val="20"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:crossAx val="10"/></c:valAx>' +
  '</c:plotArea><c:plotVisOnly val="1"/></c:chart><c:externalData r:id="rId1"><c:autoUpdate val="0"/></c:externalData></c:chartSpace>';

const rels = (items: readonly (readonly [string, string, string])[]): string =>
  `${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
  items.map(([id, type, target]) => `<Relationship Id="${id}" Type="${type}" Target="${target}"/>`).join('') +
  '</Relationships>';

const comments =
  `${XML}<p:cmLst xmlns:p="${P}">` +
  '<p:cm authorId="0" dt="2026-01-01T09:00:00.000" idx="1"><p:pos x="10" y="10"/><p:text>Check the chart numbers.</p:text></p:cm>' +
  '<p:cm authorId="0" dt="2026-01-01T09:05:00.000" idx="2"><p:pos x="10" y="10"/><p:text>Checked.</p:text>' +
  '<p:extLst><p:ext uri="{C676402C-5697-4E1C-873F-D02D1690AC5C}"><p15:threadingInfo xmlns:p15="http://schemas.microsoft.com/office/powerpoint/2012/main" timeZoneBias="0">' +
  '<p15:parentCm authorId="0" idx="1"/></p15:threadingInfo></p:ext></p:extLst></p:cm></p:cmLst>';

const authors =
  `${XML}<p:cmAuthorLst xmlns:p="${P}"><p:cmAuthor id="0" name="Fixture Author" initials="FA" lastIdx="2" clrIdx="0"/></p:cmAuthorLst>`;

const sections =
  '<p:extLst><p:ext uri="{521415D9-36F7-43E2-AB2F-B90AF26B5E84}"><p14:sectionLst xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main">' +
  `<p14:section name="Intro" id="${SECTION_IDS[0]}"><p14:sldIdLst><p14:sldId id="256"/><p14:sldId id="257"/></p14:sldIdLst></p14:section>` +
  `<p14:section name="Body" id="${SECTION_IDS[1]}"><p14:sldIdLst><p14:sldId id="258"/><p14:sldId id="259"/></p14:sldIdLst></p14:section>` +
  '</p14:sectionLst></p:ext></p:extLst>';

// Slides are related to the presentation as rId101 … rId104 (see `buildPptx`).
const customShows =
  '<p:custShowLst>' +
  '<p:custShow name="Short" id="0"><p:sldLst><p:sld r:id="rId101"/><p:sld r:id="rId102"/><p:sld r:id="rId104"/></p:sldLst></p:custShow>' +
  '<p:custShow name="Charts" id="1"><p:sldLst><p:sld r:id="rId101"/></p:sldLst></p:custShow>' +
  '</p:custShowLst>';

const creationId = (val: number): string =>
  '<p:extLst><p:ext uri="{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}">' +
  `<p14:creationId xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" val="${val}"/></p:ext></p:extLst>`;

const xfrm = (x: number, y: number, cx: number, cy: number, tag = 'a:xfrm'): string =>
  `<${tag}><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></${tag}>`;

const picture = (id: number, rId: string, x: number): string =>
  `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="Picture ${id - 1}"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>` +
  `<p:blipFill><a:blip r:embed="${rId}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
  `<p:spPr>${xfrm(x, 2 * IN, 3 * IN, 2.25 * IN)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`;

const title = (text: string): string =>
  shape({ id: 2, name: 'Title', x: 0.7 * IN, y: 0.5 * IN, cx: 11 * IN, cy: IN, text: para(run(text, 'sz="3200"')) });

/** The deck. `image` is the picture shared by slides 2 to 4; slide 2 also has its own. */
export function slideStructureDeck(image: Uint8Array, other: Uint8Array): FixtureOptions {
  return {
    slides: [
      {
        shapes:
          title('Intro 1: a chart and two comments') +
          '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="3" name="Chart 2"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>' +
          xfrm(IN, 2 * IN, 6 * IN, 4 * IN, 'p:xfrm') +
          '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">' +
          '<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rIdChart"/></a:graphicData></a:graphic></p:graphicFrame>',
        rels: [
          ['rIdChart', `${R}/chart`, '../charts/chart1.xml'],
          ['rIdCm', `${R}/comments`, '../comments/comment1.xml'],
        ],
        ext: creationId(1_000_001),
      },
      {
        shapes: title('Intro 2: a shared picture and its own picture') + picture(3, 'rIdImg', IN) + picture(4, 'rIdOwn', 5 * IN),
        rels: [
          ['rIdImg', `${R}/image`, '../media/image1.png'],
          ['rIdOwn', `${R}/image`, '../media/image2.png'],
        ],
        ext: creationId(1_000_002),
      },
      {
        shapes: title('Body 1: the shared picture') + picture(3, 'rIdImg', IN),
        rels: [['rIdImg', `${R}/image`, '../media/image1.png']],
        ext: creationId(1_000_003),
      },
      {
        shapes:
          title('Body 2: an embedded workbook') +
          '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="3" name="Object 2"/><p:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>' +
          xfrm(IN, 2 * IN, 3 * IN, 2.25 * IN, 'p:xfrm') +
          '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/presentationml/2006/ole">' +
          `<p:oleObj name="Worksheet" r:id="rIdOle" imgW="${3 * IN}" imgH="${2.25 * IN}" progId="Package"><p:embed/>` +
          '<p:pic><p:nvPicPr><p:cNvPr id="0" name=""/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>' +
          `<p:spPr>${xfrm(IN, 2 * IN, 3 * IN, 2.25 * IN)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic></p:oleObj></a:graphicData></a:graphic></p:graphicFrame>`,
        rels: [
          ['rIdOle', `${R}/package`, '../embeddings/Embedded_Worksheet2.xlsx'],
          ['rIdImg', `${R}/image`, '../media/image1.png'],
        ],
        ext: creationId(1_000_004),
      },
    ],
    media: [
      { name: 'image1.png', bytes: image, contentType: 'image/png' },
      { name: 'image2.png', bytes: other, contentType: 'image/png' },
    ],
    binaryParts: [
      {
        name: 'ppt/embeddings/Embedded_Worksheet1.xlsx',
        bytes: workbook(),
        contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      },
      {
        name: 'ppt/embeddings/Embedded_Worksheet2.xlsx',
        bytes: workbook(),
        contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      },
    ],
    extraParts: [
      ['ppt/charts/chart1.xml', 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml', chart],
      [
        'ppt/charts/_rels/chart1.xml.rels',
        'application/vnd.openxmlformats-package.relationships+xml',
        rels([['rId1', `${R}/package`, '../embeddings/Embedded_Worksheet1.xlsx']]),
      ],
      ['ppt/comments/comment1.xml', `${CT_PML}.comments+xml`, comments],
      ['ppt/commentAuthors.xml', `${CT_PML}.commentAuthors+xml`, authors],
    ],
    presentationRels: [['rIdAuthors', `${R}/commentAuthors`, 'commentAuthors.xml']],
    presentationLists: customShows,
    presentationExt: sections,
  };
}

/** Build the deck with stand-in pictures, for tests. */
export function buildSlideStructureDeck(): Uint8Array {
  // Two 1×1 PNGs that differ, so the shared and the own picture are distinct parts.
  const png = (rgb: string): Uint8Array =>
    Uint8Array.from(
      atob(
        rgb === 'red'
          ? 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8DwHwAFBQIAX8jx0gAAAABJRU5ErkJggg=='
          : 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=='
      ),
      (c) => c.charCodeAt(0)
    );
  return buildPptx(slideStructureDeck(png('red'), png('green')));
}
