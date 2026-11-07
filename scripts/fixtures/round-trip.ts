// The round-trip fixtures: content that an edit must keep. `hyperlinks` has every link kind.
// `preserved-content` has content that the editor keeps but does not edit. `svg-tab-stops` has an
// SVG picture with its PNG fallback, and paragraphs with tab stops. `header-footer` has a
// slide-number placeholder on its layouts. The feature edits (`scripts/fidelity/feature-edits.ts`)
// edit these decks and check what survives.

import { buildPptx, para, run, shape, solid, type FixtureOptions } from './builder.ts';
import { workbook } from './slide-structure.ts';
import { clickEntrances } from './timing.ts';

const IN = 914_400;
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const R_HLINK = `${R}/hyperlink`;
const R_IMAGE = `${R}/image`;
const R_MEDIA = 'http://schemas.microsoft.com/office/2007/relationships/media';
const A14 = 'http://schemas.microsoft.com/office/drawing/2010/main';
const P14 = 'http://schemas.microsoft.com/office/powerpoint/2010/main';
const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
const MATH = 'http://schemas.openxmlformats.org/officeDocument/2006/math';

const text = (id: number, name: string, y: number, body: string, h = 0.9): string =>
  shape({
    id,
    name,
    x: 0.7 * IN,
    y: y * IN,
    cx: 11 * IN,
    cy: h * IN,
    text: body,
    bodyPr: '<a:bodyPr wrap="square"><a:noAutofit/></a:bodyPr>',
  });

const linked = (label: string, rId: string, attrs = ''): string =>
  run(label, 'sz="2000"', `<a:hlinkClick r:id="${rId}"${attrs}/>`);

const title = (value: string): string => text(2, 'Title', 0.4, para(run(value, 'sz="3200"')));

// Slides are related to the presentation as rId101 … (see `buildPptx`).
const CUSTOM_SHOWS =
  '<p:custShowLst><p:custShow name="Short" id="0"><p:sldLst><p:sld r:id="rId101"/><p:sld r:id="rId103"/></p:sldLst></p:custShow></p:custShowLst>';

/** Every link kind of ECMA-376 Part 1, 21.1.2.3.5 hlinkClick, with the targets on slides 2 and 3. */
export function hyperlinksDeck(): FixtureOptions {
  const links =
    para(run('Plain text to link. ', 'sz="2000"') + run('More plain text.', 'sz="2000"')) +
    para(linked('A web page', 'rIdWeb', ' tooltip="Example &amp; more"')) +
    para(linked('An email with a subject', 'rIdMail')) +
    para(linked('Slide 3', 'rIdSlide', ' action="ppaction://hlinksldjump"')) +
    para(
      linked('Next', '', ' action="ppaction://hlinkshowjump?jump=nextslide"') +
        run(' · ', 'sz="2000"') +
        linked('Last', '', ' action="ppaction://hlinkshowjump?jump=lastslide"')
    ) +
    para(linked('A custom show', '', ' action="ppaction://customshow?id=0&amp;return=true"')) +
    para(linked('Another file', 'rIdFile', ' action="ppaction://hlinkfile"')) +
    para(linked('A link to remove', 'rIdOld')) +
    para(linked('A program', 'rIdProgram', ' action="ppaction://program"'));
  const shapeLink =
    '<p:sp><p:nvSpPr><p:cNvPr id="4" name="Linked shape"><a:hlinkClick r:id="rIdShape"/></p:cNvPr><p:cNvSpPr/><p:nvPr/></p:nvSpPr>' +
    `<p:spPr><a:xfrm><a:off x="${9 * IN}" y="${5.6 * IN}"/><a:ext cx="${2.7 * IN}" cy="${0.9 * IN}"/></a:xfrm><a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>${solid('4472C4')}</p:spPr>` +
    `<p:txBody><a:bodyPr anchor="ctr"/><a:lstStyle/>${para(run('Shape link', 'sz="1800"'), '<a:pPr algn="ctr"/>')}</p:txBody></p:sp>`;
  return {
    slides: [
      {
        shapes: title('Hyperlinks') + text(3, 'Links', 1.4, links, 5) + shapeLink,
        rels: [
          ['rIdWeb', R_HLINK, 'https://example.com/docs?q=1&amp;r=2', 'External'],
          ['rIdMail', R_HLINK, 'mailto:team@example.com?subject=Hello%20there', 'External'],
          ['rIdSlide', `${R}/slide`, 'slide3.xml'],
          ['rIdFile', R_HLINK, 'notes.txt', 'External'],
          ['rIdOld', R_HLINK, 'https://example.org/old', 'External'],
          ['rIdShape', R_HLINK, 'https://example.net/', 'External'],
          ['rIdProgram', R_HLINK, 'tool.exe', 'External'],
        ],
      },
      { shapes: title('Slide 2: next slide target') },
      { shapes: title('Slide 3: slide link target') },
    ],
    presentationLists: CUSTOM_SHOWS,
  };
}

/** A minimal EMF picture (MS-EMF 2.3): a black rectangle, 2 × 1 in. */
export function emfPicture(): Uint8Array {
  const records: number[][] = [];
  const header = [1, 88, 0, 0, 191, 95, 0, 0, 5080, 2540, 0x464d4520, 0x10000, 0, 0, 1, 0, 0, 0, 1920, 1080, 508, 286];
  // Handles (u16) and reserved (u16) share one u32 slot: 1 handle.
  records.push(header);
  records.push([37, 12, 0x80000004]); // EMR_SELECTOBJECT BLACK_BRUSH
  records.push([43, 24, 10, 10, 181, 85]); // EMR_RECTANGLE
  records.push([14, 20, 0, 16, 20]); // EMR_EOF
  const words = records.flat();
  const bytes = new Uint8Array(words.length * 4);
  const view = new DataView(bytes.buffer);
  words.forEach((w, i) => view.setInt32(i * 4, w | 0, true));
  view.setUint32(48, bytes.length, true); // nBytes
  view.setUint32(52, records.length, true); // nRecords
  view.setUint32(56, 1, true); // nHandles = 1, reserved = 0
  return bytes;
}

/** A minimal WMF picture (MS-WMF 2.3): a rectangle with the default pen and brush, 2 × 1 in. */
export function wmfPicture(): Uint8Array {
  // Records: size in 16-bit words (u32), function (u16), parameters (i16).
  const record = (fn: number, ...params: number[]): number[] => {
    const size = 3 + params.length;
    return [size & 0xffff, size >>> 16, fn, ...params];
  };
  const body = [
    ...record(0x020b, 0, 0), // META_SETWINDOWORG y, x
    ...record(0x020c, 1440, 2880), // META_SETWINDOWEXT y, x
    ...record(0x041b, 1400, 2840, 40, 40), // META_RECTANGLE bottom, right, top, left
    ...record(0x0000), // META_EOF
  ];
  const size = 9 + body.length;
  const header = [1, 9, 0x0300, size & 0xffff, size >>> 16, 0, 7, 0, 0];
  // The placeable header: key, handle, bounds, units per inch, reserved, checksum.
  const placeable = [0xcdd7, 0x9ac6, 0, 0, 0, 2880, 1440, 1440, 0, 0];
  placeable.push(placeable.reduce((a, b) => a ^ b, 0));
  const words = [...placeable, ...header, ...body];
  const bytes = new Uint8Array(words.length * 2);
  const view = new DataView(bytes.buffer);
  words.forEach((w, i) => view.setUint16(i * 2, w & 0xffff, true));
  return bytes;
}

/** The smallest MP4 box: a file type box with no media. A video that does not play. */
export function mp4Stub(): Uint8Array {
  const bytes = new Uint8Array(20);
  const view = new DataView(bytes.buffer);
  view.setUint32(0, 20);
  [...'ftypisom'].forEach((c, i) => view.setUint8(4 + i, c.charCodeAt(0)));
  view.setUint32(12, 0x200);
  [...'isom'].forEach((c, i) => view.setUint8(16 + i, c.charCodeAt(0)));
  return bytes;
}

/** A short silent WAV clip: 0.1 s of 8 kHz, 8-bit mono. */
export function wavClip(): Uint8Array {
  const samples = 800;
  const bytes = new Uint8Array(44 + samples).fill(128, 44);
  const view = new DataView(bytes.buffer);
  const ascii = (at: number, s: string): void => [...s].forEach((c, i) => view.setUint8(at + i, c.charCodeAt(0)));
  ascii(0, 'RIFF');
  view.setUint32(4, 36 + samples, true);
  ascii(8, 'WAVEfmt ');
  view.setUint32(16, 16, true);
  view.setUint16(20, 1, true);
  view.setUint16(22, 1, true);
  view.setUint32(24, 8000, true);
  view.setUint32(28, 8000, true);
  view.setUint16(32, 1, true);
  view.setUint16(34, 8, true);
  ascii(36, 'data');
  view.setUint32(40, samples, true);
  return bytes;
}

const xfrm = (x: number, y: number, cx: number, cy: number): string =>
  `<a:xfrm><a:off x="${Math.round(x * IN)}" y="${Math.round(y * IN)}"/><a:ext cx="${Math.round(cx * IN)}" cy="${Math.round(cy * IN)}"/></a:xfrm>`;

/** An extension that no reader knows, which a save must keep as it is. */
const UNKNOWN_EXT =
  '<a:extLst><a:ext uri="{7D0C4A51-2F4B-4E63-9A0E-5B8C1D2E3F40}"><x:note xmlns:x="urn:example:unknown" kept="yes">Unknown extension</x:note></a:ext></a:extLst>';

/**
 * Content that the editor keeps but does not edit: warped text, 3-D, an equation, EMF and WMF
 * pictures, audio and video clips, an embedded workbook, animations, a transition, and
 * extensions that no reader knows.
 */
export function preservedContentDeck(image: Uint8Array): FixtureOptions {
  const warp = shape({
    id: 3,
    name: 'Warped text',
    x: 0.7 * IN,
    y: 1.3 * IN,
    cx: 5 * IN,
    cy: 1.6 * IN,
    bodyPr: '<a:bodyPr wrap="none"><a:prstTxWarp prst="textArchUp"><a:avLst/></a:prstTxWarp><a:noAutofit/></a:bodyPr>',
    text: para(run('Warped text', 'sz="4000" b="1"'), '<a:pPr algn="ctr"/>'),
  });
  const threeD =
    '<p:sp><p:nvSpPr><p:cNvPr id="4" name="3-D shape"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>' +
    `<p:spPr>${xfrm(6.5, 1.3, 2.4, 1.6)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>${solid('ED7D31')}` +
    '<a:scene3d><a:camera prst="isometricOffAxis1Right"/><a:lightRig rig="threePt" dir="t"/></a:scene3d>' +
    '<a:sp3d extrusionH="228600" prstMaterial="plastic"><a:bevelT w="101600" h="38100"/></a:sp3d></p:spPr></p:sp>';
  // Normal text (`m:nor`) draws with the run font, an open font, instead of a math font.
  const mathRun = (t: string, italic = true): string =>
    `<m:r><m:rPr><m:nor/></m:rPr><a:rPr lang="en-US"${italic ? ' i="1"' : ''}><a:latin typeface="Liberation Serif"/></a:rPr><m:t>${t}</m:t></m:r>`;
  const equation =
    `<mc:AlternateContent xmlns:mc="${MC}"><mc:Choice xmlns:a14="${A14}" Requires="a14">` +
    '<p:sp><p:nvSpPr><p:cNvPr id="5" name="Equation"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>' +
    `<p:spPr>${xfrm(0.7, 3.3, 4, 0.9)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>` +
    '<p:txBody><a:bodyPr wrap="square"><a:spAutoFit/></a:bodyPr><a:lstStyle/><a:p>' +
    `<a14:m><m:oMathPara xmlns:m="${MATH}"><m:oMathParaPr><m:jc m:val="centerGroup"/></m:oMathParaPr><m:oMath>` +
    `${mathRun('E')}${mathRun('=', false)}${mathRun('m')}<m:sSup><m:e>${mathRun('c')}</m:e><m:sup>${mathRun('2', false)}</m:sup></m:sSup>` +
    '</m:oMath></m:oMathPara></a14:m><a:endParaRPr lang="en-US" sz="2800"/></a:p></p:txBody></p:sp></mc:Choice>' +
    '<mc:Fallback><p:sp><p:nvSpPr><p:cNvPr id="5" name="Equation"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>' +
    `<p:spPr>${xfrm(0.7, 3.3, 4, 0.9)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>` +
    '<a:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></a:blipFill></p:spPr>' +
    '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"><a:noFill/></a:rPr><a:t> </a:t></a:r></a:p></p:txBody></p:sp>' +
    '</mc:Fallback></mc:AlternateContent>';
  const emf =
    '<p:pic><p:nvPicPr><p:cNvPr id="6" name="EMF picture"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>' +
    '<p:blipFill><a:blip r:embed="rIdEmf"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>' +
    `<p:spPr>${xfrm(5.2, 3.3, 2, 1)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`;
  const audio =
    '<p:pic><p:nvPicPr><p:cNvPr id="7" name="Audio clip"><a:hlinkClick r:id="" action="ppaction://media"/></p:cNvPr>' +
    '<p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr><a:audioFile r:link="rIdAudio"/>' +
    `<p:extLst><p:ext uri="{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}"><p14:media xmlns:p14="${P14}" r:embed="rIdMedia"/></p:ext></p:extLst></p:nvPr></p:nvPicPr>` +
    '<p:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>' +
    `<p:spPr>${xfrm(7.8, 3.3, 0.6, 0.45)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`;
  const wmf =
    '<p:pic><p:nvPicPr><p:cNvPr id="9" name="WMF picture"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>' +
    '<p:blipFill><a:blip r:embed="rIdWmf"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>' +
    `<p:spPr>${xfrm(8.7, 3.3, 2, 1)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`;
  const video =
    '<p:pic><p:nvPicPr><p:cNvPr id="10" name="Video clip"><a:hlinkClick r:id="" action="ppaction://media"/></p:cNvPr>' +
    '<p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr><a:videoFile r:link="rIdVideo"/>' +
    `<p:extLst><p:ext uri="{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}"><p14:media xmlns:p14="${P14}" r:embed="rIdVideoMedia"/></p:ext></p:extLst></p:nvPr></p:nvPicPr>` +
    '<p:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>' +
    `<p:spPr>${xfrm(9.3, 1.3, 1.6, 1.2)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`;
  const ole =
    '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="11" name="Embedded workbook"/><p:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>' +
    `<p:xfrm><a:off x="${Math.round(11.1 * IN)}" y="${Math.round(1.3 * IN)}"/><a:ext cx="${Math.round(1.6 * IN)}" cy="${Math.round(1.2 * IN)}"/></p:xfrm>` +
    '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/presentationml/2006/ole">' +
    `<p:oleObj name="Worksheet" r:id="rIdOle" imgW="${Math.round(1.6 * IN)}" imgH="${Math.round(1.2 * IN)}" progId="Package"><p:embed/>` +
    '<p:pic><p:nvPicPr><p:cNvPr id="0" name=""/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>' +
    `<p:spPr>${xfrm(11.1, 1.3, 1.6, 1.2)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic></p:oleObj></a:graphicData></a:graphic></p:graphicFrame>`;
  const animated =
    '<p:sp><p:nvSpPr><p:cNvPr id="8" name="Animated text">' + UNKNOWN_EXT + '</p:cNvPr><p:cNvSpPr/><p:nvPr/></p:nvSpPr>' +
    `<p:spPr>${xfrm(0.7, 4.9, 11, 1.2)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>${solid('DEEBF7')}</p:spPr>` +
    '<p:txBody><a:bodyPr wrap="square"><a:noAutofit/></a:bodyPr><a:lstStyle/>' +
    para(run('This shape flies in on a click. Edits keep the animation.', 'sz="2000"')) +
    '</p:txBody></p:sp>';
  const slideExt =
    '<p:transition spd="med"><p:fade/></p:transition>' +
    clickEntrances([8, 4]) +
    `<p:extLst><p:ext uri="{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}"><p14:creationId xmlns:p14="${P14}" val="2000001"/></p:ext>` +
    '<p:ext uri="{3A86A75C-4F4B-4683-9AE1-C65F6400EC91}"><x:slideNote xmlns:x="urn:example:unknown">Unknown slide extension</x:slideNote></p:ext></p:extLst>';
  return {
    slides: [
      {
        shapes:
          title('Kept content') + warp + threeD + equation + emf + audio + wmf + video + ole + animated,
        rels: [
          ['rIdImg', R_IMAGE, '../media/image1.png'],
          ['rIdEmf', R_IMAGE, '../media/image2.emf'],
          ['rIdAudio', `${R}/audio`, '../media/media1.wav'],
          ['rIdMedia', R_MEDIA, '../media/media1.wav'],
          ['rIdWmf', R_IMAGE, '../media/image3.wmf'],
          ['rIdVideo', `${R}/video`, '../media/media2.mp4'],
          ['rIdVideoMedia', R_MEDIA, '../media/media2.mp4'],
          ['rIdOle', `${R}/package`, '../embeddings/Embedded_Worksheet1.xlsx'],
        ],
        ext: slideExt,
      },
    ],
    media: [
      { name: 'image1.png', bytes: image, contentType: 'image/png' },
      { name: 'image2.emf', bytes: emfPicture(), contentType: 'image/x-emf' },
      { name: 'media1.wav', bytes: wavClip(), contentType: 'audio/wav' },
      { name: 'image3.wmf', bytes: wmfPicture(), contentType: 'image/x-wmf' },
      { name: 'media2.mp4', bytes: mp4Stub(), contentType: 'video/mp4' },
    ],
    binaryParts: [
      {
        name: 'ppt/embeddings/Embedded_Worksheet1.xlsx',
        bytes: workbook(),
        contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      },
    ],
  };
}

/** A small SVG picture: a blue circle on a light square. */
const SVG = new TextEncoder().encode(
  '<svg xmlns="http://www.w3.org/2000/svg" width="96" height="96" viewBox="0 0 96 96">' +
    '<rect width="96" height="96" fill="#DEEBF7"/><circle cx="48" cy="48" r="36" fill="#2E75B6"/></svg>'
);

/** An SVG picture with its PNG fallback, a PNG picture, and paragraphs with tab stops. */
export function svgTabStopsDeck(image: Uint8Array): FixtureOptions {
  const svgPicture =
    '<p:pic><p:nvPicPr><p:cNvPr id="3" name="SVG picture" descr="A blue circle"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>' +
    '<p:blipFill><a:blip r:embed="rIdPng"><a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}">' +
    '<asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="rIdSvg"/></a:ext></a:extLst></a:blip>' +
    '<a:stretch><a:fillRect/></a:stretch></p:blipFill>' +
    `<p:spPr>${xfrm(0.7, 0.5, 2.5, 2.5)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`;
  const pngPicture =
    '<p:pic><p:nvPicPr><p:cNvPr id="4" name="PNG picture"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>' +
    '<p:blipFill><a:blip r:embed="rIdPng"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>' +
    `<p:spPr>${xfrm(3.6, 0.5, 3.3, 2.5)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`;
  const tabs =
    '<a:tabLst><a:tab pos="2743200" algn="l"/><a:tab pos="5486400" algn="ctr"/><a:tab pos="8229600" algn="r"/><a:tab pos="9601200" algn="dec"/></a:tabLst>';
  const row = (cells: string): string => para(run(cells, 'sz="2000"'), `<a:pPr marL="0" indent="0">${tabs}</a:pPr>`);
  const tabbed = shape({
    id: 5,
    name: 'Tab stops',
    x: 0.7 * IN,
    y: 3.4 * IN,
    cx: 11 * IN,
    cy: 2.6 * IN,
    fill: solid('F2F2F2'),
    bodyPr: '<a:bodyPr wrap="square"><a:noAutofit/></a:bodyPr>',
    text:
      row('Item\tLeft\tCentered\tRight\t1.50') +
      row('Second\tA\tB\tC\t12.25') +
      row('Third\tAlpha\tBeta\tGamma\t100.5') +
      para(run('An indented paragraph with a hanging first line.', 'sz="2000"'), '<a:pPr marL="457200" indent="-228600"/>'),
  });
  return {
    slides: [
      {
        shapes: svgPicture + pngPicture + tabbed,
        rels: [
          ['rIdPng', R_IMAGE, '../media/image1.png'],
          ['rIdSvg', R_IMAGE, '../media/image2.svg'],
        ],
      },
    ],
    media: [
      { name: 'image1.png', bytes: image, contentType: 'image/png' },
      { name: 'image2.svg', bytes: SVG, contentType: 'image/svg+xml' },
    ],
  };
}

/**
 * Layouts with a slide-number placeholder. Slide 1 shows its number and a footer with its own
 * box; slide 2 shows neither.
 */
export function headerFooterDeck(): FixtureOptions {
  const footer =
    '<p:sp><p:nvSpPr><p:cNvPr id="4" name="Footer Placeholder 3"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="ftr" sz="quarter" idx="11"/></p:nvPr></p:nvSpPr>' +
    `<p:spPr>${xfrm(4.2, 6.9, 4.9, 0.4)}</p:spPr><p:txBody><a:bodyPr/><a:lstStyle/>${para(run('Footer text & more', 'sz="1200"'), '<a:pPr algn="ctr"/>')}</p:txBody></p:sp>`;
  const number =
    '<p:sp><p:nvSpPr><p:cNvPr id="5" name="Slide Number Placeholder 4"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldNum" sz="quarter" idx="12"/></p:nvPr></p:nvSpPr><p:spPr/>' +
    '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:fld id="{B6F15528-21DE-4FAA-801E-634DDDAF4B2B}" type="slidenum"><a:rPr lang="en-US"/><a:t>1</a:t></a:fld><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>';
  return {
    footers: true,
    slides: [
      { shapes: title('Slide 1: a footer and a number') + footer + number },
      { shapes: title('Slide 2: no number yet') },
    ],
  };
}

/** Build the round-trip decks, for tests. */
export function buildRoundTripDecks(image: Uint8Array): Record<string, Uint8Array> {
  return {
    'header-footer': buildPptx(headerFooterDeck()),
    hyperlinks: buildPptx(hyperlinksDeck()),
    'preserved-content': buildPptx(preservedContentDeck(image)),
    'svg-tab-stops': buildPptx(svgTabStopsDeck(image)),
  };
}
