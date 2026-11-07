// Synthetic decks at the sizes that the editor UI reads on every change: the toolbar state,
// the slide panel, the layout menu, the selection, and the live region. Each one must open
// and edit in the browser without a long freeze.

import { strToU8, zlibSync } from 'fflate';
import { buildPptx, para, run, shape, solid, table } from '../fixtures/builder.ts';
import {
  IN,
  R_NS,
  box,
  one,
  patched,
  placeholder,
  range,
  SLIDE,
  tableRows,
  type HangDeck,
} from './deck-parts.ts';

/** A solid `w`×`h` PNG that decodes. */
function solidPng(w: number, h: number, rgb: readonly [number, number, number]): Uint8Array {
  const crcTable = range(256).map((n) => {
    let c = n;
    for (let k = 0; k < 8; k += 1) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    return c >>> 0;
  });
  const crc = (b: Uint8Array) => {
    let c = 0xffffffff;
    for (const x of b) c = crcTable[(c ^ x) & 0xff]! ^ (c >>> 8);
    return (c ^ 0xffffffff) >>> 0;
  };
  const chunk = (type: string, data: Uint8Array) => {
    const out = new Uint8Array(12 + data.length);
    const v = new DataView(out.buffer);
    v.setUint32(0, data.length);
    out.set(strToU8(type), 4);
    out.set(data, 8);
    v.setUint32(8 + data.length, crc(out.subarray(4, 8 + data.length)));
    return out;
  };
  const ihdr = new Uint8Array(13);
  new DataView(ihdr.buffer).setUint32(0, w);
  new DataView(ihdr.buffer).setUint32(4, h);
  ihdr.set([8, 2, 0, 0, 0], 8);
  const raw = new Uint8Array((w * 3 + 1) * h);
  for (let y = 0; y < h; y += 1)
    for (let x = 0; x < w; x += 1) raw.set(rgb, y * (w * 3 + 1) + 1 + x * 3);
  const parts = [
    new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    chunk('IHDR', ihdr),
    chunk('IDAT', zlibSync(raw)),
    chunk('IEND', new Uint8Array()),
  ];
  const png = new Uint8Array(parts.reduce((a, p) => a + p.length, 0));
  let at = 0;
  for (const p of parts) (png.set(p, at), (at += p.length));
  return png;
}

const R_IMAGE = `${R_NS}/image`;
const R_HLINK = `${R_NS}/hyperlink`;

/** A small shape at grid cell `i` of a `cols`-wide grid. */
const cellBox = (id: number, i: number, cols: number, text: string, extra = '') =>
  shape({
    id,
    name: `Cell ${i}`,
    x: (i % cols) * Math.floor((12 * IN) / cols),
    y: Math.floor(i / cols) * Math.floor((7 * IN) / cols),
    cx: Math.floor((12 * IN) / cols),
    cy: Math.floor((7 * IN) / cols),
    text: para(run(text, '', extra)),
  });

export function uiScaleDecks(): HangDeck[] {
  const layouts = (n: number) =>
    range(n).map((i) => ({ name: `Layout ${i}`, shapes: placeholder(2, 'title', 0) }));
  return [
    {
      id: 'ui-layouts-3000',
      target: '3000 layouts under one master: the layout menu and the toolbar state',
      bytes: () =>
        buildPptx({
          slides: [{ shapes: placeholder(2, 'title', 0), layoutName: 'Layout 7' }],
          layouts: layouts(3000),
        }),
    },
    {
      id: 'ui-slides-2000-own-layout',
      target: '2000 slides, each on its own layout',
      bytes: () =>
        buildPptx({
          slides: range(2000).map((i) => ({
            shapes: placeholder(2, 'title', 0),
            layoutName: `Layout ${i}`,
          })),
          layouts: layouts(2000),
        }),
    },
    {
      id: 'ui-placeholders-matched-20000',
      target: '20000 slide placeholders that each match a layout placeholder',
      bytes: () =>
        patched(one(''), (f) => {
          const rels = f.get('ppt/slides/_rels/slide1.xml.rels');
          const layout = `ppt/slideLayouts/${/slideLayouts\/([^"]+)"/.exec(rels)![1]}`;
          const places = range(20_000)
            .map((i) => placeholder(5000 + i, 'body', 10 + i))
            .join('');
          f.set(layout, f.get(layout).replace('</p:spTree>', `${places}</p:spTree>`));
          f.set(SLIDE, f.get(SLIDE).replace('</p:spTree>', `${places}</p:spTree>`));
        }),
    },
    {
      id: 'ui-sections-3000',
      target: '3000 slides, each in its own section',
      bytes: () => {
        const sections = range(3000)
          .map(
            (i) =>
              `<p14:section name="Section ${i}" id="{00000000-0000-0000-0000-${String(i).padStart(12, '0')}}">` +
              `<p14:sldIdLst><p14:sldId id="${256 + i}"/></p14:sldIdLst></p14:section>`
          )
          .join('');
        return buildPptx({
          slides: range(3000).map((i) => ({ shapes: box(2, para(run(`Slide ${i}`))) })),
          presentationExt:
            '<p:extLst><p:ext uri="{521415D9-36F7-43E2-AB2F-B90AF26B5E84}">' +
            `<p14:sectionLst xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main">${sections}</p14:sectionLst></p:ext></p:extLst>`,
        });
      },
    },
    {
      id: 'ui-custom-shows-2000',
      target: '2000 custom shows that each list every slide',
      bytes: () => {
        const list = range(20)
          .map((i) => `<p:sld r:id="rId${101 + i}"/>`)
          .join('');
        const shows = range(2000)
          .map(
            (i) =>
              `<p:custShow name="Show ${i}" id="${i}"><p:sldLst>${list}</p:sldLst></p:custShow>`
          )
          .join('');
        return buildPptx({
          slides: range(20).map((i) => ({ shapes: box(2, para(run(`Slide ${i}`))) })),
          presentationLists: `<p:custShowLst>${shows}</p:custShowLst>`,
        });
      },
    },
    {
      id: 'ui-images-2000',
      target: '2000 pictures on one slide, each its own media part',
      bytes: () =>
        buildPptx({
          slides: [
            {
              shapes: range(2000)
                .map(
                  (i) =>
                    `<p:pic><p:nvPicPr><p:cNvPr id="${i + 2}" name="Pic ${i}" descr="Picture ${i}"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>` +
                    `<p:blipFill><a:blip r:embed="rIdI${i}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
                    `<p:spPr><a:xfrm><a:off x="${(i % 50) * 230_000}" y="${Math.floor(i / 50) * 160_000}"/><a:ext cx="220000" cy="150000"/></a:xfrm>` +
                    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>'
                )
                .join(''),
              rels: range(2000).map((i) => [`rIdI${i}`, R_IMAGE, `../media/i${i}.png`] as const),
            },
          ],
          media: range(2000).map((i) => ({
            name: `i${i}.png`,
            bytes: solidPng(8, 8, [i % 256, (i * 7) % 256, (i * 13) % 256]),
            contentType: 'image/png',
          })),
        }),
    },
    {
      id: 'ui-hyperlinks-20000',
      target: '20000 linked runs on one slide',
      bytes: () =>
        buildPptx({
          slides: [
            {
              shapes: box(
                2,
                range(200)
                  .map((p) =>
                    para(
                      range(100)
                        .map((r) =>
                          run(`link ${p}.${r} `, '', `<a:hlinkClick r:id="rIdL${p * 100 + r}"/>`)
                        )
                        .join('')
                    )
                  )
                  .join(''),
                { cy: 6 * IN }
              ),
              rels: range(20_000).map(
                (i) => [`rIdL${i}`, R_HLINK, `https://example.com/${i}`, 'External'] as const
              ),
            },
          ],
        }),
    },
    {
      id: 'ui-connectors-5000',
      target: '5000 connectors between 5000 shapes',
      bytes: () =>
        one(
          range(5000)
            .map((i) => cellBox(i + 2, i, 100, ''))
            .join('') +
            range(5000)
              .map(
                (i) =>
                  `<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="${10_000 + i}" name="C${i}"/><p:cNvCxnSpPr>` +
                  `<a:stCxn id="${i + 2}" idx="0"/><a:endCxn id="${((i + 1) % 5000) + 2}" idx="2"/></p:cNvCxnSpPr><p:nvPr/></p:nvCxnSpPr>` +
                  `<p:spPr><a:xfrm><a:off x="${(i % 100) * 10_000}" y="${Math.floor(i / 100) * 10_000}"/><a:ext cx="100000" cy="100000"/></a:xfrm>` +
                  `<a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom><a:ln w="6350">${solid('000000')}</a:ln></p:spPr></p:cxnSp>`
              )
              .join('')
        ),
    },
    {
      id: 'ui-animations-20000',
      target: '20000 animation steps on 2000 shapes',
      bytes: () =>
        patched(
          one(
            range(2000)
              .map((i) => cellBox(i + 2, i, 50, `s${i}`))
              .join('')
          ),
          (f) => {
            let id = 3;
            const step = (spid: number) =>
              `<p:par><p:cTn id="${id++}" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst>` +
              `<p:par><p:cTn id="${id++}" presetID="10" presetClass="entr" presetSubtype="0" fill="hold" nodeType="clickEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>` +
              `<p:set><p:cBhvr><p:cTn id="${id++}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>` +
              `<p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr>` +
              '<p:to><p:strVal val="visible"/></p:to></p:set></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>';
            const steps = range(20_000)
              .map((i) => step((i % 2000) + 2))
              .join('');
            const timing =
              `<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>` +
              `<p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>${steps}</p:childTnLst></p:cTn>` +
              '<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>' +
              '<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq>' +
              '</p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>';
            const xml = f.get(SLIDE);
            f.set(SLIDE, xml.replace('</p:sld>', `${timing}</p:sld>`));
          }
        ),
    },
    {
      id: 'ui-group-wide-20000',
      target: 'one group with 20000 shapes',
      bytes: () =>
        one(
          '<p:grpSp><p:nvGrpSpPr><p:cNvPr id="1000000" name="Wide"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
            `<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${12 * IN}" cy="${7 * IN}"/><a:chOff x="0" y="0"/><a:chExt cx="${12 * IN}" cy="${7 * IN}"/></a:xfrm></p:grpSpPr>` +
            range(20_000)
              .map((i) => cellBox(i + 2, i, 200, 'g'))
              .join('') +
            '</p:grpSp>'
        ),
    },
    {
      id: 'ui-tables-500',
      target: '500 tables of 5 by 5 cells on one slide',
      bytes: () =>
        one(
          range(500)
            .map((i) =>
              table({
                id: i + 2,
                x: (i % 25) * 450_000,
                y: Math.floor(i / 25) * 320_000,
                cx: 400_000,
                cy: 300_000,
                columns: [80_000, 80_000, 80_000, 80_000, 80_000],
                tblPr: '<a:tblPr/>',
                rows: tableRows(5, 5, 'ab'),
              })
            )
            .join('')
        ),
    },
    {
      id: 'ui-names-100k',
      target: 'a shape whose name and description are 100000 characters',
      bytes: () =>
        one(
          box(2, para(run('named'))).replace(
            /<p:cNvPr id="2" name="[^"]*"/,
            `<p:cNvPr id="2" name="${'N'.repeat(100_000)}" descr="${'D '.repeat(50_000)}"`
          )
        ),
    },
    {
      id: 'ui-fonts-5000-runs',
      target: '5000 runs, each in its own font family, in one text box',
      bytes: () =>
        one(
          box(
            2,
            range(50)
              .map((p) =>
                para(
                  range(100)
                    .map((r) =>
                      run(`f${p * 100 + r} `, '', `<a:latin typeface="Face ${p * 100 + r}"/>`)
                    )
                    .join('')
                )
              )
              .join(''),
            { cy: 6 * IN }
          )
        ),
    },
    {
      id: 'ui-notes-20000-paragraphs',
      target: 'speaker notes of 20000 paragraphs',
      bytes: () =>
        buildPptx({
          slides: [
            {
              shapes: box(2, para(run('notes'))),
              notes: range(20_000)
                .map((i) => `Note line ${i}.`)
                .join('\n'),
            },
          ],
        }),
    },
    {
      id: 'ui-slides-1000-heavy',
      target: '1000 slides with 50 text shapes each',
      bytes: () =>
        buildPptx({
          slides: range(1000).map((s) => ({
            shapes: range(50)
              .map((i) => cellBox(i + 2, i, 10, `S${s} shape ${i}`))
              .join(''),
          })),
        }),
    },
    {
      id: 'ui-text-boxes-same-word-5000',
      target: '5000 text boxes with the same word, for find and replace',
      bytes: () =>
        one(
          range(5000)
            .map((i) => cellBox(i + 2, i, 80, 'match match match'))
            .join('')
        ),
    },
  ];
}
