import { describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import {
  buildPptx,
  cell,
  para,
  run,
  shape,
  table,
} from '../../../../../scripts/fixtures/builder.ts';
import { readPptx } from '../../store/index.ts';
import { resolveSlide } from '../../cascade/index.ts';
import { APPROXIMATIONS } from '../approximations.ts';
import { createFixedMeasurer, layoutSlide } from '../index.ts';

const FIDELITY_DOC = join(import.meta.dir, '../../../../../docs/site/content/pptx-fidelity.mdx');

/** The bold titles of the items under "Known approximations" on the public fidelity page (docs/site/content/pptx-fidelity.mdx). */
function documentedTitles(): string[] {
  const doc = readFileSync(FIDELITY_DOC, 'utf8');
  const section = doc.slice(
    doc.indexOf('## Known approximations'),
    doc.indexOf('\n## ', doc.indexOf('## Known approximations') + 3)
  );
  return [...section.matchAll(/^- \*\*(.+?)\.\*\*/gm)].map((m) => m[1]);
}

function approximationsOf(shapes: string): string[] {
  const r = readPptx(buildPptx({ slides: [{ shapes, layout: 'blank' }] }));
  if (!r.ok) throw new Error('open failed');
  return layoutSlide(resolveSlide(r.pkg, 0)!, createFixedMeasurer())
    .diagnostics.filter((d) => d.startsWith('approximation: '))
    .map((d) => d.slice('approximation: '.length));
}

describe('the approximation catalog', () => {
  test('lists exactly the approximations of the public fidelity page, in order', () => {
    expect(APPROXIMATIONS.map((a) => a.title)).toEqual(documentedTitles());
  });

  test('ids are unique', () => {
    expect(new Set(APPROXIMATIONS.map((a) => a.id)).size).toBe(APPROXIMATIONS.length);
  });
});

describe('approximations a slide uses', () => {
  const IN = 914_400;

  test('a plain shape uses none', () => {
    expect(approximationsOf(shape({ id: 2, x: 0, y: 0, cx: IN, cy: IN }))).toEqual([]);
  });

  test('text, effects, and line ends are detected', () => {
    const found = approximationsOf(
      shape({
        id: 2,
        x: 0,
        y: 0,
        cx: 4 * IN,
        cy: IN,
        effects:
          '<a:effectLst><a:outerShdw blurRad="50800" kx="600000"><a:srgbClr val="000000"/></a:outerShdw>' +
          '<a:glow rad="1000"><a:srgbClr val="FF0000"/></a:glow></a:effectLst>',
        text:
          para(run('Small', 'cap="small" u="wavy" baseline="30000"')) +
          para(run('日本語')) +
          para(run('Centered'), '<a:pPr algn="ctr"><a:buChar char="•"/></a:pPr>') +
          para(run('Thai number'), '<a:pPr><a:buAutoNum type="thaiAlphaPeriod"/></a:pPr>') +
          para(
            run('Check'),
            '<a:pPr><a:buFont typeface="Wingdings"/><a:buChar char="ü"/></a:pPr>'
          ) +
          '<a:p><a:pPr rtl="1"/><a:r><a:rPr lang="he-IL"/><a:t>abc</a:t></a:r><a:br><a:rPr lang="he-IL"/></a:br></a:p>',
      }) +
        '<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="3" name="L"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr><p:spPr>' +
        '<a:xfrm><a:off x="0" y="2000000"/><a:ext cx="2000000" cy="0"/></a:xfrm><a:prstGeom prst="line"><a:avLst/></a:prstGeom>' +
        '<a:ln w="12700" cmpd="dbl"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:tailEnd type="triangle"/></a:ln></p:spPr></p:cxnSp>'
    );
    expect(found).toEqual([
      'line-model',
      'blur',
      'shadow-skew',
      'glow-soft-edge',
      'arrowhead-joins',
      'compound-lines',
      'small-caps',
      'underline-styles',
      'script-fonts',
      'rtl',
      'script-size',
      'aligned-bullets',
      'numbering-schemes',
      'symbol-bullets',
      'trailing-break',
    ]);
  });

  test('fill overlays, two-stop colors, outside foci, and unmeasured tiles are detected', () => {
    const stops =
      '<a:gsLst><a:gs pos="0"><a:srgbClr val="FFFFFF"/></a:gs><a:gs pos="100000"><a:srgbClr val="C00000"/></a:gs></a:gsLst>';
    const at = (fill: string, effects?: string): string =>
      shape({ id: 2, x: 0, y: 0, cx: 914_400, cy: 914_400, fill, ...(effects ? { effects } : {}) });
    expect(
      approximationsOf(
        at(
          '<a:solidFill><a:srgbClr val="808080"/></a:solidFill>',
          `<a:effectLst><a:fillOverlay blend="mult"><a:gradFill>${stops}<a:path path="rect"><a:fillToRect t="-80000" b="180000"/></a:path></a:gradFill></a:fillOverlay></a:effectLst>`
        )
      )
    ).toEqual(['fill-overlay', 'gradient-colors', 'path-gradient-focus', 'rect-path-gradient']);
    // The measured corner preset: a circle with a tile rectangle that doubles the box.
    expect(
      approximationsOf(
        at(
          `<a:gradFill>${stops}<a:path path="circle"><a:fillToRect r="100000" b="100000"/></a:path><a:tileRect r="-100000" b="-100000"/></a:gradFill>`
        )
      )
    ).toEqual(['gradient-colors']);
    expect(
      approximationsOf(
        at(`<a:gradFill>${stops}<a:lin ang="0" scaled="0"/><a:tileRect r="-100000"/></a:gradFill>`)
      )
    ).toEqual(['gradient-colors', 'gradient-tiles']);
  });

  test('group fills used by a child, table backgrounds, and dropped overlays are checked', () => {
    const two =
      '<a:gsLst><a:gs pos="0"><a:srgbClr val="FFFFFF"/></a:gs><a:gs pos="100000"><a:srgbClr val="000000"/></a:gs></a:gsLst>';
    const group =
      '<p:grpSp><p:nvGrpSpPr><p:cNvPr id="5" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
      '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/><a:chOff x="0" y="0"/><a:chExt cx="914400" cy="914400"/></a:xfrm>' +
      `<a:gradFill>${two}<a:lin ang="0"/></a:gradFill></p:grpSpPr>` +
      shape({ id: 6, x: 0, y: 0, cx: 914_400, cy: 914_400, fill: '<a:grpFill/>' }) +
      '</p:grpSp>';
    expect(approximationsOf(group)).toEqual(['gradient-colors']);
    const many = Array.from(
      { length: 70 },
      (_, i) =>
        `<a:path w="100" h="100"><a:moveTo><a:pt x="${i}" y="0"/></a:moveTo><a:lnTo><a:pt x="100" y="100"/></a:lnTo><a:close/></a:path>`
    ).join('');
    const overlaid =
      '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Many"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm>' +
      `<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/><a:pathLst>${many}</a:pathLst></a:custGeom>` +
      '<a:solidFill><a:srgbClr val="808080"/></a:solidFill><a:effectLst><a:fillOverlay blend="mult"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:fillOverlay></a:effectLst></p:spPr></p:sp>';
    // Paint drops the overlay, so the report says so and does not call it approximate.
    expect(approximationsOf(overlaid)).toEqual(['effects-not-drawn']);
  });

  test('table backgrounds and text fills get the fill flags', () => {
    const two =
      '<a:gsLst><a:gs pos="0"><a:srgbClr val="FFFFFF"/></a:gs><a:gs pos="100000"><a:srgbClr val="000000"/></a:gs></a:gsLst>';
    const rect = `<a:gradFill>${two}<a:path path="rect"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path></a:gradFill>`;
    const IN = 914_400;
    const framed = approximationsOf(
      table({
        id: 12,
        x: 0,
        y: 0,
        cx: IN,
        cy: IN,
        columns: [IN],
        tblPr: `<a:tblPr>${rect}</a:tblPr>`,
        rows: `<a:tr h="${IN}">${cell('y')}</a:tr>`,
      })
    );
    expect(framed).toContain('gradient-colors');
    expect(framed).toContain('rect-path-gradient');
    const bordered = approximationsOf(
      table({
        id: 13,
        x: 0,
        y: 0,
        cx: IN,
        cy: IN,
        columns: [IN],
        tblPr: '<a:tblPr/>',
        rows: `<a:tr h="${IN}">${cell('y', { tcPr: `<a:lnL w="12700"><a:gradFill>${two}<a:lin ang="0"/></a:gradFill></a:lnL>` })}</a:tr>`,
      })
    );
    expect(bordered).toContain('gradient-colors');
    const text = approximationsOf(
      shape({ id: 2, x: 0, y: 0, cx: IN, cy: IN, text: para(run('x', '', rect)) })
    );
    expect(text).toContain('gradient-text');
    expect(text).toContain('rect-path-gradient');
    // Lines and text runs with `a:grpFill` report the group's fill.
    const grouped = approximationsOf(
      '<p:grpSp><p:nvGrpSpPr><p:cNvPr id="5" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
        `<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${IN}" cy="${IN}"/><a:chOff x="0" y="0"/><a:chExt cx="${IN}" cy="${IN}"/></a:xfrm>${rect}</p:grpSpPr>` +
        shape({ id: 6, x: 0, y: 0, cx: IN, cy: IN, text: para(run('x', '', '<a:grpFill/>')) }) +
        '</p:grpSp>'
    );
    expect(grouped).toContain('rect-path-gradient');
    expect(grouped).toContain('gradient-text');
    const outlined = approximationsOf(
      '<p:grpSp><p:nvGrpSpPr><p:cNvPr id="5" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
        `<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${IN}" cy="${IN}"/><a:chOff x="0" y="0"/><a:chExt cx="${IN}" cy="${IN}"/></a:xfrm>${rect}</p:grpSpPr>` +
        shape({ id: 6, x: 0, y: 0, cx: IN, cy: IN, line: '<a:ln w="12700"><a:grpFill/></a:ln>' }) +
        '</p:grpSp>'
    );
    expect(outlined).toContain('rect-path-gradient');
  });

  test('a slide background is checked too', () => {
    const r = readPptx(
      buildPptx({
        slides: [
          {
            shapes: '',
            layout: 'blank',
            background:
              '<p:bg><p:bgPr><a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FFFFFF"/></a:gs><a:gs pos="100000"><a:srgbClr val="000000"/></a:gs></a:gsLst><a:lin ang="0"/></a:gradFill><a:effectLst/></p:bgPr></p:bg>',
          },
        ],
      })
    );
    if (!r.ok) throw new Error('open failed');
    const diagnostics = layoutSlide(resolveSlide(r.pkg, 0)!, createFixedMeasurer()).diagnostics;
    expect(diagnostics).toContain('approximation: gradient-colors');
  });

  test('text in a flipped group and table borders are detected', () => {
    const found = approximationsOf(
      '<p:grpSp><p:nvGrpSpPr><p:cNvPr id="10" name="G"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
        `<p:grpSpPr><a:xfrm flipH="1"><a:off x="0" y="0"/><a:ext cx="${IN}" cy="${IN}"/>` +
        `<a:chOff x="0" y="0"/><a:chExt cx="${IN}" cy="${IN}"/></a:xfrm></p:grpSpPr>` +
        shape({ id: 11, x: 0, y: 0, cx: IN, cy: IN, text: para(run('x')) }) +
        '</p:grpSp>' +
        table({
          id: 12,
          x: 0,
          y: 2 * IN,
          cx: IN,
          cy: IN,
          columns: [IN],
          tblPr: '<a:tblPr/>',
          rows: `<a:tr h="${IN}">${cell('y', { tcPr: '<a:lnL w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:lnL>' })}</a:tr>`,
        })
    );
    expect(found).toEqual(['line-model', 'flipped-text', 'coincident-borders']);
  });

  test('every text body uses the line model; a body without text does not', () => {
    const box = (text: string): string =>
      shape({ id: 2, x: 0, y: 0, cx: 4 * IN, cy: IN, bodyPr: '<a:bodyPr/>', text });
    expect(approximationsOf(box(para(run('x'))))).toEqual(['line-model']);
    expect(approximationsOf(box('<a:p><a:endParaRPr lang="en-US"/></a:p>'))).toEqual([]);
  });

  test('an empty text body reports no text approximation, whatever its body properties', () => {
    // An empty body lays out one line for the caret; it draws nothing, so it approximates nothing.
    const empty = shape({
      id: 2,
      x: 0,
      y: 0,
      cx: 4 * IN,
      cy: IN,
      flipH: true,
      bodyPr: '<a:bodyPr vert="eaVert" anchor="ctr" anchorCtr="1"/>',
      text: '<a:p><a:endParaRPr lang="en-US"/></a:p>',
    });
    expect(approximationsOf(empty)).toEqual([]);
  });

  test('every dashed picture border is flagged, round cap or not; a solid one is not', () => {
    const picture = (ln: string): string =>
      '<p:pic><p:nvPicPr><p:cNvPr id="2" name="P"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>' +
      '<p:blipFill><a:blip/><a:stretch><a:fillRect/></a:stretch></p:blipFill>' +
      `<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${IN}" cy="${IN}"/></a:xfrm>` +
      `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>${ln}</p:spPr></p:pic>`;
    const line = (dash: string, cap = ''): string =>
      `<a:ln w="38100"${cap}><a:solidFill><a:srgbClr val="000000"/></a:solidFill>` +
      `<a:prstDash val="${dash}"/></a:ln>`;
    expect(approximationsOf(picture(line('dash')))).toContain('picture-border-dots');
    expect(approximationsOf(picture(line('sysDot', ' cap="rnd"')))).toContain(
      'picture-border-dots'
    );
    expect(approximationsOf(picture(line('solid')))).not.toContain('picture-border-dots');
  });
});
