// The Shape Format fixtures (G4): the XML the editor writes for the Shadow, Reflection, Glow,
// and Soft Edges galleries and for the Shape Styles gallery, theme effect styles on and off,
// and flips, turns, the aspect lock, and hidden shapes. The XML is written out here, not taken
// from the editor's own tables, so a wrong preset shows as a difference against the reference PDF.

import { para, run, shape, solid, type FixtureSlide } from './builder.ts';

const IN = 914_400;
const black = (alpha: number): string =>
  `<a:prstClr val="black"><a:alpha val="${alpha}"/></a:prstClr>`;
const outer = (dir: number, algn: string): string =>
  `<a:outerShdw blurRad="50800" dist="38100" dir="${dir}" algn="${algn}">${black(40000)}</a:outerShdw>`;
const inner = (dir: number): string =>
  `<a:innerShdw blurRad="63500" dist="50800" dir="${dir}">${black(50000)}</a:innerShdw>`;
const reflection = (endPos: number, dist: number): string =>
  `<a:reflection blurRad="6350" stA="50000" endA="300" endPos="${endPos}"${dist ? ` dist="${dist}"` : ''} dir="5400000" sy="-100000" algn="bl" rotWithShape="0"/>`;
const glow = (rad: number, accent: string): string =>
  `<a:glow rad="${rad}"><a:schemeClr val="${accent}"><a:satMod val="175000"/><a:alpha val="40000"/></a:schemeClr></a:glow>`;
const fx = (inside: string): string => `<a:effectLst>${inside}</a:effectLst>`;
const label = (text: string): string => para(run(text, 'sz="1200"'), '<a:pPr algn="ctr"/>');
const centered = '<a:bodyPr anchor="ctr"/>';

/** A shape in a grid cell: `col` and `row` from 0, 2.1 in apart. */
function cell(id: number, col: number, row: number, extra: Partial<Parameters<typeof shape>[0]>): string {
  return shape({
    id,
    x: (0.5 + col * 2.1) * IN,
    y: (0.6 + row * 1.6) * IN,
    cx: 1.6 * IN,
    cy: 1.0 * IN,
    bodyPr: centered,
    ...extra,
  });
}

/** `p:style` of a Shape Styles gallery cell, as the editor writes it. */
function styleXml(ln: [number, string], fill: [number, string], effect: number, font: string, c: string): string {
  const color = (v: string): string =>
    v === 'shade' ? `<a:schemeClr val="${c}"><a:shade val="15000"/></a:schemeClr>` : `<a:schemeClr val="${v === 'C' ? c : v}"/>`;
  return (
    `<p:style><a:lnRef idx="${ln[0]}">${color(ln[1])}</a:lnRef><a:fillRef idx="${fill[0]}">${color(fill[1])}</a:fillRef>` +
    `<a:effectRef idx="${effect}">${color('C')}</a:effectRef><a:fontRef idx="minor">${color(font)}</a:fontRef></p:style>`
  );
}

const ROWS: readonly [string, [number, string], [number, string], number, string][] = [
  ['Colored outline', [2, 'C'], [1, 'lt1'], 0, 'dk1'],
  ['Colored fill', [2, 'shade'], [1, 'C'], 0, 'lt1'],
  ['Light 1 outline', [3, 'lt1'], [1, 'C'], 1, 'lt1'],
  ['Subtle effect', [1, 'C'], [2, 'C'], 1, 'dk1'],
  ['Moderate effect', [1, 'C'], [3, 'C'], 2, 'lt1'],
  ['Intense effect', [0, 'C'], [3, 'C'], 3, 'lt1'],
];

export const shapeFormatFixtures: Record<string, FixtureSlide[]> = {
  'shape-effects-presets': [
    {
      // Outer and inner shadow presets.
      layout: 'blank',
      shapes: [
        ['Offset: bottom right', outer(2_700_000, 'tl')],
        ['Offset: bottom', outer(5_400_000, 't')],
        ['Offset: top left', outer(13_500_000, 'br')],
        ['Offset: left', outer(10_800_000, 'r')],
        ['Offset: center', `<a:outerShdw blurRad="63500" sx="102000" sy="102000" algn="ctr">${black(40000)}</a:outerShdw>`],
        ['Inside: top left', inner(13_500_000)],
        ['Inside: bottom', inner(5_400_000)],
        ['Inside: right', inner(0)],
        ['Inside: center', '<a:innerShdw blurRad="114300"><a:prstClr val="black"/></a:innerShdw>'],
      ]
        .map(([text, effect], i) =>
          cell(2 + i, i % 5, Math.floor(i / 5) * 2, { fill: solid('FFFFFF'), line: '<a:ln w="9525"><a:solidFill><a:srgbClr val="A5A5A5"/></a:solidFill></a:ln>', effects: fx(effect), text: label(text) }))
        .join(''),
    },
    {
      // Reflection, glow, and soft edge presets; glow and a shadow together, in schema order.
      layout: 'blank',
      shapes: [
        ['Tight reflection, touching', reflection(35_000, 0)],
        ['Half reflection, 4 pt', reflection(55_500, 50_800)],
        ['Full reflection, 8 pt', reflection(90_000, 101_600)],
        ['Glow 5 pt, Accent 1', glow(63_500, 'accent1')],
        ['Glow 18 pt, Accent 6', glow(228_600, 'accent6')],
        ['Soft edges 2.5 pt', '<a:softEdge rad="31750"/>'],
        ['Soft edges 10 pt', '<a:softEdge rad="127000"/>'],
        ['Glow and shadow', glow(101_600, 'accent2') + outer(2_700_000, 'tl')],
      ]
        .map(([text, effect], i) =>
          cell(2 + i, i % 4, Math.floor(i / 4) * 2, { fill: solid('5B9BD5'), effects: fx(effect), text: label(text) }))
        .join(''),
    },
    {
      // Theme effect styles: effectRef 1, 2, 3 (the theme's shadow), and 3 turned off by an
      // empty effect list (Shadow > No Shadow), and 3 with a glow added (the theme shadow kept).
      layout: 'blank',
      shapes:
        [1, 2, 3]
          .map((idx, i) =>
            cell(2 + i, i, 0, { style: styleXml([2, 'shade'], [1, 'C'], idx, 'lt1', 'accent1'), text: label(`effectRef ${idx}`) }))
          .join('') +
        cell(5, 0, 2, { style: styleXml([2, 'shade'], [1, 'C'], 3, 'lt1', 'accent1'), effects: '<a:effectLst/>', text: label('Shadow off') }) +
        cell(6, 1, 2, {
          style: styleXml([2, 'shade'], [1, 'C'], 3, 'lt1', 'accent1'),
          effects: fx(glow(63_500, 'accent2') + `<a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="63000"/></a:srgbClr></a:outerShdw>`),
          text: label('Theme shadow and glow'),
        }),
    },
  ],
  'shape-styles': [
    {
      // The six theme style rows in Accent 1, Accent 2, Accent 6, and Dark 1.
      layout: 'blank',
      shapes: ROWS.flatMap(([name, ln, fill, effect, font], row) =>
        ['accent1', 'accent2', 'accent6', 'dk1'].map((c, col) =>
          shape({
            id: 2 + row * 4 + col,
            name: `${name} ${c}`,
            x: (0.5 + col * 3) * IN,
            y: (0.3 + row * 1.12) * IN,
            cx: 2.4 * IN,
            cy: 0.8 * IN,
            prst: 'roundRect',
            bodyPr: centered,
            style: styleXml(ln, fill, effect, font, c),
            text: label(name),
          })
        )
      ).join(''),
    },
  ],
  'arrange-flip-hidden': [
    {
      // Flips and turns, with text (text is never mirrored), a locked aspect ratio, and a
      // hidden shape that must not draw. The names are the selection pane's list.
      layout: 'blank',
      shapes:
        shape({ id: 2, name: 'Arrow', x: 0.5 * IN, y: 0.6 * IN, cx: 2.4 * IN, cy: 1.2 * IN, prst: 'rightArrow', fill: solid('4472C4'), bodyPr: centered, text: label('Original') }) +
        shape({ id: 3, name: 'Arrow flipped H', x: 3.3 * IN, y: 0.6 * IN, cx: 2.4 * IN, cy: 1.2 * IN, prst: 'rightArrow', flipH: true, fill: solid('4472C4'), bodyPr: centered, text: label('Flip horizontal') }) +
        shape({ id: 4, name: 'Arrow flipped V', x: 6.1 * IN, y: 0.6 * IN, cx: 2.4 * IN, cy: 1.2 * IN, prst: 'rtTriangle', flipV: true, fill: solid('ED7D31'), bodyPr: centered, text: label('Flip vertical') }) +
        shape({ id: 5, name: 'Turned and flipped', x: 8.9 * IN, y: 0.6 * IN, cx: 2.4 * IN, cy: 1.2 * IN, prst: 'rightArrow', rot: 330 * 60_000, flipH: true, fill: solid('70AD47'), bodyPr: centered, text: label('30° then flip') }) +
        shape({ id: 6, name: 'Rotated right 90', x: 0.9 * IN, y: 3.2 * IN, cx: 2.4 * IN, cy: 1.2 * IN, prst: 'rightArrow', rot: 90 * 60_000, fill: solid('A5A5A5'), bodyPr: centered, text: label('Rotate right 90°') }) +
        '<p:sp><p:nvSpPr><p:cNvPr id="7" name="Locked ratio"/><p:cNvSpPr><a:spLocks noChangeAspect="1"/></p:cNvSpPr><p:nvPr/></p:nvSpPr>' +
        `<p:spPr><a:xfrm><a:off x="${3.8 * IN}" y="${3 * IN}"/><a:ext cx="${1.6 * IN}" cy="${1.6 * IN}"/></a:xfrm><a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>${solid('FFC000')}</p:spPr>` +
        `<p:txBody>${centered}<a:lstStyle/>${label('Locked')}</p:txBody></p:sp>` +
        '<p:sp><p:nvSpPr><p:cNvPr id="8" name="Hidden shape" hidden="1"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>' +
        `<p:spPr><a:xfrm><a:off x="${6 * IN}" y="${3 * IN}"/><a:ext cx="${2 * IN}" cy="${1.6 * IN}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>${solid('C00000')}</p:spPr>` +
        `<p:txBody>${centered}<a:lstStyle/>${label('Must not draw')}</p:txBody></p:sp>` +
        // Three boxes with equal gaps (Distribute Horizontally).
        [6, 7.4, 9.1]
          .map((x, i) => shape({ id: 9 + i, name: `Distributed ${i + 1}`, x: x * IN, y: 5.2 * IN, cx: (0.8 + i * 0.3) * IN, cy: 0.8 * IN, fill: solid('5B9BD5') }))
          .join(''),
    },
  ],
};
