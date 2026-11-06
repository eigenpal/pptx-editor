// Chart and diagram fixtures (ECMA-376 Part 1, 21.2 DrawingML Charts and 21.4 DrawingML
// Diagrams). Each chart slide holds one chart part with cached values and no workbook, so a
// reader draws it from the cache. The diagram parts in `diagrams/` are the anonymous diagram
// parts of DOCX Editor's default sample (docx-editor/scripts/demo-doc/diagrams, EigenPal):
// data, layout, quick style, colors, and the saved drawing. `drawing1.xml` is the drawing that
// the presentation app saves for this frame: a reader that lays the diagram out again gets the
// same text sizes and insets as a reader of the saved drawing.

import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { para, run, shape, type FixtureOptions, type FixtureSlide } from './builder.ts';

const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const C = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const DGM = 'http://schemas.openxmlformats.org/drawingml/2006/diagram';
const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';
const IN = 914_400;
const CHART_CT = 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml';

const title = (text: string): string =>
  shape({
    id: 2,
    x: 0.5 * IN,
    y: 0.25 * IN,
    cx: 12 * IN,
    cy: 0.6 * IN,
    text: para(run(text, 'sz="2000" b="1"')),
    bodyPr: '<a:bodyPr wrap="square" rtlCol="0" anchor="t"><a:noAutofit/></a:bodyPr>',
  });

/** A graphic frame that shows one chart or diagram. */
function frame(id: number, name: string, box: number[], uri: string, data: string): string {
  const [x, y, cx, cy] = box.map(Math.round);
  return (
    `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${id}" name="${name}"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>` +
    `<p:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></p:xfrm>` +
    `<a:graphic><a:graphicData uri="${uri}">${data}</a:graphicData></a:graphic></p:graphicFrame>`
  );
}

// --- Charts ----------------------------------------------------------------------------------

const CATEGORIES = ['North', 'South', 'East', 'West'];
const SERIES: readonly { name: string; values: number[]; color: string }[] = [
  { name: 'Plan', values: [42, 35, 51, 28], color: '4472C4' },
  { name: 'Actual', values: [38, 41, 47, 33], color: 'ED7D31' },
  { name: 'Forecast', values: [45, 39, 55, 36], color: 'A5A5A5' },
];

const strCache = (ref: string, values: readonly string[]): string =>
  `<c:strRef><c:f>${ref}</c:f><c:strCache><c:ptCount val="${values.length}"/>${values
    .map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`)
    .join('')}</c:strCache></c:strRef>`;

const numCache = (ref: string, values: readonly number[], format = 'General'): string =>
  `<c:numRef><c:f>${ref}</c:f><c:numCache><c:formatCode>${format}</c:formatCode><c:ptCount val="${values.length}"/>${values
    .map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`)
    .join('')}</c:numCache></c:numRef>`;

const solidFill = (color: string): string => `<a:solidFill><a:srgbClr val="${color}"/></a:solidFill>`;
const column = (i: number): string => String.fromCharCode(66 + i);

/** One category series: its name, its fill (or line), its categories, and its values. */
function series(i: number, kind: 'fill' | 'line' | 'marker', extra = ''): string {
  const s = SERIES[i]!;
  const spPr =
    kind === 'fill'
      ? `<c:spPr>${solidFill(s.color)}</c:spPr>`
      : `<c:spPr><a:ln w="28575" cap="rnd">${solidFill(s.color)}<a:round/></a:ln></c:spPr>`;
  const marker =
    kind === 'marker'
      ? `<c:marker><c:symbol val="circle"/><c:size val="7"/><c:spPr>${solidFill(s.color)}<a:ln w="9525">${solidFill(s.color)}</a:ln></c:spPr></c:marker>`
      : kind === 'line'
        ? '<c:marker><c:symbol val="none"/></c:marker>'
        : '';
  return (
    `<c:ser><c:idx val="${i}"/><c:order val="${i}"/><c:tx>${strCache(`Sheet1!$${column(i)}$1`, [s.name])}</c:tx>${spPr}${marker}${extra}` +
    `<c:cat>${strCache('Sheet1!$A$2:$A$5', CATEGORIES)}</c:cat>` +
    `<c:val>${numCache(`Sheet1!$${column(i)}$2:$${column(i)}$5`, s.values)}</c:val>` +
    (kind === 'fill' ? '' : '<c:smooth val="0"/>') +
    '</c:ser>'
  );
}

const textProps = (size = 1197, color = '595959'): string =>
  `<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="${size}">${solidFill(color)}</a:defRPr></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr>`;

const gridLine = `<c:majorGridlines><c:spPr><a:ln w="9525">${solidFill('D9D9D9')}</a:ln></c:spPr></c:majorGridlines>`;
const axisLine = `<c:spPr><a:ln w="9525">${solidFill('BFBFBF')}</a:ln></c:spPr>`;

/** A category axis and a value axis, `catPos` for the category side (`b` or `l`). */
function axes(catPos: 'b' | 'l', valuePos: 'l' | 'b', valueFormat = 'General', size = 1197): string {
  return (
    `<c:catAx><c:axId val="1001"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${catPos}"/>` +
    `<c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>${axisLine}${textProps(size)}` +
    '<c:crossAx val="1002"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>' +
    `<c:valAx><c:axId val="1002"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${valuePos}"/>${gridLine}` +
    `<c:numFmt formatCode="${valueFormat}" sourceLinked="0"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:spPr><a:ln><a:noFill/></a:ln></c:spPr>${textProps(size)}` +
    '<c:crossAx val="1001"/><c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx>'
  );
}

const AX_IDS = '<c:axId val="1001"/><c:axId val="1002"/>';

function chartTitle(text: string): string {
  return (
    `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1862" b="0">${solidFill('595959')}</a:defRPr></a:pPr>` +
    `<a:r><a:rPr lang="en-US" sz="1862" b="0">${solidFill('595959')}</a:rPr><a:t>${text}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>`
  );
}

function legend(pos: 'b' | 'r' | 't', size = 1197): string {
  return `<c:legend><c:legendPos val="${pos}"/><c:overlay val="0"/>${textProps(size)}</c:legend>`;
}

function chartSpace(
  plot: string,
  opts: { title?: string; legend?: 'b' | 'r' | 't'; legendSize?: number }
): string {
  return (
    `${XML}<c:chartSpace xmlns:c="${C}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="${R}">` +
    '<c:date1904 val="0"/><c:roundedCorners val="0"/><c:chart>' +
    (opts.title ? chartTitle(opts.title) : '') +
    `<c:autoTitleDeleted val="${opts.title ? 0 : 1}"/><c:plotArea><c:layout/>${plot}</c:plotArea>` +
    (opts.legend ? legend(opts.legend, opts.legendSize) : '') +
    '<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart>' +
    `<c:spPr>${solidFill('FFFFFF')}<a:ln><a:noFill/></a:ln></c:spPr>${textProps(1000, '404040')}</c:chartSpace>`
  );
}

const dataLabels = (show: 'val' | 'percent', position: string): string =>
  `<c:dLbls><c:numFmt formatCode="${show === 'percent' ? '0%' : 'General'}" sourceLinked="0"/><c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>${textProps(1197, show === 'percent' ? 'FFFFFF' : '404040')}` +
  `<c:dLblPos val="${position}"/><c:showLegendKey val="0"/><c:showVal val="${show === 'val' ? 1 : 0}"/><c:showCatName val="0"/><c:showSerName val="0"/>` +
  `<c:showPercent val="${show === 'percent' ? 1 : 0}"/><c:showBubbleSize val="0"/></c:dLbls>`;

/** Varied-color pie points: one fill per category. */
const PIE_COLORS = ['4472C4', 'ED7D31', 'A5A5A5', 'FFC000'];
const piePoints = PIE_COLORS.map(
  (color, i) =>
    `<c:dPt><c:idx val="${i}"/><c:bubble3D val="0"/><c:spPr>${solidFill(color)}<a:ln w="19050">${solidFill('FFFFFF')}</a:ln></c:spPr></c:dPt>`
).join('');

function pieSeries(labelSize = 1197): string {
  const s = SERIES[0]!;
  return (
    `<c:ser><c:idx val="0"/><c:order val="0"/><c:tx>${strCache('Sheet1!$B$1', [s.name])}</c:tx>${piePoints}` +
    dataLabels('percent', 'ctr')
      .replace('<c:dLblPos val="ctr"/>', '')
      .replace('sz="1197"', `sz="${labelSize}"`) +
    `<c:cat>${strCache('Sheet1!$A$2:$A$5', CATEGORIES)}</c:cat><c:val>${numCache('Sheet1!$B$2:$B$5', s.values)}</c:val></c:ser>`
  );
}

function scatterSeries(i: number): string {
  const s = SERIES[i]!;
  const xs = [1, 2.5, 4, 5.5, 7];
  const ys = xs.map((x) => Math.round((s.values[0]! / 10 + x * (i + 1) * 0.9) * 10) / 10);
  return (
    `<c:ser><c:idx val="${i}"/><c:order val="${i}"/><c:tx>${strCache(`Sheet1!$${column(i)}$1`, [s.name])}</c:tx>` +
    `<c:spPr><a:ln w="19050"><a:noFill/></a:ln></c:spPr>` +
    `<c:marker><c:symbol val="circle"/><c:size val="8"/><c:spPr>${solidFill(s.color)}<a:ln w="9525">${solidFill(s.color)}</a:ln></c:spPr></c:marker>` +
    `<c:xVal>${numCache('Sheet1!$A$2:$A$6', xs)}</c:xVal><c:yVal>${numCache(`Sheet1!$${column(i)}$2:$${column(i)}$6`, ys)}</c:yVal><c:smooth val="0"/></c:ser>`
  );
}

function scatterAxes(): string {
  const valAx = (id: number, cross: number, pos: 'b' | 'l', grid: boolean): string =>
    `<c:valAx><c:axId val="${id}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${pos}"/>${grid ? gridLine : ''}` +
    `<c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>${axisLine}${textProps()}` +
    `<c:crossAx val="${cross}"/><c:crosses val="autoZero"/><c:crossBetween val="midCat"/></c:valAx>`;
  return valAx(1001, 1002, 'b', false) + valAx(1002, 1001, 'l', true);
}

const all = (kind: 'fill' | 'line' | 'marker', extra = ''): string =>
  SERIES.map((_, i) => series(i, kind, extra)).join('');

/** The chart slides: `[slide title, chart part XML]`. */
const CHARTS: readonly (readonly [string, string])[] = [
  [
    'Clustered column with title, legend, and axes',
    chartSpace(
      `<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>${all('fill')}<c:gapWidth val="150"/><c:overlap val="-20"/>${AX_IDS}</c:barChart>${axes('b', 'l')}`,
      { title: 'Regional sales', legend: 'b' }
    ),
  ],
  [
    'Stacked bar with value labels',
    chartSpace(
      `<c:barChart><c:barDir val="bar"/><c:grouping val="stacked"/><c:varyColors val="0"/>${all('fill', dataLabels('val', 'ctr').replace('404040', 'FFFFFF'))}<c:gapWidth val="80"/><c:overlap val="100"/>${AX_IDS}</c:barChart>${axes('l', 'b')}`,
      { legend: 'r' }
    ),
  ],
  [
    'Line with markers',
    chartSpace(
      `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${all('marker')}<c:marker val="1"/>${AX_IDS}</c:lineChart>${axes('b', 'l')}`,
      { title: 'Monthly trend', legend: 't' }
    ),
  ],
  [
    'Stacked area',
    chartSpace(
      `<c:areaChart><c:grouping val="stacked"/><c:varyColors val="0"/>${all('fill')}${AX_IDS}</c:areaChart>${axes('b', 'l')}`,
      { legend: 'b' }
    ),
  ],
  [
    'Scatter with markers only',
    chartSpace(
      `<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>${SERIES.map((_, i) => scatterSeries(i)).join('')}${AX_IDS}</c:scatterChart>${scatterAxes()}`,
      { title: 'Effort and output', legend: 'r' }
    ),
  ],
  [
    'Pie with percentage labels',
    chartSpace(`<c:pieChart><c:varyColors val="1"/>${pieSeries()}<c:firstSliceAng val="0"/></c:pieChart>`, {
      title: 'Share by region',
      legend: 'r',
    }),
  ],
  [
    'Doughnut',
    chartSpace(
      `<c:doughnutChart><c:varyColors val="1"/>${pieSeries()}<c:firstSliceAng val="0"/><c:holeSize val="55"/></c:doughnutChart>`,
      { legend: 'b' }
    ),
  ],
  [
    'Radar with markers',
    chartSpace(
      `<c:radarChart><c:radarStyle val="marker"/><c:varyColors val="0"/>${all('marker')}${AX_IDS}</c:radarChart>` +
        axes('b', 'l'),
      { legend: 'r' }
    ),
  ],
];

/** One slide per chart type, each drawn from its cached values. */
export function chartsDeck(): FixtureOptions {
  const slides: FixtureSlide[] = CHARTS.map(([name], i) => ({
    layout: 'blank',
    shapes:
      title(name) +
      frame(3, `Chart ${i + 1}`, [1.5 * IN, 1.1 * IN, 10 * IN, 5.9 * IN], C, `<c:chart xmlns:c="${C}" r:id="rIdChart"/>`),
    rels: [['rIdChart', `${R}/chart`, `../charts/chart${i + 1}.xml`]],
  }));
  return {
    slides,
    extraParts: CHARTS.map(([, xml], i): [string, string, string] => [`ppt/charts/chart${i + 1}.xml`, CHART_CT, xml]),
  };
}

// --- Chart sizes -----------------------------------------------------------------------------

/** Frame sizes in inches, each centred under the slide title. */
const FRAMES = { large: [10, 5.9], medium: [7.5, 4.4], small: [5.5, 3.3] } as const;

type ChartSizeSlide = readonly [
  kind: 'column' | 'pie' | 'radar' | 'line',
  legend: 'b' | 'r' | 't',
  sizePt: number,
  frame: keyof typeof FRAMES,
];

/**
 * Legend-bearing charts at three frame sizes and three text sizes. Every text element of a chart
 * (axes, legend, data labels) takes the slide's size, so the legend and radar spacing rules are
 * measured beyond one text size and one frame.
 */
const CHART_SIZES: readonly ChartSizeSlide[] = [
  ['column', 'r', 9, 'large'],
  ['column', 'r', 14, 'medium'],
  ['column', 'r', 20, 'small'],
  ['column', 'b', 9, 'small'],
  ['line', 'b', 9, 'small'],
  ['column', 'b', 20, 'large'],
  ['pie', 'b', 14, 'small'],
  ['line', 't', 20, 'medium'],
  ['radar', 'r', 9, 'small'],
  ['radar', 'r', 9, 'large'],
  ['radar', 'r', 14, 'medium'],
  ['radar', 'r', 20, 'large'],
];

function sizedChart([kind, legendPos, sizePt]: ChartSizeSlide): string {
  const sz = sizePt * 100;
  const opts = { legend: legendPos, legendSize: sz };
  switch (kind) {
    case 'column':
      return chartSpace(
        `<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>${all('fill')}<c:gapWidth val="150"/><c:overlap val="-20"/>${AX_IDS}</c:barChart>${axes('b', 'l', 'General', sz)}`,
        opts
      );
    case 'line':
      return chartSpace(
        `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${all('marker')}<c:marker val="1"/>${AX_IDS}</c:lineChart>${axes('b', 'l', 'General', sz)}`,
        opts
      );
    case 'pie':
      return chartSpace(
        `<c:pieChart><c:varyColors val="1"/>${pieSeries(sz)}<c:firstSliceAng val="0"/></c:pieChart>`,
        opts
      );
    case 'radar':
      return chartSpace(
        `<c:radarChart><c:radarStyle val="marker"/><c:varyColors val="0"/>${all('marker')}${AX_IDS}</c:radarChart>${axes('b', 'l', 'General', sz)}`,
        opts
      );
  }
}

const KIND_NAMES = { column: 'Column', pie: 'Pie', radar: 'Radar', line: 'Line' } as const;
const LEGEND_NAMES = { b: 'bottom', r: 'right', t: 'top' } as const;

/** One legend-bearing chart per slide, at the text size and frame size of `CHART_SIZES`. */
export function chartSizesDeck(): FixtureOptions {
  const slides: FixtureSlide[] = CHART_SIZES.map((spec, i) => {
    const [kind, legendPos, sizePt, size] = spec;
    const [w, h] = FRAMES[size];
    const box = [((13.333 - w) / 2) * IN, (1.1 + (5.9 - h) / 2) * IN, w * IN, h * IN];
    return {
      layout: 'blank',
      shapes:
        title(`${KIND_NAMES[kind]}, ${LEGEND_NAMES[legendPos]} legend, ${sizePt} pt text, ${w} x ${h} in`) +
        frame(3, `Chart ${i + 1}`, box, C, `<c:chart xmlns:c="${C}" r:id="rIdChart"/>`),
      rels: [['rIdChart', `${R}/chart`, `../charts/chart${i + 1}.xml`]],
    };
  });
  return {
    slides,
    extraParts: CHART_SIZES.map((spec, i): [string, string, string] => [
      `ppt/charts/chart${i + 1}.xml`,
      CHART_CT,
      sizedChart(spec),
    ]),
  };
}

// --- Diagrams --------------------------------------------------------------------------------

const DIAGRAM_PARTS = [
  ['data', 'diagramData', `${R}/diagramData`, 'application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml'],
  ['layout', 'diagramLayout', `${R}/diagramLayout`, 'application/vnd.openxmlformats-officedocument.drawingml.diagramLayout+xml'],
  ['quickStyle', 'diagramQuickStyle', `${R}/diagramQuickStyle`, 'application/vnd.openxmlformats-officedocument.drawingml.diagramStyle+xml'],
  ['colors', 'diagramColors', `${R}/diagramColors`, 'application/vnd.openxmlformats-officedocument.drawingml.diagramColors+xml'],
  ['drawing', 'diagramDrawing', 'http://schemas.microsoft.com/office/2007/relationships/diagramDrawing', 'application/vnd.ms-office.drawingml.diagramDrawing+xml'],
] as const;

const diagramPart = (file: string): string =>
  readFileSync(join(import.meta.dirname, 'diagrams', file), 'utf8');

/**
 * Two diagrams, each with its saved drawing: a nested target (`target2`) and a grouped
 * process list (`lProcess2`). Each frame keeps the extent the drawing was laid out in.
 */
export function diagramsDeck(): FixtureOptions {
  const extents: readonly (readonly [number, number])[] = [
    [5_000_000, 2_314_000],
    [4_400_000, 2_531_000],
  ];
  const slides: FixtureSlide[] = extents.map(([cx, cy], i) => {
    const n = i + 1;
    // The saved drawing is laid out in a frame of exactly this extent, so the frame keeps it.
    const box = [(12_192_000 - cx) / 2, 1.1 * IN + (5.9 * IN - cy) / 2, cx, cy];
    return {
      layout: 'blank',
      shapes:
        title(n === 1 ? 'Nested target diagram' : 'Grouped process list') +
        frame(
          3,
          `Diagram ${n}`,
          box,
          DGM,
          `<dgm:relIds xmlns:dgm="${DGM}" r:dm="rIdDiagramData${n}" r:lo="rIdDiagramLayout${n}" r:qs="rIdDiagramQuickStyle${n}" r:cs="rIdDiagramColors${n}"/>`
        ),
      rels: DIAGRAM_PARTS.map(([file, rel, type]) => [
        rel === 'diagramDrawing' ? `rIdDemoDiagramDrawing${n}` : `rId${rel[0]!.toUpperCase()}${rel.slice(1)}${n}`,
        type,
        `../diagrams/${file}${n}.xml`,
      ]) as FixtureSlide['rels'],
    };
  });
  return {
    slides,
    extraParts: extents.flatMap((_, i) =>
      DIAGRAM_PARTS.map(([file, , , contentType]): [string, string, string] => [
        `ppt/diagrams/${file}${i + 1}.xml`,
        contentType,
        diagramPart(`${file}${i + 1}.xml`),
      ])
    ),
  };
}
