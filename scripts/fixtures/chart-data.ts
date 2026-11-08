// The `chart-data` fixture: charts whose data the Edit Data dialog and `setChartData` edit
// (ECMA-376 Part 1, 21.2). Each chart names its cells in an embedded workbook (`c:externalData`
// to a SpreadsheetML package, ECMA-376 Part 1, 18), and caches the same values. Slide 1 has a
// clustered column chart over a workbook table (`xl/tables/table1.xml`) with shared strings, a
// data point with its own fill (`c:dPt`), and a data label of its own (`c:dLbl`). Slide 2 has a
// scatter chart (`c:xVal`, `c:yVal`) over a workbook with no table.

import { strToU8, zipSync } from 'fflate';
import { para, run, shape, type FixtureOptions, type FixtureSlide } from './builder.ts';

const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const C = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const S = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';
const IN = 914_400;
const CHART_CT = 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml';
const RELS_CT = 'application/vnd.openxmlformats-package.relationships+xml';
const XLSX_CT = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
const SML = 'application/vnd.openxmlformats-officedocument.spreadsheetml';

/** The column chart's data: categories down column A, one series per column. */
export const COLUMN_CATEGORIES = ['North', 'South', 'East', 'West'];
export const COLUMN_SERIES: readonly { name: string; values: number[]; color: string }[] = [
  { name: 'Plan', values: [42, 35, 51, 28], color: '4472C4' },
  { name: 'Actual', values: [38, 41, 47, 33], color: 'ED7D31' },
  { name: 'Forecast', values: [45, 39, 55, 36], color: 'A5A5A5' },
];

/** The scatter chart's data: X values in column A, one Y series per column. */
export const SCATTER_X = [1, 2, 3, 4, 5];
export const SCATTER_SERIES: readonly { name: string; values: number[]; color: string }[] = [
  { name: 'Sample A', values: [2.5, 3.1, 4.8, 5.2, 6.9], color: '4472C4' },
  { name: 'Sample B', values: [1.2, 2.9, 3.3, 4.1, 4.4], color: 'ED7D31' },
];

const col = (i: number): string => String.fromCharCode(65 + i);
const solidFill = (color: string): string => `<a:solidFill><a:srgbClr val="${color}"/></a:solidFill>`;

const strRef = (ref: string, values: readonly string[]): string =>
  `<c:strRef><c:f>${ref}</c:f><c:strCache><c:ptCount val="${values.length}"/>${values
    .map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`)
    .join('')}</c:strCache></c:strRef>`;

const numRef = (ref: string, values: readonly number[]): string =>
  `<c:numRef><c:f>${ref}</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${values.length}"/>${values
    .map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`)
    .join('')}</c:numCache></c:numRef>`;

const textProps = (size = 1197, color = '595959'): string =>
  `<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="${size}">${solidFill(color)}</a:defRPr></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr>`;

const gridLine = `<c:majorGridlines><c:spPr><a:ln w="9525">${solidFill('D9D9D9')}</a:ln></c:spPr></c:majorGridlines>`;

function valueAxis(id: number, cross: number, pos: string, grid: boolean): string {
  return (
    `<c:valAx><c:axId val="${id}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${pos}"/>${grid ? gridLine : ''}` +
    `<c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>` +
    `<c:spPr><a:ln w="9525">${solidFill('BFBFBF')}</a:ln></c:spPr>${textProps()}` +
    `<c:crossAx val="${cross}"/><c:crosses val="autoZero"/><c:crossBetween val="${pos === 'b' ? 'midCat' : 'between'}"/></c:valAx>`
  );
}

const categoryAxis =
  '<c:catAx><c:axId val="1001"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/>' +
  `<c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:spPr><a:ln w="9525">${solidFill('BFBFBF')}</a:ln></c:spPr>${textProps()}` +
  '<c:crossAx val="1002"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>';

function chartSpace(plot: string, title: string): string {
  return (
    `${XML}<c:chartSpace xmlns:c="${C}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="${R}">` +
    '<c:date1904 val="0"/><c:roundedCorners val="0"/><c:chart>' +
    `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1862" b="0">${solidFill('595959')}</a:defRPr></a:pPr>` +
    `<a:r><a:rPr lang="en-US" sz="1862" b="0">${solidFill('595959')}</a:rPr><a:t>${title}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>` +
    `<c:autoTitleDeleted val="0"/><c:plotArea><c:layout/>${plot}</c:plotArea>` +
    `<c:legend><c:legendPos val="b"/><c:overlay val="0"/>${textProps()}</c:legend>` +
    '<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart>' +
    `<c:spPr>${solidFill('FFFFFF')}<a:ln><a:noFill/></a:ln></c:spPr>${textProps(1000, '404040')}` +
    '<c:externalData r:id="rId1"><c:autoUpdate val="0"/></c:externalData></c:chartSpace>'
  );
}

/** One column series. The second has a point of its own color and a data label of its own. */
function columnSeries(i: number): string {
  const s = COLUMN_SERIES[i]!;
  const own =
    i === 1
      ? `<c:dPt><c:idx val="3"/><c:invertIfNegative val="0"/><c:bubble3D val="0"/><c:spPr>${solidFill('C00000')}</c:spPr></c:dPt>` +
        '<c:dLbls><c:dLbl><c:idx val="3"/><c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbl>' +
        `${textProps(1000, '404040')}<c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>`
      : '';
  const n = COLUMN_CATEGORIES.length + 1;
  return (
    `<c:ser><c:idx val="${i}"/><c:order val="${i}"/><c:tx>${strRef(`Sheet1!$${col(i + 1)}$1`, [s.name])}</c:tx>` +
    `<c:spPr>${solidFill(s.color)}</c:spPr><c:invertIfNegative val="0"/>${own}` +
    `<c:cat>${strRef(`Sheet1!$A$2:$A$${n}`, COLUMN_CATEGORIES)}</c:cat>` +
    `<c:val>${numRef(`Sheet1!$${col(i + 1)}$2:$${col(i + 1)}$${n}`, s.values)}</c:val></c:ser>`
  );
}

function scatterSeries(i: number): string {
  const s = SCATTER_SERIES[i]!;
  const n = SCATTER_X.length + 1;
  return (
    `<c:ser><c:idx val="${i}"/><c:order val="${i}"/><c:tx>${strRef(`Sheet1!$${col(i + 1)}$1`, [s.name])}</c:tx>` +
    `<c:spPr><a:ln w="19050">${solidFill(s.color)}</a:ln></c:spPr>` +
    `<c:marker><c:symbol val="circle"/><c:size val="7"/><c:spPr>${solidFill(s.color)}</c:spPr></c:marker>` +
    `<c:xVal>${numRef(`Sheet1!$A$2:$A$${n}`, SCATTER_X)}</c:xVal>` +
    `<c:yVal>${numRef(`Sheet1!$${col(i + 1)}$2:$${col(i + 1)}$${n}`, s.values)}</c:yVal><c:smooth val="0"/></c:ser>`
  );
}

const columnChart = chartSpace(
  '<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>' +
    `${COLUMN_SERIES.map((_, i) => columnSeries(i)).join('')}<c:gapWidth val="150"/><c:overlap val="-20"/>` +
    `<c:axId val="1001"/><c:axId val="1002"/></c:barChart>${categoryAxis}${valueAxis(1002, 1001, 'l', true)}`,
  'Sales by region'
);

const scatterChart = chartSpace(
  '<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>' +
    `${SCATTER_SERIES.map((_, i) => scatterSeries(i)).join('')}<c:axId val="1001"/><c:axId val="1002"/></c:scatterChart>` +
    `${valueAxis(1001, 1002, 'b', false)}${valueAxis(1002, 1001, 'l', true)}`,
  'Samples'
);

// --- Workbooks --------------------------------------------------------------------------------

const cellRef = (c: number, r: number): string => `${col(c)}${r}`;

/**
 * A workbook of one sheet with `header` in row 1 and `rows` below it. Text goes to the shared
 * strings; a table over the whole range when `table` is set.
 */
function workbook(header: readonly string[], rows: readonly (string | number)[][], table: boolean): Uint8Array {
  const strings: string[] = [];
  const sst = (v: string): number => {
    const at = strings.indexOf(v);
    return at >= 0 ? at : strings.push(v) - 1;
  };
  const cell = (c: number, r: number, v: string | number): string =>
    typeof v === 'number'
      ? `<c r="${cellRef(c, r)}"><v>${v}</v></c>`
      : `<c r="${cellRef(c, r)}" t="s"><v>${sst(v)}</v></c>`;
  const all = [header, ...rows];
  const sheetRows = all
    .map((row, r) => `<row r="${r + 1}" spans="1:${row.length}">${row.map((v, c) => cell(c, r + 1, v)).join('')}</row>`)
    .join('');
  const ref = `A1:${cellRef(header.length - 1, all.length)}`;
  const files: Record<string, Uint8Array> = {
    '[Content_Types].xml': strToU8(
      `${XML}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
        `<Default Extension="rels" ContentType="${RELS_CT}"/><Default Extension="xml" ContentType="application/xml"/>` +
        `<Override PartName="/xl/workbook.xml" ContentType="${SML}.sheet.main+xml"/>` +
        `<Override PartName="/xl/worksheets/sheet1.xml" ContentType="${SML}.worksheet+xml"/>` +
        `<Override PartName="/xl/styles.xml" ContentType="${SML}.styles+xml"/>` +
        `<Override PartName="/xl/sharedStrings.xml" ContentType="${SML}.sharedStrings+xml"/>` +
        (table ? `<Override PartName="/xl/tables/table1.xml" ContentType="${SML}.table+xml"/>` : '') +
        '</Types>'
    ),
    '_rels/.rels': strToU8(
      `${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="xl/workbook.xml"/></Relationships>`
    ),
    'xl/workbook.xml': strToU8(
      `${XML}<workbook xmlns="${S}" xmlns:r="${R}"><bookViews><workbookView xWindow="0" yWindow="0" windowWidth="16000" windowHeight="9000"/></bookViews>` +
        '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>'
    ),
    'xl/_rels/workbook.xml.rels': strToU8(
      `${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rId1" Type="${R}/worksheet" Target="worksheets/sheet1.xml"/>` +
        `<Relationship Id="rId2" Type="${R}/styles" Target="styles.xml"/>` +
        `<Relationship Id="rId3" Type="${R}/sharedStrings" Target="sharedStrings.xml"/></Relationships>`
    ),
    'xl/styles.xml': strToU8(
      `${XML}<styleSheet xmlns="${S}"><fonts count="1"><font><sz val="11"/><name val="Carlito"/><family val="2"/></font></fonts>` +
        '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>' +
        '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
        '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>' +
        '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles><dxfs count="0"/><tableStyles count="0"/></styleSheet>'
    ),
  };
  const sheet =
    `${XML}<worksheet xmlns="${S}" xmlns:r="${R}"><dimension ref="${ref}"/><sheetViews><sheetView workbookViewId="0"/></sheetViews>` +
    `<sheetFormatPr defaultRowHeight="15"/><sheetData>${sheetRows}</sheetData>` +
    '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>' +
    (table ? '<tableParts count="1"><tablePart r:id="rId1"/></tableParts>' : '') +
    '</worksheet>';
  files['xl/worksheets/sheet1.xml'] = strToU8(sheet);
  if (table) {
    files['xl/worksheets/_rels/sheet1.xml.rels'] = strToU8(
      `${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/table" Target="../tables/table1.xml"/></Relationships>`
    );
    files['xl/tables/table1.xml'] = strToU8(
      `${XML}<table xmlns="${S}" id="1" name="Table1" displayName="Table1" ref="${ref}" totalsRowShown="0">` +
        `<tableColumns count="${header.length}">${header.map((h, i) => `<tableColumn id="${i + 1}" name="${h}"/>`).join('')}</tableColumns>` +
        '<tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/></table>'
    );
  }
  files['xl/sharedStrings.xml'] = strToU8(
    `${XML}<sst xmlns="${S}" count="${strings.length}" uniqueCount="${strings.length}">${strings.map((v) => `<si><t>${v}</t></si>`).join('')}</sst>`
  );
  return zipSync(files, { mtime: new Date(Date.UTC(1980, 0, 2, 12)) });
}

const rels = (target: string): string =>
  `${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
  `<Relationship Id="rId1" Type="${R}/package" Target="${target}"/></Relationships>`;

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

const frame =
  '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="3" name="Chart 1"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>' +
  `<p:xfrm><a:off x="${Math.round(1.5 * IN)}" y="${Math.round(1.1 * IN)}"/><a:ext cx="${10 * IN}" cy="${Math.round(5.9 * IN)}"/></p:xfrm>` +
  `<a:graphic><a:graphicData uri="${C}"><c:chart xmlns:c="${C}" r:id="rIdChart"/></a:graphicData></a:graphic></p:graphicFrame>`;

/** The deck: a column chart and a scatter chart, each with its embedded workbook. */
export function chartDataDeck(): FixtureOptions {
  const slides: FixtureSlide[] = ['Column chart with a workbook table', 'Scatter chart'].map(
    (name, i) => ({
      layout: 'blank',
      shapes: title(name) + frame,
      rels: [['rIdChart', `${R}/chart`, `../charts/chart${i + 1}.xml`]],
    })
  );
  const columnBook = workbook(
    ['Region', ...COLUMN_SERIES.map((s) => s.name)],
    COLUMN_CATEGORIES.map((c, k) => [c, ...COLUMN_SERIES.map((s) => s.values[k]!)]),
    true
  );
  const scatterBook = workbook(
    ['X', ...SCATTER_SERIES.map((s) => s.name)],
    SCATTER_X.map((x, k) => [x, ...SCATTER_SERIES.map((s) => s.values[k]!)]),
    false
  );
  return {
    slides,
    binaryParts: [
      { name: 'ppt/embeddings/Workbook1.xlsx', bytes: columnBook, contentType: XLSX_CT },
      { name: 'ppt/embeddings/Workbook2.xlsx', bytes: scatterBook, contentType: XLSX_CT },
    ],
    extraParts: [
      ['ppt/charts/chart1.xml', CHART_CT, columnChart],
      ['ppt/charts/_rels/chart1.xml.rels', RELS_CT, rels('../embeddings/Workbook1.xlsx')],
      ['ppt/charts/chart2.xml', CHART_CT, scatterChart],
      ['ppt/charts/_rels/chart2.xml.rels', RELS_CT, rels('../embeddings/Workbook2.xlsx')],
    ],
  };
}
