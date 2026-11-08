// A new chart (`insertChart`): a chart part (ECMA-376 Part 1, 21.2.2.27 chartSpace) with the
// data in its caches, an embedded workbook (a SpreadsheetML package, ECMA-376 Part 1, 18) that
// holds the same data in the table that the chart's cell references name, and a graphic frame
// on the slide that shows the chart (§19.3.1.21 graphicFrame). The data edit (`setChartData`)
// then writes the same table, so the chart and its workbook stay in agreement. DOM-free.

import { writeZip } from '@eigenpal/ooxml-core';
import {
  escapeXml,
  escapeXmlAttribute,
  isValidXmlText,
  NS,
  parseXmlPart,
  REL,
  serializeXmlPart,
} from '../../store/index.ts';
import type { ChartType, NewChartData } from '../../contracts/neutral.ts';
import { toBase64 } from './base64.ts';
import {
  MAX_CHART_LABEL,
  MAX_CHART_POINTS,
  MAX_CHART_SERIES,
  columnLetters,
  rangeFormula,
} from './chart-data-model.ts';
import { axisKindOf, changeChartType, defaultAxesXml, isChartType } from './chart-type.ts';
import { cChild, replaceKid } from './chart-xml.ts';
import { CommandRefusal, type OpBatch } from './op-batch.ts';
import {
  addOverride,
  addRelationship,
  ensureContentType,
  freePartName,
  nextShapeId,
  spTreeOf,
} from './package-edits.ts';
import {
  freeInsertPosition,
  placeholderIndex,
  removeFilledPlaceholder,
  type EmptyPlaceholder,
} from './placeholder-fill.ts';
import { specFromXml } from './shape-commands.ts';

const CHART_CT = 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml';
const XLSX_CT = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
const SML = 'application/vnd.openxmlformats-officedocument.spreadsheetml';
const RELS_CT = 'application/vnd.openxmlformats-package.relationships+xml';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const S = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const PKG_RELS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const REL_PACKAGE = `${R}/package`;
const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';
/** The sheet that holds the chart's table. */
const SHEET = 'Sheet1';

/** The sample data of a new chart: three series over four categories, or one for a pie. */
export function sampleChartData(type: ChartType): NewChartData {
  const categories = ['Category 1', 'Category 2', 'Category 3', 'Category 4'];
  if (type === 'pie' || type === 'doughnut')
    return { categories, series: [{ name: 'Series 1', values: [8.2, 3.2, 1.4, 1.2] }] };
  return {
    categories: axisKindOf(type) === 'xy' ? ['1', '2', '3', '4'] : categories,
    series: [
      { name: 'Series 1', values: [4.3, 2.5, 3.5, 4.5] },
      { name: 'Series 2', values: [2.4, 4.4, 1.8, 2.8] },
      { name: 'Series 3', values: [2, 2, 3, 5] },
    ],
  };
}

const isNumber = (label: string): boolean => label.trim() !== '' && Number.isFinite(Number(label));

/** Check the data of a new chart. Throws `CommandRefusal`. */
function checkData(data: unknown, xy: boolean): NewChartData {
  if (!data || typeof data !== 'object') throw new CommandRefusal('errors.unsupported', 'data');
  const { categories, series } = data as NewChartData;
  if (!Array.isArray(categories) || !Array.isArray(series))
    throw new CommandRefusal('errors.unsupported', 'data');
  if (categories.length === 0 || series.length === 0)
    throw new CommandRefusal('errors.chartData.empty');
  if (categories.length > MAX_CHART_POINTS || series.length > MAX_CHART_SERIES)
    throw new CommandRefusal('errors.chartData.tooLarge');
  const text = (v: unknown): string => {
    if (typeof v !== 'string' || v.length > MAX_CHART_LABEL || !isValidXmlText(v))
      throw new CommandRefusal('errors.chartData.invalidText');
    return v;
  };
  const labels = categories.map(text);
  if (xy && !labels.every(isNumber)) throw new CommandRefusal('errors.chartData.notNumber');
  const rows = series.map((s) => {
    if (!s || typeof s !== 'object' || !Array.isArray(s.values))
      throw new CommandRefusal('errors.unsupported', 'series');
    if (s.values.length !== labels.length) throw new CommandRefusal('errors.unsupported', 'values');
    const values = (s.values as readonly unknown[]).map((v) => {
      if (v === null) return null;
      if (typeof v !== 'number' || !Number.isFinite(v))
        throw new CommandRefusal('errors.chartData.notNumber');
      return v;
    });
    return { name: text(s.name), values };
  });
  return { categories: labels, series: rows };
}

/** A string reference with its cache. */
const strRef = (f: string, values: readonly string[]): string =>
  `<c:strRef><c:f>${escapeXml(f)}</c:f><c:strCache><c:ptCount val="${values.length}"/>${values
    .map((v, i) => `<c:pt idx="${i}"><c:v>${escapeXml(v)}</c:v></c:pt>`)
    .join('')}</c:strCache></c:strRef>`;

/** A number reference with its cache; a blank cell has no point. */
const numRef = (f: string, values: readonly (number | null)[]): string =>
  `<c:numRef><c:f>${escapeXml(f)}</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${values.length}"/>${values
    .map((v, i) => (v === null ? '' : `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`))
    .join('')}</c:numCache></c:numRef>`;

/** Chart text: 9 pt (or `size`), text color at 65 %. */
const textProps = (size = 1197): string =>
  `<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="${size}" b="0" i="0" u="none" strike="noStrike" baseline="0"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr>`;

/** The chart part of a new chart of `type` over `data`, with its workbook as relationship `bookRel`. */
export function newChartSpaceXml(type: ChartType, data: NewChartData, bookRel: string): string {
  const xy = axisKindOf(type) === 'xy';
  const n = data.categories.length;
  const cat = rangeFormula(SHEET, 1, 2, 1, n + 1);
  const categories = xy ? numRef(cat, data.categories.map(Number)) : strRef(cat, data.categories);
  // Built as a clustered column chart, then changed to `type` by the same rules as a type
  // change, so a new chart and a changed one have the same markup.
  const series = data.series
    .map((s, i) => {
      const col = i + 2;
      return (
        `<c:ser><c:idx val="${i}"/><c:order val="${i}"/>` +
        `<c:tx>${strRef(rangeFormula(SHEET, col, 1, col, 1), [s.name])}</c:tx>` +
        `<c:invertIfNegative val="0"/><c:cat>${categories}</c:cat>` +
        `<c:val>${numRef(rangeFormula(SHEET, col, 2, col, n + 1), s.values)}</c:val></c:ser>`
      );
    })
    .join('');
  const ids: [string, string] = ['500000001', '500000002'];
  const xml =
    `<c:chartSpace xmlns:c="${NS.c}" xmlns:a="${NS.a}" xmlns:r="${NS.r}">` +
    '<c:date1904 val="0"/><c:lang val="en-US"/><c:roundedCorners val="0"/>' +
    '<c:chart><c:autoTitleDeleted val="1"/><c:plotArea><c:layout/>' +
    '<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>' +
    `${series}<c:gapWidth val="219"/><c:overlap val="-27"/><c:axId val="${ids[0]}"/><c:axId val="${ids[1]}"/></c:barChart>` +
    `${defaultAxesXml('columnClustered', ids)}<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr></c:plotArea>` +
    `<c:legend><c:legendPos val="b"/><c:overlay val="0"/><c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>${textProps()}</c:legend>` +
    '<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart>' +
    '<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>' +
    `${textProps(1000)}<c:externalData r:id="${escapeXmlAttribute(bookRel)}"><c:autoUpdate val="0"/></c:externalData></c:chartSpace>`;
  if (type === 'columnClustered') return `${XML}${xml}`;
  const parsed = parseXmlPart('/ppt/charts/chart.xml', xml);
  if (!parsed.ok) throw new CommandRefusal('errors.unsupported', 'template');
  const root = parsed.part.root;
  const chart = cChild(root, 'chart')!;
  const plot = cChild(cChild(chart, 'plotArea'), 'barChart')!;
  const next = replaceKid(root, chart, changeChartType(chart, plot, type));
  return serializeXmlPart({ partName: '/ppt/charts/chart.xml', root: next, nextId: 0 });
}

/** A workbook of one sheet: the series names in row 1, the categories in column A. */
export function newChartWorkbook(data: NewChartData, xy: boolean): Uint8Array {
  const strings: string[] = [];
  const sst = (v: string): number => {
    const at = strings.indexOf(v);
    return at >= 0 ? at : strings.push(v) - 1;
  };
  const ref = (c: number, r: number): string => `${columnLetters(c)}${r}`;
  const text = (c: number, r: number, v: string): string =>
    `<c r="${ref(c, r)}" t="s"><v>${sst(v)}</v></c>`;
  const num = (c: number, r: number, v: number | null): string =>
    v === null ? '' : `<c r="${ref(c, r)}"><v>${v}</v></c>`;
  const width = data.series.length + 1;
  const rows: string[] = [
    `<row r="1">${data.series.map((s, i) => text(i + 2, 1, s.name)).join('')}</row>`,
  ];
  data.categories.forEach((label, k) => {
    const r = k + 2;
    const first = xy ? num(1, r, Number(label)) : text(1, r, label);
    rows.push(
      `<row r="${r}">${first}${data.series.map((s, i) => num(i + 2, r, s.values[k] ?? null)).join('')}</row>`
    );
  });
  const encode = (s: string): Uint8Array => new TextEncoder().encode(s);
  const files = new Map<string, Uint8Array>([
    [
      '/[Content_Types].xml',
      encode(
        `${XML}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
          `<Default Extension="rels" ContentType="${RELS_CT}"/><Default Extension="xml" ContentType="application/xml"/>` +
          `<Override PartName="/xl/workbook.xml" ContentType="${SML}.sheet.main+xml"/>` +
          `<Override PartName="/xl/worksheets/sheet1.xml" ContentType="${SML}.worksheet+xml"/>` +
          `<Override PartName="/xl/styles.xml" ContentType="${SML}.styles+xml"/>` +
          `<Override PartName="/xl/sharedStrings.xml" ContentType="${SML}.sharedStrings+xml"/></Types>`
      ),
    ],
    [
      '/_rels/.rels',
      encode(
        `${XML}<Relationships xmlns="${PKG_RELS}"><Relationship Id="rId1" Type="${R}/officeDocument" Target="xl/workbook.xml"/></Relationships>`
      ),
    ],
    [
      '/xl/workbook.xml',
      encode(
        `${XML}<workbook xmlns="${S}" xmlns:r="${R}"><bookViews><workbookView xWindow="0" yWindow="0" windowWidth="16000" windowHeight="9000"/></bookViews>` +
          `<sheets><sheet name="${SHEET}" sheetId="1" r:id="rId1"/></sheets></workbook>`
      ),
    ],
    [
      '/xl/_rels/workbook.xml.rels',
      encode(
        `${XML}<Relationships xmlns="${PKG_RELS}">` +
          `<Relationship Id="rId1" Type="${R}/worksheet" Target="worksheets/sheet1.xml"/>` +
          `<Relationship Id="rId2" Type="${R}/styles" Target="styles.xml"/>` +
          `<Relationship Id="rId3" Type="${R}/sharedStrings" Target="sharedStrings.xml"/></Relationships>`
      ),
    ],
    [
      '/xl/styles.xml',
      encode(
        `${XML}<styleSheet xmlns="${S}"><fonts count="1"><font><sz val="11"/><name val="Calibri"/><family val="2"/></font></fonts>` +
          '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>' +
          '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
          '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
          '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>' +
          '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles><dxfs count="0"/><tableStyles count="0"/></styleSheet>'
      ),
    ],
    [
      '/xl/worksheets/sheet1.xml',
      encode(
        `${XML}<worksheet xmlns="${S}" xmlns:r="${R}"><dimension ref="A1:${ref(width, data.categories.length + 1)}"/>` +
          '<sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/>' +
          `<sheetData>${rows.join('')}</sheetData>` +
          '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>'
      ),
    ],
  ]);
  files.set(
    '/xl/sharedStrings.xml',
    encode(
      `${XML}<sst xmlns="${S}" count="${strings.length}" uniqueCount="${strings.length}">${strings
        .map((v) => `<si><t xml:space="preserve">${escapeXml(v)}</t></si>`)
        .join('')}</sst>`
    )
  );
  return writeZip(files);
}

/**
 * A new chart frame of `type` (`insertChart`). It is centered on the slide, or it takes the place
 * of the empty placeholder `into`: its box, and its index (§19.3.1.36). Returns the frame's node
 * id.
 */
export function insertChartFrame(
  batch: OpBatch,
  part: string,
  type: ChartType,
  data: NewChartData | undefined,
  slideW: number,
  slideH: number,
  into?: EmptyPlaceholder
): string {
  if (!isChartType(type)) throw new CommandRefusal('errors.unsupported', 'chartType');
  const xy = axisKindOf(type) === 'xy';
  const checked = checkData(data ?? sampleChartData(type), xy);
  // The workbook, then the chart part, which names the workbook by its relationship.
  const chartPart = freePartName(batch, '/ppt/charts/chart', '.xml');
  const book = freePartName(batch, '/ppt/embeddings/Workbook', '.xlsx');
  batch.push({ op: 'addPart', part: book, base64: toBase64(newChartWorkbook(checked, xy)) });
  ensureContentType(batch, 'xlsx', XLSX_CT, book);
  const bookRel = addRelationship(batch, chartPart, REL_PACKAGE, book);
  batch.push({
    op: 'addPart',
    part: chartPart,
    base64: toBase64(new TextEncoder().encode(newChartSpaceXml(type, checked, bookRel))),
  });
  addOverride(batch, chartPart, CHART_CT);
  const rId = addRelationship(batch, part, REL.chart, chartPart);
  // Three fifths of the slide wide, at most 4:3 and three fifths high, or the placeholder's
  // box. Without a placeholder, it is centered below the top fifth, where a title usually is.
  const cx = into ? into.cx : Math.round(slideW * 0.6);
  const cy = into ? into.cy : Math.round(Math.min(slideH * 0.6, cx * 0.75));
  const tree = spTreeOf(batch, part);
  const { x, y } = into
    ? into
    : freeInsertPosition(
        tree,
        Math.round((slideW - cx) / 2),
        Math.round(slideH * 0.2 + (slideH * 0.8 - cy) / 2),
        {
          cx,
          cy,
          slideW,
          slideH,
        }
      );
  const id = nextShapeId(batch, tree);
  const name = into ? escapeXmlAttribute(into.name) : `Chart ${id - 1}`;
  const xml =
    `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${id}" name="${name}"/>` +
    '<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr>' +
    `<p:nvPr>${into ? into.ph : ''}</p:nvPr></p:nvGraphicFramePr>` +
    `<p:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></p:xfrm>` +
    `<a:graphic><a:graphicData uri="${NS.c}"><c:chart xmlns:c="${NS.c}" r:id="${rId}"/></a:graphicData></a:graphic></p:graphicFrame>`;
  const at = into ? placeholderIndex(batch, part, into) : -1;
  const ext = tree.children.findIndex((c) => c.type === 'element' && c.local === 'extLst');
  const frame = batch.insert({
    op: 'insertNode',
    part,
    parent: tree.id,
    index: at >= 0 ? at : ext < 0 ? tree.children.length : ext,
    spec: specFromXml(xml),
  });
  if (into && at >= 0) removeFilledPlaceholder(batch, part, into, String(id));
  return frame;
}
