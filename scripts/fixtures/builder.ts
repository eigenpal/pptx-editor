// A minimal, valid PresentationML package builder for synthetic fixtures and tests.
//
// Writes the parts that a reader requires to open a deck without a repair prompt: content types,
// package and presentation relationships, document properties, presentation, presProps,
// viewProps, tableStyles, one theme, one master, title and blank layouts, slides, and optional
// notes and media. Slide content is supplied as `p:spTree` children XML.

import { zipSync, strToU8 } from 'fflate';

export const NS_DECL =
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';

const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const CT = 'application/vnd.openxmlformats-officedocument.presentationml';

/** One slide. `shapes` is the XML of `p:spTree` children after the group properties. */
export interface FixtureSlide {
  readonly shapes: string;
  /** `p:bg` XML, e.g. `<p:bg><p:bgPr>…</p:bgPr></p:bg>`. */
  readonly background?: string;
  readonly layout?: 'title' | 'blank';
  /** The `name` of one of `FixtureOptions.layouts`; it wins over `layout`. */
  readonly layoutName?: string;
  readonly notes?: string;
  /** Notes paragraphs as `a:p` XML (formatted runs); used instead of `notes` when given. */
  readonly notesXml?: string;
  /** Extra relationships: `[id, type, target]`, or `[id, type, target, 'External']`. */
  readonly rels?: readonly (
    | readonly [string, string, string]
    | readonly [string, string, string, 'External']
  )[];
  /** The slide's `p:extLst` XML (after `p:clrMapOvr`). */
  readonly ext?: string;
  /**
   * The child of the slide's `p:clrMapOvr`: `<a:masterClrMapping/>` (the default) or an
   * `a:overrideClrMapping` (ECMA-376 Part 1, 19.3.1.7 clrMapOvr).
   */
  readonly clrMapOvr?: string;
  /** A hidden slide (`p:sld/@show="0"`): a slide show skips it. */
  readonly hidden?: boolean;
}

/** A slide layout after the two built-in ones (`slideLayout3.xml` …), on the one master. */
export interface FixtureLayout {
  /** `p:cSld/@name`; a slide picks the layout with `FixtureSlide.layoutName`. */
  readonly name: string;
  /** `p:sldLayout/@type` (`title`, `twoObj`, `titleOnly`, …); omitted when absent. */
  readonly type?: string;
  /** The `p:spTree` children XML. */
  readonly shapes: string;
}

/** A media part: `ppt/media/<name>`. */
export interface FixtureMedia {
  readonly name: string;
  readonly bytes: Uint8Array;
  readonly contentType: string;
}

export interface FixtureOptions {
  /** Slide size in EMU. Default 16:9, 12192000 × 6858000. */
  readonly width?: number;
  readonly height?: number;
  readonly slides: readonly FixtureSlide[];
  readonly media?: readonly FixtureMedia[];
  /**
   * Give the master and layouts the standard slide-number placeholder (default theme position),
   * so Insert > Slide numbers works. Default false: fixtures keep their parts unchanged.
   */
  readonly footers?: boolean;
  /** Theme colors (hex, no `#`) for every slot. Default: the built-in default palette. */
  readonly colors?: ThemeColors;
  /** Theme major/minor Latin fonts. Default Calibri Light / Calibri. */
  readonly majorFont?: string;
  readonly minorFont?: string;
  /** The theme's `a:theme/@name`, which a layout picker shows. Default `Fixture`. */
  readonly themeName?: string;
  /** Extra XML parts: `[partName without leading slash, contentType, xml]`. */
  readonly extraParts?: readonly (readonly [string, string, string])[];
  /** `a:tblStyle` elements written into `ppt/tableStyles.xml`. */
  readonly tableStyles?: string;
  /** Extra presentation relationships: `[id, type, target relative to ppt/]`. */
  readonly presentationRels?: readonly (readonly [string, string, string])[];
  /** XML after `p:notesSz` and before `p:defaultTextStyle` (`p:custShowLst`, for example). */
  readonly presentationLists?: string;
  /** The presentation's `p:extLst` XML (sections, for example). */
  readonly presentationExt?: string;
  /** Binary parts by full name (`ppt/embeddings/x.xlsx`), typed by a default per extension. */
  readonly binaryParts?: readonly FixtureMedia[];
  /** More layouts after Title and Content and Blank. */
  readonly layouts?: readonly FixtureLayout[];
  /**
   * The order of the master's layout list (`p:sldLayoutIdLst`), which a layout picker shows, by
   * layout name. Layouts it does not name follow in their default order. Default: Title and
   * Content, Blank, then `layouts`.
   */
  readonly layoutOrder?: readonly string[];
}

/** The twelve theme color slots, as hex without `#`. */
export interface ThemeColors {
  readonly dk1: string;
  readonly lt1: string;
  readonly dk2: string;
  readonly lt2: string;
  readonly accent1: string;
  readonly accent2: string;
  readonly accent3: string;
  readonly accent4: string;
  readonly accent5: string;
  readonly accent6: string;
  readonly hlink: string;
  readonly folHlink: string;
}

export function spTree(children: string): string {
  return (
    '<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
    '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/>' +
    `<a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>${children}</p:spTree>`
  );
}

const rels = (items: readonly (readonly string[])[]): string =>
  `${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
  items
    .map(([id, type, target, mode]) => `<Relationship Id="${id}" Type="${type}" Target="${target}"${mode ? ` TargetMode="${mode}"` : ''}/>`)
    .join('') +
  '</Relationships>';

function theme(major: string, minor: string, colors?: ThemeColors, name = 'Fixture'): string {
  const clr = (name: string, hex: string): string => `<a:${name}><a:srgbClr val="${colors?.[name as keyof ThemeColors] ?? hex}"/></a:${name}>`;
  const sys = (name: 'dk1' | 'lt1', val: string, last: string): string =>
    colors ? clr(name, colors[name]) : `<a:${name}><a:sysClr val="${val}" lastClr="${last}"/></a:${name}>`;
  const fill = '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>';
  const line = (w: number): string =>
    `<a:ln w="${w}" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>`;
  return (
    `${XML}<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="${name}"><a:themeElements>` +
    '<a:clrScheme name="Default">' +
    sys('dk1', 'windowText', '000000') +
    sys('lt1', 'window', 'FFFFFF') +
    clr('dk2', '44546A') + clr('lt2', 'E7E6E6') + clr('accent1', '4472C4') + clr('accent2', 'ED7D31') +
    clr('accent3', 'A5A5A5') + clr('accent4', 'FFC000') + clr('accent5', '5B9BD5') + clr('accent6', '70AD47') +
    clr('hlink', '0563C1') + clr('folHlink', '954F72') + '</a:clrScheme>' +
    `<a:fontScheme name="${name}"><a:majorFont><a:latin typeface="${major}"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>` +
    `<a:minorFont><a:latin typeface="${minor}"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme>` +
    '<a:fmtScheme name="Default"><a:fillStyleLst>' + fill + fill + fill + '</a:fillStyleLst>' +
    '<a:lnStyleLst>' + line(6350) + line(12700) + line(19050) + '</a:lnStyleLst>' +
    '<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle>' +
    '<a:effectStyle><a:effectLst><a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0">' +
    '<a:srgbClr val="000000"><a:alpha val="63000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle></a:effectStyleLst>' +
    '<a:bgFillStyleLst>' + fill + fill + fill + '</a:bgFillStyleLst></a:fmtScheme></a:themeElements>' +
    '<a:objectDefaults/><a:extraClrSchemeLst/></a:theme>'
  );
}

const level = (n: number, sz: number, extra = ''): string =>
  `<a:lvl${n}pPr marL="${(n - 1) * 457200 + (extra ? 228600 : 0)}" indent="${extra ? -228600 : 0}" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">` +
  `<a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:spcBef><a:spcPts val="${extra ? 1000 : 0}"/></a:spcBef>${extra}` +
  `<a:defRPr sz="${sz}" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl${n}pPr>`;

/** The slide-number placeholder of the default theme: on the master with its frame and style. */
const MASTER_SLIDE_NUMBER =
  '<p:sp><p:nvSpPr><p:cNvPr id="6" name="Slide Number Placeholder 5"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldNum" sz="quarter" idx="4"/></p:nvPr></p:nvSpPr>' +
  '<p:spPr><a:xfrm><a:off x="8610600" y="6356350"/><a:ext cx="2743200" cy="365125"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>' +
  '<p:txBody><a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0" anchor="ctr"/><a:lstStyle><a:lvl1pPr algn="r"><a:defRPr sz="1200"><a:solidFill><a:schemeClr val="tx1"><a:tint val="75000"/></a:schemeClr></a:solidFill></a:defRPr></a:lvl1pPr></a:lstStyle>' +
  '<a:p><a:fld id="{B6F15528-21DE-4FAA-801E-634DDDAF4B2B}" type="slidenum"><a:rPr lang="en-US"/><a:t>‹#›</a:t></a:fld><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>';
export const LAYOUT_SLIDE_NUMBER =
  '<p:sp><p:nvSpPr><p:cNvPr id="6" name="Slide Number Placeholder 5"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldNum" sz="quarter" idx="12"/></p:nvPr></p:nvSpPr><p:spPr/>' +
  '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:fld id="{B6F15528-21DE-4FAA-801E-634DDDAF4B2B}" type="slidenum"><a:rPr lang="en-US"/><a:t>‹#›</a:t></a:fld><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>';

function master(footers = false, extraLayouts: readonly string[] = [], order?: readonly string[]): string {
  const ids = [
    { name: 'Title and Content', xml: '<p:sldLayoutId id="2147483649" r:id="rId1"/>' },
    { name: 'Blank', xml: '<p:sldLayoutId id="2147483650" r:id="rId2"/>' },
    ...extraLayouts.map((name, i) => ({ name, xml: `<p:sldLayoutId id="${2147483651 + i}" r:id="rIdLayout${i + 3}"/>` })),
  ];
  const rank = (name: string): number => {
    const at = order?.indexOf(name) ?? -1;
    return at < 0 ? Number.MAX_SAFE_INTEGER : at;
  };
  const listed = order ? [...ids].sort((a, b) => rank(a.name) - rank(b.name)) : ids;
  const bullet = '<a:buFont typeface="Arial" panose="020B0604020202020204"/><a:buChar char="&#8226;"/>';
  return (
    `${XML}<p:sldMaster ${NS_DECL}><p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg>` +
    spTree(
      '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title Placeholder 1"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>' +
        '<p:spPr><a:xfrm><a:off x="838200" y="365125"/><a:ext cx="10515600" cy="1325563"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>' +
        '<p:txBody><a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0" anchor="ctr"><a:normAutofit/></a:bodyPr><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>Click to edit Master title style</a:t></a:r></a:p></p:txBody></p:sp>' +
        '<p:sp><p:nvSpPr><p:cNvPr id="3" name="Text Placeholder 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>' +
        '<p:spPr><a:xfrm><a:off x="838200" y="1825625"/><a:ext cx="10515600" cy="4351338"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>' +
        '<p:txBody><a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0"><a:normAutofit/></a:bodyPr><a:lstStyle/><a:p><a:pPr lvl="0"/><a:r><a:rPr lang="en-US"/><a:t>Click to edit Master text styles</a:t></a:r></a:p></p:txBody></p:sp>' +
        (footers ? MASTER_SLIDE_NUMBER : '')
    ) +
    '</p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>' +
    `<p:sldLayoutIdLst>${listed.map((l) => l.xml).join('')}</p:sldLayoutIdLst>` +
    '<p:txStyles><p:titleStyle><a:lvl1pPr algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:spcBef><a:spcPct val="0"/></a:spcBef><a:buNone/>' +
    '<a:defRPr sz="4400" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mj-lt"/><a:ea typeface="+mj-ea"/><a:cs typeface="+mj-cs"/></a:defRPr></a:lvl1pPr></p:titleStyle>' +
    `<p:bodyStyle>${level(1, 2800, bullet)}${level(2, 2400, bullet)}${level(3, 2000, bullet)}</p:bodyStyle>` +
    `<p:otherStyle><a:defPPr><a:defRPr lang="en-US"/></a:defPPr>${level(1, 1800)}</p:otherStyle></p:txStyles></p:sldMaster>`
  );
}

function layout(kind: 'title' | 'blank', footers = false): string {
  const shapes =
    kind === 'title'
      ? '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title 1"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr/>' +
        '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>Click to edit Master title style</a:t></a:r></a:p></p:txBody></p:sp>' +
        '<p:sp><p:nvSpPr><p:cNvPr id="3" name="Content Placeholder 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph idx="1"/></p:nvPr></p:nvSpPr><p:spPr/>' +
        '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:pPr lvl="0"/><a:r><a:rPr lang="en-US"/><a:t>Click to edit Master text styles</a:t></a:r></a:p></p:txBody></p:sp>'
      : '';
  return (
    `${XML}<p:sldLayout ${NS_DECL} type="${kind === 'title' ? 'obj' : 'blank'}" preserve="1"><p:cSld name="${kind === 'title' ? 'Title and Content' : 'Blank'}">` +
    spTree(shapes + (footers ? LAYOUT_SLIDE_NUMBER : '')) +
    '</p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>'
  );
}

function notesMaster(): string {
  return (
    `${XML}<p:notesMaster ${NS_DECL}><p:cSld>` +
    spTree(
      '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Notes Placeholder 1"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" sz="quarter" idx="1"/></p:nvPr></p:nvSpPr>' +
        '<p:spPr><a:xfrm><a:off x="685800" y="4400550"/><a:ext cx="5486400" cy="3600450"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>' +
        '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>'
    ) +
    '</p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>' +
    `<p:notesStyle>${level(1, 1200)}</p:notesStyle></p:notesMaster>`
  );
}

export function escapeText(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function notesSlide(text: string, xml?: string): string {
  const paras =
    xml ??
    text
      .split('\n')
      .map((t) => `<a:p><a:r><a:rPr lang="en-US" dirty="0"/><a:t>${escapeText(t)}</a:t></a:r></a:p>`)
      .join('');
  return (
    `${XML}<p:notes ${NS_DECL}><p:cSld>` +
    spTree(
      '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Slide Image Placeholder 1"/><p:cNvSpPr><a:spLocks noGrp="1" noRot="1" noChangeAspect="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldImg"/></p:nvPr></p:nvSpPr><p:spPr/></p:sp>' +
        `<p:sp><p:nvSpPr><p:cNvPr id="3" name="Notes Placeholder 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>${paras}</p:txBody></p:sp>`
    ) +
    '</p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:notes>'
  );
}

/** Build a `.pptx` package. */
export function buildPptx(opts: FixtureOptions): Uint8Array {
  const width = opts.width ?? 12_192_000;
  const height = opts.height ?? 6_858_000;
  const files: Record<string, Uint8Array> = {};
  const put = (name: string, text: string): void => {
    files[name] = strToU8(text);
  };
  const overrides: [string, string][] = [
    ['/ppt/presentation.xml', `${CT}.presentation.main+xml`],
    ['/ppt/slideMasters/slideMaster1.xml', `${CT}.slideMaster+xml`],
    ['/ppt/slideLayouts/slideLayout1.xml', `${CT}.slideLayout+xml`],
    ['/ppt/slideLayouts/slideLayout2.xml', `${CT}.slideLayout+xml`],
    ...(opts.layouts ?? []).map((_, i): [string, string] => [`/ppt/slideLayouts/slideLayout${i + 3}.xml`, `${CT}.slideLayout+xml`]),
    ['/ppt/theme/theme1.xml', 'application/vnd.openxmlformats-officedocument.theme+xml'],
    ['/ppt/theme/theme2.xml', 'application/vnd.openxmlformats-officedocument.theme+xml'],
    ['/ppt/notesMasters/notesMaster1.xml', `${CT}.notesMaster+xml`],
    ['/ppt/presProps.xml', `${CT}.presProps+xml`],
    ['/ppt/viewProps.xml', `${CT}.viewProps+xml`],
    ['/ppt/tableStyles.xml', `${CT}.tableStyles+xml`],
    ['/docProps/core.xml', 'application/vnd.openxmlformats-package.core-properties+xml'],
    ['/docProps/app.xml', 'application/vnd.openxmlformats-officedocument.extended-properties+xml'],
  ];
  const presRels: [string, string, string][] = [
    ['rId1', `${R}/slideMaster`, 'slideMasters/slideMaster1.xml'],
    ['rId2', `${R}/notesMaster`, 'notesMasters/notesMaster1.xml'],
    ['rId3', `${R}/presProps`, 'presProps.xml'],
    ['rId4', `${R}/viewProps`, 'viewProps.xml'],
    ['rId5', `${R}/theme`, 'theme/theme1.xml'],
    ['rId6', `${R}/tableStyles`, 'tableStyles.xml'],
  ];
  const sldIds: string[] = [];
  opts.slides.forEach((slide, i) => {
    const n = i + 1;
    const rId = `rId${100 + n}`;
    presRels.push([rId, `${R}/slide`, `slides/slide${n}.xml`]);
    sldIds.push(`<p:sldId id="${255 + n}" r:id="${rId}"/>`);
    overrides.push([`/ppt/slides/slide${n}.xml`, `${CT}.slide+xml`]);
    const named = slide.layoutName === undefined ? -1 : (opts.layouts ?? []).findIndex((l) => l.name === slide.layoutName);
    if (slide.layoutName !== undefined && named < 0) throw new Error(`no layout ${slide.layoutName}`);
    const layoutNumber = named >= 0 ? named + 3 : slide.layout === 'blank' ? 2 : 1;
    const slideRels: string[][] = [
      ['rId1', `${R}/slideLayout`, `../slideLayouts/slideLayout${layoutNumber}.xml`],
      ...(slide.rels ?? []).map((r) => [...r]),
    ];
    if (slide.notes !== undefined || slide.notesXml !== undefined) {
      overrides.push([`/ppt/notesSlides/notesSlide${n}.xml`, `${CT}.notesSlide+xml`]);
      slideRels.push(['rIdNotes', `${R}/notesSlide`, `../notesSlides/notesSlide${n}.xml`]);
      put(`ppt/notesSlides/notesSlide${n}.xml`, notesSlide(slide.notes ?? '', slide.notesXml));
      put(
        `ppt/notesSlides/_rels/notesSlide${n}.xml.rels`,
        rels([
          ['rId1', `${R}/notesMaster`, '../notesMasters/notesMaster1.xml'],
          ['rId2', `${R}/slide`, `../slides/slide${n}.xml`],
        ])
      );
    }
    put(
      `ppt/slides/slide${n}.xml`,
      `${XML}<p:sld ${NS_DECL}${slide.hidden ? ' show="0"' : ''}><p:cSld>${slide.background ?? ''}${spTree(slide.shapes)}</p:cSld><p:clrMapOvr>${slide.clrMapOvr ?? '<a:masterClrMapping/>'}</p:clrMapOvr>${slide.ext ?? ''}</p:sld>`
    );
    put(`ppt/slides/_rels/slide${n}.xml.rels`, rels(slideRels));
  });
  for (const [name, type, xml] of opts.extraParts ?? []) {
    overrides.push([`/${name}`, type]);
    put(name, xml);
  }
  const defaults = new Map<string, string>([
    ['rels', 'application/vnd.openxmlformats-package.relationships+xml'],
    ['xml', 'application/xml'],
  ]);
  for (const m of opts.media ?? []) {
    files[`ppt/media/${m.name}`] = m.bytes;
    defaults.set(m.name.slice(m.name.lastIndexOf('.') + 1), m.contentType);
  }
  for (const b of opts.binaryParts ?? []) {
    files[b.name] = b.bytes;
    defaults.set(b.name.slice(b.name.lastIndexOf('.') + 1), b.contentType);
  }
  for (const r of opts.presentationRels ?? []) presRels.push([...r] as [string, string, string]);
  put(
    '[Content_Types].xml',
    `${XML}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      [...defaults].map(([e, t]) => `<Default Extension="${e}" ContentType="${t}"/>`).join('') +
      overrides.map(([p, t]) => `<Override PartName="${p}" ContentType="${t}"/>`).join('') +
      '</Types>'
  );
  put(
    '_rels/.rels',
    rels([
      ['rId1', `${R}/officeDocument`, 'ppt/presentation.xml'],
      ['rId2', 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties', 'docProps/core.xml'],
      ['rId3', `${R}/extended-properties`, 'docProps/app.xml'],
    ])
  );
  put(
    'docProps/core.xml',
    `${XML}<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>Fixture</dc:title><dc:creator>PPTX Editor fixtures</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">2026-01-01T00:00:00Z</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">2026-01-01T00:00:00Z</dcterms:modified></cp:coreProperties>`
  );
  put(
    'docProps/app.xml',
    `${XML}<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>PPTX Editor</Application><Slides>${opts.slides.length}</Slides></Properties>`
  );
  put(
    'ppt/presentation.xml',
    `${XML}<p:presentation ${NS_DECL} saveSubsetFonts="1"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>` +
      `<p:notesMasterIdLst><p:notesMasterId r:id="rId2"/></p:notesMasterIdLst><p:sldIdLst>${sldIds.join('')}</p:sldIdLst>` +
      `<p:sldSz cx="${width}" cy="${height}"/><p:notesSz cx="6858000" cy="9144000"/>${opts.presentationLists ?? ''}` +
      `<p:defaultTextStyle><a:defPPr><a:defRPr lang="en-US"/></a:defPPr>${level(1, 1800)}${level(2, 1800)}</p:defaultTextStyle>${opts.presentationExt ?? ''}</p:presentation>`
  );
  put('ppt/_rels/presentation.xml.rels', rels(presRels));
  put('ppt/presProps.xml', `${XML}<p:presentationPr ${NS_DECL}/>`);
  put(
    'ppt/viewProps.xml',
    `${XML}<p:viewPr ${NS_DECL}><p:normalViewPr><p:restoredLeft sz="15620"/><p:restoredTop sz="94660"/></p:normalViewPr><p:gridSpacing cx="76200" cy="76200"/></p:viewPr>`
  );
  put(
    'ppt/tableStyles.xml',
    opts.tableStyles
      ? `${XML}<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}">${opts.tableStyles}</a:tblStyleLst>`
      : `${XML}<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"/>`
  );
  put(
    'ppt/theme/theme1.xml',
    theme(opts.majorFont ?? 'Calibri Light', opts.minorFont ?? 'Calibri', opts.colors, opts.themeName)
  );
  put('ppt/theme/theme2.xml', theme('Calibri Light', 'Calibri'));
  const extra = opts.layouts ?? [];
  put(
    'ppt/slideMasters/slideMaster1.xml',
    master(
      opts.footers,
      extra.map((l) => l.name),
      opts.layoutOrder
    )
  );
  put(
    'ppt/slideMasters/_rels/slideMaster1.xml.rels',
    rels([
      ['rId1', `${R}/slideLayout`, '../slideLayouts/slideLayout1.xml'],
      ['rId2', `${R}/slideLayout`, '../slideLayouts/slideLayout2.xml'],
      ['rId3', `${R}/theme`, '../theme/theme1.xml'],
      ...extra.map((_, i) => [`rIdLayout${i + 3}`, `${R}/slideLayout`, `../slideLayouts/slideLayout${i + 3}.xml`]),
    ])
  );
  extra.forEach((l, i) => {
    put(
      `ppt/slideLayouts/slideLayout${i + 3}.xml`,
      `${XML}<p:sldLayout ${NS_DECL}${l.type ? ` type="${l.type}"` : ''} preserve="1"><p:cSld name="${escapeText(l.name)}">` +
        spTree(l.shapes) +
        '</p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>'
    );
    put(`ppt/slideLayouts/_rels/slideLayout${i + 3}.xml.rels`, rels([['rId1', `${R}/slideMaster`, '../slideMasters/slideMaster1.xml']]));
  });
  for (const [n, kind] of [
    [1, 'title'],
    [2, 'blank'],
  ] as const) {
    put(`ppt/slideLayouts/slideLayout${n}.xml`, layout(kind, opts.footers));
    put(`ppt/slideLayouts/_rels/slideLayout${n}.xml.rels`, rels([['rId1', `${R}/slideMaster`, '../slideMasters/slideMaster1.xml']]));
  }
  put('ppt/notesMasters/notesMaster1.xml', notesMaster());
  put('ppt/notesMasters/_rels/notesMaster1.xml.rels', rels([['rId1', `${R}/theme`, '../theme/theme2.xml']]));
  return zipSync(files, { mtime: new Date(Date.UTC(1980, 0, 2, 12)) });
}

/** A `p:sp` with preset geometry and optional text body XML. */
export function shape(o: {
  id: number;
  name?: string;
  x: number;
  y: number;
  cx: number;
  cy: number;
  prst?: string;
  rot?: number;
  flipH?: boolean;
  flipV?: boolean;
  fill?: string;
  line?: string;
  effects?: string;
  text?: string;
  bodyPr?: string;
  /** The `a:lstStyle` content (level properties). */
  lstStyle?: string;
  ph?: string;
  avLst?: string;
  style?: string;
}): string {
  const n = (v: number): number => Math.round(v);
  const xfrmAttrs = `${o.rot ? ` rot="${n(o.rot)}"` : ''}${o.flipH ? ' flipH="1"' : ''}${o.flipV ? ' flipV="1"' : ''}`;
  return (
    `<p:sp><p:nvSpPr><p:cNvPr id="${o.id}" name="${o.name ?? `Shape ${o.id}`}"/><p:cNvSpPr/><p:nvPr>${o.ph ?? ''}</p:nvPr></p:nvSpPr>` +
    `<p:spPr><a:xfrm${xfrmAttrs}><a:off x="${n(o.x)}" y="${n(o.y)}"/><a:ext cx="${n(o.cx)}" cy="${n(o.cy)}"/></a:xfrm>` +
    `<a:prstGeom prst="${o.prst ?? 'rect'}"><a:avLst>${o.avLst ?? ''}</a:avLst></a:prstGeom>${o.fill ?? ''}${o.line ?? ''}${o.effects ?? ''}</p:spPr>` +
    (o.style ?? '') +
    (o.text !== undefined
      ? `<p:txBody>${o.bodyPr ?? '<a:bodyPr wrap="square" rtlCol="0"><a:spAutoFit/></a:bodyPr>'}${o.lstStyle ? `<a:lstStyle>${o.lstStyle}</a:lstStyle>` : '<a:lstStyle/>'}${o.text}</p:txBody>`
      : '') +
    '</p:sp>'
  );
}

/** One run: `a:rPr` attributes and children, then the text. */
export function run(text: string, attrs = '', children = ''): string {
  return `<a:r><a:rPr lang="en-US" dirty="0"${attrs ? ` ${attrs}` : ''}>${children}</a:rPr><a:t>${escapeText(text)}</a:t></a:r>`;
}

/** A paragraph from run XML and optional `a:pPr` XML. */
export function para(runs: string, pPr = ''): string {
  return `<a:p>${pPr}${runs}</a:p>`;
}

/** A table graphic frame. `rows` are `a:tr` elements; `tblPr` is the whole `a:tblPr`. */
export function table(o: {
  id: number;
  x: number;
  y: number;
  cx: number;
  cy: number;
  columns: readonly number[];
  tblPr: string;
  rows: string;
}): string {
  return (
    `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${o.id}" name="Table ${o.id - 1}"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>` +
    `<p:xfrm><a:off x="${o.x}" y="${o.y}"/><a:ext cx="${o.cx}" cy="${o.cy}"/></p:xfrm>` +
    '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl>' +
    o.tblPr +
    `<a:tblGrid>${o.columns.map((w) => `<a:gridCol w="${w}"/>`).join('')}</a:tblGrid>${o.rows}</a:tbl></a:graphicData></a:graphic></p:graphicFrame>`
  );
}

/** One `a:tc`. `attrs` go on `a:tc`; `tcPr` is the content of `a:tcPr` and its attributes. */
export function cell(text: string, o: { attrs?: string; tcPrAttrs?: string; tcPr?: string; rPr?: string } = {}): string {
  const body = text
    .split('\n')
    .map((t) => (t ? `<a:p><a:r><a:rPr lang="en-US"${o.rPr ?? ''}/><a:t>${escapeText(t)}</a:t></a:r></a:p>` : '<a:p><a:endParaRPr lang="en-US"/></a:p>'))
    .join('');
  return `<a:tc${o.attrs ?? ''}><a:txBody><a:bodyPr/><a:lstStyle/>${body}</a:txBody><a:tcPr${o.tcPrAttrs ?? ''}>${o.tcPr ?? ''}</a:tcPr></a:tc>`;
}

export const solid = (hex: string): string => `<a:solidFill><a:srgbClr val="${hex}"/></a:solidFill>`;
