// Corpus census: counts every PresentationML / DrawingML feature used across a set of decks.
//
// Reports COUNTS AND FEATURE NAMES ONLY. It never prints or writes deck text, file names,
// image bytes, or anything else from the decks. Decks are referred to as deck1…deckN, sorted
// by file name, as AGENTS.md requires.
//
//   bun scripts/fidelity/census.ts                  # private corpus (demo-presentations/)
//   bun scripts/fidelity/census.ts --dir fixtures   # synthetic fixtures
//   bun scripts/fidelity/census.ts --out local/census/corpus-census.md

import { readdirSync, readFileSync, writeFileSync, mkdirSync, existsSync } from 'node:fs';
import { join, dirname } from 'node:path';
import { unzipSync, strFromU8 } from 'fflate';
import { XMLParser } from 'fast-xml-parser';

type Counter = Map<string, number>;
const bump = (c: Counter, key: string, n = 1): void => {
  c.set(key, (c.get(key) ?? 0) + n);
};

interface Census {
  decks: number;
  slides: number;
  layouts: number;
  masters: number;
  slideSizes: Counter;
  elements: Counter;
  presetGeometries: Counter;
  fills: Counter;
  gradientKinds: Counter;
  lines: Counter;
  effects: Counter;
  colorKinds: Counter;
  colorTransforms: Counter;
  bodyProps: Counter;
  autofit: Counter;
  paragraphProps: Counter;
  runProps: Counter;
  bullets: Counter;
  fonts: Counter;
  embeddedFonts: Counter;
  graphicFrames: Counter;
  alternateContent: Counter;
  pictures: Counter;
  placeholders: Counter;
  backgrounds: Counter;
  transforms: Counter;
  contentTypes: Counter;
  perDeck: { name: string; slides: number; embeddedFonts: number; size: string }[];
}

const newCensus = (): Census => ({
  decks: 0,
  slides: 0,
  layouts: 0,
  masters: 0,
  slideSizes: new Map(),
  elements: new Map(),
  presetGeometries: new Map(),
  fills: new Map(),
  gradientKinds: new Map(),
  lines: new Map(),
  effects: new Map(),
  colorKinds: new Map(),
  colorTransforms: new Map(),
  bodyProps: new Map(),
  autofit: new Map(),
  paragraphProps: new Map(),
  runProps: new Map(),
  bullets: new Map(),
  fonts: new Map(),
  embeddedFonts: new Map(),
  graphicFrames: new Map(),
  alternateContent: new Map(),
  pictures: new Map(),
  placeholders: new Map(),
  backgrounds: new Map(),
  transforms: new Map(),
  contentTypes: new Map(),
  perDeck: [],
});

const parser = new XMLParser({
  preserveOrder: true,
  ignoreAttributes: false,
  attributeNamePrefix: '',
  processEntities: false,
  parseTagValue: false,
  parseAttributeValue: false,
  trimValues: false,
});

interface El {
  name: string;
  attrs: Record<string, string>;
  children: El[];
}

function toEls(items: unknown[]): El[] {
  const out: El[] = [];
  for (const item of items as Record<string, unknown>[]) {
    const tag = Object.keys(item).find((k) => k !== ':@' && k !== '#text');
    if (!tag || tag.startsWith('?')) continue;
    out.push({
      name: tag,
      attrs: (item[':@'] as Record<string, string>) ?? {},
      children: toEls((item[tag] as unknown[]) ?? []),
    });
  }
  return out;
}

const COLOR_ELEMENTS = new Set([
  'a:srgbClr',
  'a:schemeClr',
  'a:sysClr',
  'a:prstClr',
  'a:scrgbClr',
  'a:hslClr',
]);
const FILL_ELEMENTS = new Set([
  'a:solidFill',
  'a:gradFill',
  'a:blipFill',
  'p:blipFill',
  'a:pattFill',
  'a:noFill',
  'a:grpFill',
]);
const EFFECTS = new Set([
  'a:outerShdw',
  'a:innerShdw',
  'a:prstShdw',
  'a:glow',
  'a:softEdge',
  'a:reflection',
  'a:blur',
  'a:fillOverlay',
]);
const BLIP_EFFECTS = new Set([
  'a:grayscl',
  'a:duotone',
  'a:alphaModFix',
  'a:lum',
  'a:biLevel',
  'a:clrChange',
  'a:alphaBiLevel',
  'a:tint',
  'a:hsl',
]);

function walk(el: El, c: Census, parent: string | null): void {
  bump(c.elements, el.name);
  const a = el.attrs;
  switch (el.name) {
    case 'a:prstGeom':
      bump(c.presetGeometries, a.prst ?? '(none)');
      break;
    case 'a:custGeom':
      bump(c.presetGeometries, '(custGeom)');
      break;
    case 'a:gradFill': {
      const path = el.children.find((ch) => ch.name === 'a:path');
      bump(c.gradientKinds, path ? `path:${path.attrs.path}` : 'linear');
      if (a.rotWithShape) bump(c.gradientKinds, `rotWithShape=${a.rotWithShape}`);
      break;
    }
    case 'a:ln':
      if (a.w) bump(c.lines, 'w');
      if (a.cap) bump(c.lines, `cap=${a.cap}`);
      if (a.cmpd) bump(c.lines, `cmpd=${a.cmpd}`);
      break;
    case 'a:prstDash':
      bump(c.lines, `dash=${a.val}`);
      break;
    case 'a:custDash':
      bump(c.lines, 'custDash');
      break;
    case 'a:headEnd':
    case 'a:tailEnd':
      if (a.type && a.type !== 'none') bump(c.lines, `${el.name.slice(2)}=${a.type}`);
      break;
    case 'a:round':
    case 'a:bevel':
    case 'a:miter':
      if (parent === 'a:ln') bump(c.lines, `join=${el.name.slice(2)}`);
      break;
    case 'a:bodyPr':
      for (const [k, v] of Object.entries(a)) bump(c.bodyProps, `${k}=${v}`);
      break;
    case 'a:normAutofit':
      bump(c.autofit, 'normAutofit');
      if (a.fontScale) bump(c.autofit, 'normAutofit@fontScale');
      if (a.lnSpcReduction) bump(c.autofit, 'normAutofit@lnSpcReduction');
      break;
    case 'a:spAutoFit':
    case 'a:noAutofit':
      bump(c.autofit, el.name.slice(2));
      break;
    case 'a:pPr':
    case 'a:lvl1pPr':
    case 'a:lvl2pPr':
    case 'a:lvl3pPr':
    case 'a:defPPr':
      for (const k of Object.keys(a)) bump(c.paragraphProps, k);
      for (const ch of el.children)
        if (['a:lnSpc', 'a:spcBef', 'a:spcAft'].includes(ch.name)) {
          const kind = ch.children[0]?.name ?? '?';
          bump(c.paragraphProps, `${ch.name.slice(2)}:${kind.slice(2)}`);
        }
      break;
    case 'a:rPr':
    case 'a:defRPr':
    case 'a:endParaRPr':
      for (const k of Object.keys(a)) bump(c.runProps, k);
      break;
    case 'a:buChar':
    case 'a:buAutoNum':
    case 'a:buNone':
    case 'a:buBlip':
    case 'a:buFont':
    case 'a:buSzPct':
    case 'a:buSzPts':
    case 'a:buClr':
    case 'a:buFontTx':
    case 'a:buClrTx':
    case 'a:buSzTx':
      bump(c.bullets, el.name === 'a:buAutoNum' ? `buAutoNum:${a.type}` : el.name.slice(2));
      break;
    case 'a:latin':
    case 'a:ea':
    case 'a:cs':
    case 'a:sym':
      if (a.typeface) bump(c.fonts, `${el.name.slice(2)}:${a.typeface}`);
      break;
    case 'a:graphicData':
      bump(c.graphicFrames, a.uri ?? '(none)');
      break;
    case 'mc:Choice':
      bump(c.alternateContent, `Requires=${a.Requires}`);
      break;
    case 'p:pic':
      bump(c.pictures, 'p:pic');
      break;
    case 'a:blip':
      bump(c.pictures, a['r:embed'] ? 'blip:embed' : a['r:link'] ? 'blip:link(external)' : 'blip');
      break;
    case 'a:srcRect':
      bump(c.pictures, 'srcRect');
      break;
    case 'a:tile':
      bump(c.pictures, 'tile');
      break;
    case 'a:stretch':
      bump(c.pictures, 'stretch');
      break;
    case 'asvg:svgBlip':
      bump(c.pictures, 'svgBlip');
      break;
    case 'p:ph':
      bump(c.placeholders, a.type ?? 'obj(default)');
      break;
    case 'p:bgPr':
    case 'p:bgRef':
      bump(c.backgrounds, el.name);
      break;
    case 'a:xfrm':
    case 'p:xfrm':
      if (a.rot && a.rot !== '0') bump(c.transforms, 'rot');
      if (a.flipH === '1') bump(c.transforms, 'flipH');
      if (a.flipV === '1') bump(c.transforms, 'flipV');
      break;
  }
  if (FILL_ELEMENTS.has(el.name)) bump(c.fills, el.name.replace('p:', 'a:'));
  if (EFFECTS.has(el.name)) bump(c.effects, el.name.slice(2));
  if (BLIP_EFFECTS.has(el.name) && parent === 'a:blip') bump(c.effects, `blip:${el.name.slice(2)}`);
  if (COLOR_ELEMENTS.has(el.name)) {
    bump(c.colorKinds, el.name === 'a:schemeClr' ? `schemeClr:${a.val}` : el.name.slice(2));
    for (const t of el.children) bump(c.colorTransforms, t.name.slice(2));
  }
  for (const ch of el.children) walk(ch, c, el.name);
}

function censusDeck(bytes: Uint8Array, name: string, c: Census): void {
  const files = unzipSync(bytes);
  c.decks += 1;
  let slides = 0;
  let embedded = 0;
  let size = '?';
  for (const [path, data] of Object.entries(files)) {
    if (path === '[Content_Types].xml') {
      const els = toEls(parser.parse(strFromU8(data)));
      for (const el of els[0]?.children ?? [])
        bump(c.contentTypes, `${el.name}:${el.attrs.ContentType ?? ''}`.slice(0, 120));
    }
    if (path.startsWith('ppt/fonts/')) embedded += 1;
    if (!path.endsWith('.xml') || path.includes('_rels')) continue;
    const isSlide = /^ppt\/slides\/slide\d+\.xml$/.test(path);
    const isLayout = /^ppt\/slideLayouts\/slideLayout\d+\.xml$/.test(path);
    const isMaster = /^ppt\/slideMasters\/slideMaster\d+\.xml$/.test(path);
    const isPresentation = path === 'ppt/presentation.xml';
    const isTheme = /^ppt\/theme\/theme\d+\.xml$/.test(path);
    const isDrawing = /^ppt\/diagrams\/drawing\d+\.xml$/.test(path);
    if (isSlide) slides += 1;
    if (isLayout) c.layouts += 1;
    if (isMaster) c.masters += 1;
    if (!(isSlide || isLayout || isMaster || isPresentation || isTheme || isDrawing)) continue;
    const els = toEls(parser.parse(strFromU8(data)));
    if (isPresentation) {
      const root = els[0];
      const sz = root?.children.find((e) => e.name === 'p:sldSz');
      if (sz) {
        size = `${(Number(sz.attrs.cx) / 914400).toFixed(3)}x${(Number(sz.attrs.cy) / 914400).toFixed(3)}in`;
        bump(c.slideSizes, size);
      }
      const efl = root?.children.find((e) => e.name === 'p:embeddedFontLst');
      for (const ef of efl?.children ?? []) {
        const font = ef.children.find((e) => e.name === 'p:font');
        const faces = ef.children
          .filter((e) => e.name !== 'p:font')
          .map((e) => e.name.slice(2))
          .join('+');
        bump(c.embeddedFonts, `${font?.attrs.typeface ?? '?'} [${faces}]`);
      }
    }
    if (isSlide || isLayout || isMaster || isDrawing) for (const el of els) walk(el, c, null);
  }
  c.slides += slides;
  c.perDeck.push({ name, slides, embeddedFonts: embedded, size });
}

function table(title: string, counter: Counter, limit = 60): string {
  const rows = [...counter.entries()].sort((a, b) => b[1] - a[1]).slice(0, limit);
  if (rows.length === 0) return `### ${title}\n\nNone.\n`;
  const body = rows
    .map(([k, v]) => `| \`${k.replace(/\\/g, '\\\\').replace(/\|/g, '\\|')}\` | ${v} |`)
    .join('\n');
  return `### ${title}\n\n| Feature | Count |\n| --- | ---: |\n${body}\n`;
}

/**
 * The census as Markdown. `perDeck` adds one row per deck; it stays off for the private
 * corpus, because a deck's slide count, size, and fonts together can identify a client.
 */
export function renderCensus(c: Census, label: string, perDeck = true): string {
  const parts = [
    `# ${label} census\n`,
    `Generated by \`scripts/fidelity/census.ts\`. Counts and feature names only${perDeck ? '. Decks are named by sorted position' : ', totals across all decks'}.\n`,
    `- Decks: ${c.decks}\n- Slides: ${c.slides}\n- Slide layouts: ${c.layouts}\n- Slide masters: ${c.masters}\n`,
    perDeck
      ? `## Per deck\n\n| Deck | Slides | Embedded font parts | Slide size |\n| --- | ---: | ---: | --- |\n` +
        c.perDeck
          .map((d) => `| ${d.name} | ${d.slides} | ${d.embeddedFonts} | ${d.size} |`)
          .join('\n') +
        '\n'
      : '',
    table('Slide sizes', c.slideSizes),
    table('Preset geometries', c.presetGeometries),
    table('Fills', c.fills),
    table('Gradient kinds', c.gradientKinds),
    table('Lines', c.lines),
    table('Effects', c.effects),
    table('Color kinds', c.colorKinds),
    table('Color transforms', c.colorTransforms),
    table('Text body properties', c.bodyProps),
    table('Autofit', c.autofit),
    table('Paragraph properties', c.paragraphProps),
    table('Run properties', c.runProps),
    table('Bullets', c.bullets),
    table('Fonts referenced', c.fonts),
    table('Embedded fonts', c.embeddedFonts),
    table('Graphic frames (charts, tables, diagrams, OLE)', c.graphicFrames),
    table('mc:AlternateContent', c.alternateContent),
    table('Pictures', c.pictures),
    table('Placeholders', c.placeholders),
    table('Backgrounds', c.backgrounds),
    table('Transforms', c.transforms),
    table('Elements (top 80)', c.elements, 80),
    table('Content types', c.contentTypes),
  ];
  return parts.filter(Boolean).join('\n');
}

function main(): void {
  const args = process.argv.slice(2);
  const dir = args.includes('--dir') ? args[args.indexOf('--dir') + 1] : 'demo-presentations';
  const out = args.includes('--out')
    ? args[args.indexOf('--out') + 1]
    : dir === 'demo-presentations'
      ? 'local/census/corpus-census.md'
      : `local/census/${dir.replace(/\W+/g, '-')}-census.md`;
  if (!existsSync(dir)) {
    console.log(`No decks at ${dir}. Nothing to count.`);
    return;
  }
  const files = readdirSync(dir)
    .filter((f) => f.toLowerCase().endsWith('.pptx'))
    .sort();
  const c = newCensus();
  files.forEach((f, i) => censusDeck(readFileSync(join(dir, f)), `deck${i + 1}`, c));
  const label = dir === 'demo-presentations' ? 'Private corpus' : `Fixture set (${dir})`;
  mkdirSync(dirname(out), { recursive: true });
  writeFileSync(out, renderCensus(c, label, dir !== 'demo-presentations'));
  console.log(`${c.decks} decks, ${c.slides} slides. Wrote ${out}.`);
}

if (import.meta.main) main();
