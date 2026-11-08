// Generate the preset shape table from ECMA-376 Part 1 `presetShapeDefinitions.xml`, and the
// preset text warp table from `presetTextWarpDefinitions.xml` (20.1.9.19 `prstTxWarp`).
//
// The definitions are the spec's own guide formulas and paths. The engine evaluates them; it
// never hand-draws an approximation. Re-run after updating the reference files:
//
//   bun scripts/geometry/build-presets.ts
//
// Output: packages/ooxml-core/src/drawingml-preset-geometry.generated.ts and
// packages/ooxml-core/src/drawingml-preset-text-warp.generated.ts

import { readFileSync, writeFileSync } from 'node:fs';
import { XMLParser } from 'fast-xml-parser';

const GEOMETRIES = 'reference/ecma-376/part1/drawingml-geometries';

const parser = new XMLParser({
  preserveOrder: true,
  ignoreAttributes: false,
  attributeNamePrefix: '',
  parseAttributeValue: false,
  parseTagValue: false,
});

type Raw = Record<string, unknown>;
interface El {
  name: string;
  attrs: Record<string, string>;
  children: El[];
}

function toEls(items: unknown[]): El[] {
  const out: El[] = [];
  for (const item of items as Raw[]) {
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

/** Every definition of a definitions file, keyed by its element name. */
function readDefinitions(source: string): Record<string, unknown> {
  const xml = readFileSync(source, 'utf8').replace(/^﻿/, '');
  const root = toEls(parser.parse(xml))[0];
  const shapes: Record<string, unknown> = {};
  for (const shape of root.children) {
    const find = (n: string): El | undefined => shape.children.find((c) => c.name === n);
    const guides = (n: string): [string, string][] =>
      (find(n)?.children ?? [])
        .filter((g) => g.name === 'gd')
        .map((g) => [g.attrs.name, g.attrs.fmla]);
    const rect = find('rect');
    // Adjust handles (ECMA-376 Part 1, 20.1.9.1 ahLst): each guide a handle moves, with its range.
    const handles: [string, string, string | null, string | null][] = [];
    // Where each handle sits (`a:pos`), with the guides it moves: x and y, or radius and angle.
    const places: [string, string, string, string | null, string | null][] = [];
    for (const h of find('ahLst')?.children ?? []) {
      const pos = h.children.find((c) => c.name === 'pos');
      const polar = h.name === 'ahPolar';
      if (pos)
        places.push([
          polar ? 'polar' : 'xy',
          pos.attrs.x,
          pos.attrs.y,
          (polar ? h.attrs.gdRefR : h.attrs.gdRefX) ?? null,
          (polar ? h.attrs.gdRefAng : h.attrs.gdRefY) ?? null,
        ]);
      const axes: [string, string, string, string][] =
        h.name === 'ahPolar'
          ? [
              ['gdRefR', 'r', 'minR', 'maxR'],
              ['gdRefAng', 'ang', 'minAng', 'maxAng'],
            ]
          : [
              ['gdRefX', 'x', 'minX', 'maxX'],
              ['gdRefY', 'y', 'minY', 'maxY'],
            ];
      for (const [ref, axis, min, max] of axes)
        if (h.attrs[ref])
          handles.push([h.attrs[ref], axis, h.attrs[min] ?? null, h.attrs[max] ?? null]);
    }
    // Connection sites (ECMA-376 Part 1, 20.1.9.10 cxnLst): the angle, then the position.
    const sites: [string, string, string][] = [];
    for (const c of find('cxnLst')?.children ?? []) {
      const pos = c.children.find((x) => x.name === 'pos');
      if (pos) sites.push([c.attrs.ang, pos.attrs.x, pos.attrs.y]);
    }
    const paths = (find('pathLst')?.children ?? []).map((p) => {
      const cmds = p.children.map((c) => {
        const pts = c.children.filter((x) => x.name === 'pt').map((pt) => [pt.attrs.x, pt.attrs.y]);
        switch (c.name) {
          case 'moveTo':
            return ['M', ...pts[0]];
          case 'lnTo':
            return ['L', ...pts[0]];
          case 'arcTo':
            return ['A', c.attrs.wR, c.attrs.hR, c.attrs.stAng, c.attrs.swAng];
          case 'quadBezTo':
            return ['Q', ...pts[0], ...pts[1]];
          case 'cubicBezTo':
            return ['C', ...pts[0], ...pts[1], ...pts[2]];
          case 'close':
            return ['Z'];
        }
        throw new Error(`unknown path command ${c.name}`);
      });
      const attrs: Record<string, string | number | boolean> = {};
      if (p.attrs.w) attrs.w = Number(p.attrs.w);
      if (p.attrs.h) attrs.h = Number(p.attrs.h);
      if (p.attrs.fill) attrs.fill = p.attrs.fill;
      if (p.attrs.stroke === 'false' || p.attrs.stroke === '0') attrs.stroke = false;
      return { ...attrs, cmds };
    });
    shapes[shape.name] = {
      av: guides('avLst'),
      gd: guides('gdLst'),
      ...(rect ? { rect: [rect.attrs.l, rect.attrs.t, rect.attrs.r, rect.attrs.b] } : {}),
      paths,
      ...(handles.length > 0 ? { ah: handles } : {}),
      ...(places.length > 0 ? { hp: places } : {}),
      ...(sites.length > 0 ? { cxn: sites } : {}),
    };
  }
  return shapes;
}

// The Ecma notice must travel with every copy of the data, also inside a consumer's app bundle.
// A legal comment (`/*!`) survives bundling and minification, so the notice is written as one.
const notice = readFileSync('reference/ecma-376/NOTICE', 'utf8');
const ecmaNotice = notice.slice(notice.indexOf('COPYRIGHT NOTICE')).trim();
/** `text` wrapped at `width` columns, each line prefixed for a block comment. */
function commentLines(text: string, width = 96): string {
  const out: string[] = [];
  for (const paragraph of text.split('\n')) {
    let line = '';
    for (const word of paragraph.split(/\s+/).filter(Boolean)) {
      if (line && line.length + word.length + 1 > width) {
        out.push(line);
        line = word;
      } else line = line ? `${line} ${word}` : word;
    }
    out.push(line);
  }
  return out.map((line) => (line ? ` * ${line}` : ' *')).join('\n');
}

/** Write one generated table: `file` is the definitions file, `what` names its content. */
function writeTable(o: {
  file: string;
  out: string;
  what: string;
  doc: string;
  constName: string;
}): void {
  const defs = readDefinitions(`${GEOMETRIES}/${o.file}`);
  const header = `// GENERATED by scripts/geometry/build-presets.ts from ECMA-376 Part 1, 5th edition (2016),
// ${o.file}. Do not edit by hand.

import type { PresetShapeDefinition } from './drawingml-preset-geometry-types.ts';

/*!
${commentLines(
  `ECMA-376 ${o.what}. The guide formulas and paths below are derived from ` +
    `${o.file} in ECMA-376 Part 1, 5th edition (2016), (c) Ecma International. ` +
    'They are a derivative work under item (iv) of the Ecma copyright notice, and the Ecma ' +
    'terms below apply to them, not the Apache License of @pptx-editor.dev/core.'
)}
 *
${commentLines(ecmaNotice)}
 */
/** ${o.doc} */
export const ${o.constName}: Readonly<Record<string, PresetShapeDefinition>> = `;
  writeFileSync(
    o.out,
    `${header}${JSON.stringify(defs)} as unknown as Record<string, PresetShapeDefinition>;\n`
  );
  console.log(`${Object.keys(defs).length} definitions written to ${o.out}`);
}

writeTable({
  file: 'presetShapeDefinitions.xml',
  out: 'packages/ooxml-core/src/drawingml-preset-geometry.generated.ts',
  what: 'preset shape geometry',
  doc: 'Every ECMA-376 preset shape, keyed by `a:prstGeom/@prst`.',
  constName: 'PRESET_SHAPES',
});
writeTable({
  file: 'presetTextWarpDefinitions.xml',
  out: 'packages/ooxml-core/src/drawingml-preset-text-warp.generated.ts',
  what: 'preset text warp geometry',
  doc: 'Every ECMA-376 preset text warp, keyed by `a:prstTxWarp/@prst`.',
  constName: 'PRESET_TEXT_WARPS',
});
