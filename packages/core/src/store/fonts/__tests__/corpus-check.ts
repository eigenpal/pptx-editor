// Corpus check for the EOT/MTX decoder (not a unit test; run with `bun <this file>`).
// Reads every ppt/fonts/*.fntdata part from the private demo decks *in memory*, decodes it,
// validates the sfnt structure, walks every glyph, and checks that 'A'-'Z' map to glyphs
// with outlines. Prints only font names, counts and pass/fail. Never writes font bytes.
// Decks are read with the store's bounded zip reader, never an unbounded one.

import { existsSync, readdirSync, readFileSync } from 'node:fs';
import { join, resolve } from 'node:path';
import { readZip } from '@eigenpal/ooxml-core';
import { decodeEmbeddedFont, validateSfnt } from '../index.ts';

const CORPUS = resolve(import.meta.dir, '../../../../../../demo-presentations');

interface ZipEntry {
  name: string;
  read: () => Uint8Array;
}

/** The deck's entries through the store's bounded zip reader (entry, size, and ratio caps). */
function listZip(buf: Uint8Array): ZipEntry[] {
  const r = readZip(buf);
  if (!r.ok) throw new Error(`zip refused: ${r.reason}`);
  // Part names start with `/`; the checks below match zip entry names.
  return [...r.entries].map(([name, bytes]) => ({
    name: name.replace(/^\//, ''),
    read: () => bytes,
  }));
}

// --- minimal sfnt inspection ------------------------------------------------------------

const u16 = (d: Uint8Array, p: number): number => (d[p] << 8) | d[p + 1];
const i16 = (d: Uint8Array, p: number): number => (u16(d, p) << 16) >> 16;
const u32 = (d: Uint8Array, p: number): number =>
  ((d[p] << 24) >>> 0) + ((d[p + 1] << 16) | (d[p + 2] << 8) | d[p + 3]);

function tables(font: Uint8Array): Map<string, Uint8Array> {
  const out = new Map<string, Uint8Array>();
  const n = u16(font, 4);
  for (let i = 0; i < n; i++) {
    const r = 12 + i * 16;
    const tag = String.fromCharCode(font[r], font[r + 1], font[r + 2], font[r + 3]);
    const off = u32(font, r + 8);
    out.set(tag, font.subarray(off, off + u32(font, r + 12)));
  }
  return out;
}

function cmapLookup(cmap: Uint8Array): (cp: number) => number {
  const n = u16(cmap, 2);
  let best = -1;
  let bestScore = -1;
  for (let i = 0; i < n; i++) {
    const pid = u16(cmap, 4 + i * 8);
    const eid = u16(cmap, 6 + i * 8);
    const off = u32(cmap, 8 + i * 8);
    const fmt = u16(cmap, off);
    const score =
      (pid === 3 && eid === 10 && fmt === 12 ? 4 : 0) ||
      (pid === 0 && fmt === 12 ? 3 : 0) ||
      (pid === 3 && eid === 1 && fmt === 4 ? 2 : 0) ||
      (pid === 0 && fmt === 4 ? 1 : 0);
    if (score > bestScore) {
      bestScore = score;
      best = off;
    }
  }
  if (best < 0 || bestScore <= 0) return () => 0;
  const fmt = u16(cmap, best);
  if (fmt === 12) {
    const groups = u32(cmap, best + 12);
    return (cp) => {
      for (let g = 0; g < groups; g++) {
        const p = best + 16 + g * 12;
        const start = u32(cmap, p);
        const end = u32(cmap, p + 4);
        if (cp >= start && cp <= end) return u32(cmap, p + 8) + cp - start;
      }
      return 0;
    };
  }
  const segX2 = u16(cmap, best + 6);
  const ends = best + 14;
  const starts = ends + segX2 + 2;
  const deltas = starts + segX2;
  const rangeOffsets = deltas + segX2;
  return (cp) => {
    for (let s = 0; s < segX2 / 2; s++) {
      if (cp > u16(cmap, ends + s * 2)) continue;
      const start = u16(cmap, starts + s * 2);
      if (cp < start) return 0;
      const delta = u16(cmap, deltas + s * 2);
      const ro = u16(cmap, rangeOffsets + s * 2);
      if (ro === 0) return (cp + delta) & 0xffff;
      const gp = rangeOffsets + s * 2 + ro + (cp - start) * 2;
      const g = u16(cmap, gp);
      return g === 0 ? 0 : (g + delta) & 0xffff;
    }
    return 0;
  };
}

// --- CFF (OpenType 'CFF ' table, Adobe TN#5176 / TN#5177) -------------------------------

interface CffIndex {
  count: number;
  item: (i: number) => Uint8Array;
  end: number;
}

function readIndex(d: Uint8Array, p: number): CffIndex {
  const count = u16(d, p);
  if (count === 0) return { count, item: () => new Uint8Array(0), end: p + 2 };
  const offSize = d[p + 2];
  const off = (i: number): number => {
    let v = 0;
    for (let k = 0; k < offSize; k++) v = v * 256 + d[p + 3 + i * offSize + k];
    return v;
  };
  const dataStart = p + 3 + (count + 1) * offSize - 1;
  return {
    count,
    item: (i) => d.subarray(dataStart + off(i), dataStart + off(i + 1)),
    end: dataStart + off(count),
  };
}

/** Returns operand list for a DICT operator (only integer operands are needed here). */
function dictLookup(dict: Uint8Array, op: number): number[] | null {
  const stack: number[] = [];
  let p = 0;
  while (p < dict.length) {
    const b = dict[p];
    if (b <= 21) {
      let key = b;
      p++;
      if (b === 12) key = 1200 + dict[p++];
      if (key === op) return stack;
      stack.length = 0;
    } else if (b === 28) {
      stack.push(i16(dict, p + 1));
      p += 3;
    } else if (b === 29) {
      stack.push(u32(dict, p + 1) | 0);
      p += 5;
    } else if (b === 30) {
      p++;
      while (p < dict.length && (dict[p] & 0x0f) !== 0x0f && (dict[p] & 0xf0) !== 0xf0) p++;
      p++;
      stack.push(0);
    } else if (b >= 32 && b <= 246) {
      stack.push(b - 139);
      p++;
    } else if (b >= 247 && b <= 250) {
      stack.push((b - 247) * 256 + dict[p + 1] + 108);
      p += 2;
    } else if (b >= 251 && b <= 254) {
      stack.push(-(b - 251) * 256 - dict[p + 1] - 108);
      p += 2;
    } else p++;
  }
  return null;
}

/** True when a Type 2 charstring draws something (a moveto, or a subroutine call). */
function charstringHasOutline(cs: Uint8Array): boolean {
  let p = 0;
  let args = 0;
  let stems = 0;
  while (p < cs.length) {
    const b = cs[p];
    if (b === 28) {
      args++;
      p += 3;
    } else if (b >= 32) {
      args++;
      p += b <= 246 ? 1 : b === 255 ? 5 : 2;
    } else {
      p++;
      if (b === 21 || b === 22 || b === 4 || b === 10 || b === 29) return true;
      if (b === 1 || b === 3 || b === 18 || b === 23) stems += args >> 1;
      if (b === 19 || b === 20) {
        stems += args >> 1;
        p += (stems + 7) >> 3;
      }
      if (b === 12) p++;
      if (b === 14) return false;
      args = 0;
    }
  }
  return false;
}

function cffCharStrings(cff: Uint8Array): CffIndex {
  const nameIndex = readIndex(cff, cff[2]);
  const topIndex = readIndex(cff, nameIndex.end);
  const charStringsOffset = dictLookup(topIndex.item(0), 17);
  if (!charStringsOffset || charStringsOffset.length === 0) throw new Error('no CharStrings');
  return readIndex(cff, charStringsOffset[charStringsOffset.length - 1]);
}

interface GlyphStats {
  empty: number;
  simple: number;
  composite: number;
  errors: number;
  bboxMismatch: number;
}

/** Walks one TrueType glyph and checks its internal structure fits its loca slot. */
function checkGlyph(g: Uint8Array, numGlyphs: number, stats: GlyphStats): boolean {
  if (g.length === 0) {
    stats.empty++;
    return true;
  }
  if (g.length < 10) return false;
  const nc = i16(g, 0);
  if (nc < 0) {
    stats.composite++;
    let p = 10;
    let flags: number;
    let instr = false;
    do {
      if (p + 4 > g.length) return false;
      flags = u16(g, p);
      if (u16(g, p + 2) >= numGlyphs) return false;
      p += 4 + (flags & 1 ? 4 : 2);
      if (flags & 0x8) p += 2;
      else if (flags & 0x40) p += 4;
      else if (flags & 0x80) p += 8;
      if (flags & 0x100) instr = true;
    } while (flags & 0x20);
    if (instr) p += 2 + u16(g, p);
    return p <= g.length;
  }
  stats.simple++;
  let p = 10 + nc * 2;
  if (p + 2 > g.length) return false;
  const numPoints = nc ? u16(g, p - 2) + 1 : 0;
  p += 2 + u16(g, p);
  const flags: number[] = [];
  while (flags.length < numPoints) {
    if (p >= g.length) return false;
    const f = g[p++];
    flags.push(f);
    if (f & 8) {
      const rep = g[p++];
      for (let k = 0; k < rep; k++) flags.push(f);
    }
  }
  if (flags.length !== numPoints) return false;
  let x = 0;
  let y = 0;
  let xMin = Infinity;
  let xMax = -Infinity;
  let yMin = Infinity;
  let yMax = -Infinity;
  const xs: number[] = [];
  for (const f of flags) {
    if (f & 2) x += f & 0x10 ? g[p++] : -g[p++];
    else if (!(f & 0x10)) {
      x += i16(g, p);
      p += 2;
    }
    xs.push(x);
  }
  for (let i = 0; i < numPoints; i++) {
    const f = flags[i];
    if (f & 4) y += f & 0x20 ? g[p++] : -g[p++];
    else if (!(f & 0x20)) {
      y += i16(g, p);
      p += 2;
    }
    xMin = Math.min(xMin, xs[i]);
    xMax = Math.max(xMax, xs[i]);
    yMin = Math.min(yMin, y);
    yMax = Math.max(yMax, y);
  }
  if (p > g.length) return false;
  if (
    numPoints &&
    (xMin !== i16(g, 2) || yMin !== i16(g, 4) || xMax !== i16(g, 6) || yMax !== i16(g, 8))
  ) {
    stats.bboxMismatch++;
  }
  return true;
}

function main(): void {
  if (!existsSync(CORPUS)) {
    console.log(`corpus not found at ${CORPUS}; skipping`);
    return;
  }
  const decks = readdirSync(CORPUS)
    .filter((f) => f.toLowerCase().endsWith('.pptx'))
    .sort();
  let total = 0;
  let passed = 0;
  const seen = new Map<string, string>();
  decks.forEach((deck, di) => {
    const entries = listZip(new Uint8Array(readFileSync(join(CORPUS, deck))));
    const fonts = entries.filter((e) => /^ppt\/fonts\/.*\.fntdata$/i.test(e.name));
    console.log(`deck #${di + 1}: ${fonts.length} embedded font part(s)`);
    for (const e of fonts) {
      total++;
      const label = e.name.replace(/^ppt\/fonts\//, '');
      const res = decodeEmbeddedFont(e.read());
      if (!res.ok) {
        console.log(`  FAIL ${label}: ${res.reason} (${res.detail})`);
        continue;
      }
      const { info, sfnt } = res;
      const name = `${info.familyName} / ${info.styleName}`;
      const problems: string[] = [];
      const v = validateSfnt(sfnt);
      if (!v.ok) problems.push(...v.errors);
      const t = tables(sfnt);
      const head = t.get('head');
      const loca = t.get('loca');
      const glyf = t.get('glyf');
      const cmap = t.get('cmap');
      const cff = t.get('CFF ');
      const stats: GlyphStats = { empty: 0, simple: 0, composite: 0, errors: 0, bboxMismatch: 0 };
      let az = 0;
      let azMapped = 0;
      let flavor = 'TrueType';
      if (cff && cmap) {
        flavor = 'CFF';
        try {
          const cs = cffCharStrings(cff);
          if (cs.count !== v.numGlyphs) problems.push(`CFF has ${cs.count} charstrings`);
          for (let g = 0; g < cs.count; g++) {
            if (charstringHasOutline(cs.item(g))) stats.simple++;
            else stats.empty++;
          }
          const lookup = cmapLookup(cmap);
          for (let cp = 0x41; cp <= 0x5a; cp++) {
            const gid = lookup(cp);
            if (gid === 0 || gid >= cs.count) continue;
            azMapped++;
            if (charstringHasOutline(cs.item(gid))) az++;
          }
        } catch (err) {
          problems.push(`CFF parse failed: ${(err as Error).message}`);
        }
      } else if (head && loca && glyf && cmap) {
        const long = u16(head, 50) === 1;
        const off = (g: number): number => (long ? u32(loca, g * 4) : u16(loca, g * 2) * 2);
        for (let g = 0; g < v.numGlyphs; g++) {
          if (!checkGlyph(glyf.subarray(off(g), off(g + 1)), v.numGlyphs, stats)) stats.errors++;
        }
        const lookup = cmapLookup(cmap);
        for (let cp = 0x41; cp <= 0x5a; cp++) {
          const gid = lookup(cp);
          if (gid === 0 || gid >= v.numGlyphs) continue;
          azMapped++;
          const len = off(gid + 1) - off(gid);
          if (len >= 10 && i16(glyf, off(gid)) !== 0) az++;
        }
      } else {
        problems.push('missing cmap and glyf/loca or CFF');
      }
      if (stats.errors) problems.push(`${stats.errors} malformed glyph(s)`);
      if (azMapped !== az) problems.push(`A-Z: ${azMapped - az} mapped glyph(s) without outline`);
      if (!info.subset && azMapped !== 26) problems.push(`A-Z: only ${azMapped}/26 mapped`);
      if (azMapped === 0) problems.push('A-Z: none mapped');
      const ok = problems.length === 0;
      if (ok) passed++;
      const glyphSummary =
        flavor === 'CFF'
          ? `${stats.simple} with outlines, ${stats.empty} empty`
          : `${stats.simple} simple, ${stats.composite} composite, ${stats.empty} empty`;
      const summary =
        `${flavor}, ${info.numGlyphs} glyphs (${glyphSummary}), ` +
        `A-Z outlines ${az}/26${info.subset ? ' [subset]' : ''}` +
        (stats.bboxMismatch ? `, ${stats.bboxMismatch} stored-bbox differs` : '') +
        (info.warnings.length ? `, warnings: ${info.warnings.join('; ')}` : '');
      console.log(`  ${ok ? 'PASS' : 'FAIL'} ${label} [${name}] ${summary}`);
      if (!ok) console.log(`       ${problems.join('; ')}`);
      seen.set(name, ok ? 'pass' : 'fail');
    }
  });
  console.log(`\n${passed}/${total} font parts decoded and valid (${seen.size} distinct faces)`);
  if (passed !== total) process.exitCode = 1;
}

main();
