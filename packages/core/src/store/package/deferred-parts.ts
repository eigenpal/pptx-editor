// Deferred XML parts: a progressive open reads a long deck's slides when they are first needed.
//
// `readPptx` with `defer` keeps some XML parts (the slides and notes slides other than the
// start slide) as their bytes. Such a part is an ordinary `PackagePart` of kind `xml` whose
// `xml` parses the bytes on first read and keeps the tree, so every reader of the package stays
// correct whether or not the part was parsed: a deferred part is a parse that has not happened
// yet, never different content. The editor parses the rest in background slices
// (`editor/progressive-open.ts`), and whole-deck readers (save, a collaboration room) finish it
// first.
//
// The parse is the reader's own (`parseXmlPart`, bounded, entity-free), with the same limits.
// A part that does not parse cannot refuse the deck at the trust boundary any more, because the
// boundary has passed. It reads as a stand-in with an empty shape tree, so no reader fails,
// and its refusal is recorded and reported at once (`onReject`): the editor then blocks every
// commit, so no edit is built on a stand-in. `writePptx` refuses to write a package that holds
// a stand-in, also one that an edit copied (the copy keeps the stand-in's namespace list), so a
// stand-in never leaves the session.

import type { PackagePart, PackageRejection, PptxPackage } from './package.ts';
import { parseXmlPart, type XmlPart } from './tree.ts';

/**
 * What happens to a deferred part's tree right after it parses (the editor drops its marks).
 * @internal
 */
export type DeferredPartFinish = (part: XmlPart) => XmlPart;

/** Called once, where the read happens, when a deferred part does not parse. @internal */
export type DeferredPartReject = (rejection: Extract<PackageRejection, { reason: 'xml' }>) => void;

interface DeferredState {
  bytes: Uint8Array | null;
  xml: XmlPart | null;
  rejection: Extract<PackageRejection, { reason: 'xml' }> | null;
  /** Root attributes read from the start tag, so a second read does not scan again. */
  peeked?: Map<string, string | undefined | null>;
}

const states = new WeakMap<PackagePart, DeferredState>();
/** The namespace lists of stand-in roots: an edit that copies a root keeps its list. */
const standIns = new WeakSet<object>();
let parsedCount = 0;
let rejectionCount = 0;
let standInReadCount = 0;

/**
 * The stand-in tree of a deferred part that did not parse: the root of a slide or a notes slide
 * with an empty shape tree. It is never saved (see the module comment).
 */
function standIn(partName: string, notes: boolean): XmlPart {
  const local = notes ? 'notes' : 'sld';
  const xml = `<p:${local} xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld></p:${local}>`;
  const parsed = parseXmlPart(partName, xml);
  if (!parsed.ok) throw new Error('the deferred stand-in does not parse');
  standIns.add(parsed.part.root.nsDecls);
  return parsed.part;
}

/**
 * Whether `part` is, or was edited from, the stand-in of a deferred part that did not parse:
 * `writePptx` refuses to write it. @internal
 */
export function holdsStandIn(part: PackagePart): boolean {
  if (states.get(part)?.rejection) return true;
  return part.kind === 'xml' && !isDeferredPending(part) && standIns.has(part.xml.root.nsDecls);
}

/**
 * A deferred XML part over `bytes`. `decode` turns the bytes into text, as the reader does;
 * `finish` runs on the tree once it parses; `onReject` hears a part that does not parse;
 * `notes` picks the stand-in's root. @internal
 */
export function deferredXmlPart(
  partName: string,
  bytes: Uint8Array,
  decode: (bytes: Uint8Array) => string,
  options: {
    readonly notes: boolean;
    readonly finish?: DeferredPartFinish;
    readonly onReject?: DeferredPartReject;
  }
): PackagePart {
  const state: DeferredState = { bytes, xml: null, rejection: null };
  const part = {
    kind: 'xml' as const,
    partName,
    get xml(): XmlPart {
      if (state.xml) {
        if (state.rejection) standInReadCount += 1;
        return state.xml;
      }
      const source = state.bytes!;
      state.bytes = null;
      parsedCount += 1;
      const tree = parseXmlPart(partName, decode(source));
      if (tree.ok) state.xml = options.finish ? options.finish(tree.part) : tree.part;
      else {
        state.rejection = { reason: 'xml', partName, detail: tree.reason };
        rejectionCount += 1;
        state.xml = standIn(partName, options.notes);
        standInReadCount += 1;
        options.onReject?.(state.rejection);
      }
      return state.xml;
    },
  };
  states.set(part, state);
  return part;
}

/**
 * How many times, in this page, a reader got the stand-in of a part that did not parse: a
 * reader that compares it before and after its work knows whether it read an empty stand-in
 * in place of a slide or a notes slide. @internal
 */
export function standInReads(): number {
  return standInReadCount;
}

/** Whether `part` is a deferred part whose bytes have not been parsed yet. @internal */
export function isDeferredPending(part: PackagePart | undefined): boolean {
  if (!part) return false;
  const state = states.get(part);
  return !!state && state.xml === null;
}

/** Whether `part` was read as a deferred part (parsed since or not). @internal */
export function isDeferredPart(part: PackagePart | undefined): boolean {
  return !!part && states.has(part);
}

/**
 * The names of the deferred parts of `pkg` that have not been parsed yet, in package order.
 * @internal
 */
export function pendingDeferredParts(pkg: PptxPackage): string[] {
  const out: string[] = [];
  for (const part of pkg.parts.values()) if (isDeferredPending(part)) out.push(part.partName);
  return out;
}

/**
 * Parse the deferred part `partName` of `pkg` now. False when there was nothing to parse.
 * @internal
 */
export function parseDeferredPart(pkg: PptxPackage, partName: string): boolean {
  const part = pkg.parts.get(partName);
  if (!part || !isDeferredPending(part) || part.kind !== 'xml') return false;
  void part.xml;
  return true;
}

/** Parse every deferred part of `pkg` that has not been parsed yet. @internal */
export function parseDeferredParts(pkg: PptxPackage): void {
  for (const part of pkg.parts.values())
    if (isDeferredPending(part) && part.kind === 'xml') void part.xml;
}

/** The refusal of a deferred part of `pkg` that did not parse, or undefined. @internal */
export function deferredRejection(
  pkg: PptxPackage
): Extract<PackageRejection, { reason: 'xml' }> | undefined {
  for (const part of pkg.parts.values()) {
    const rejection = states.get(part)?.rejection;
    if (rejection) return rejection;
  }
  return undefined;
}

/** The refusal of `part` when it is a deferred part that did not parse. @internal */
export function deferredPartRejection(
  part: PackagePart
): Extract<PackageRejection, { reason: 'xml' }> | undefined {
  return states.get(part)?.rejection ?? undefined;
}

/**
 * The unprefixed attribute `name` of the root start tag of a deferred part that has not been
 * parsed, read from its first bytes without parsing the part. Undefined when the attribute is
 * absent; null when the start tag cannot be read this way, and the caller then reads `xml`:
 * a UTF-16 part, a start tag past the first 8 KiB, or a value that needs entity decoding.
 * @internal
 */
export function peekDeferredRootAttribute(
  part: PackagePart,
  name: string
): string | undefined | null {
  const state = states.get(part);
  if (!state || state.xml || !state.bytes) return null;
  const known = state.peeked?.get(name);
  if (known !== undefined || state.peeked?.has(name)) return known;
  const tag = rootStartTag(state.bytes);
  const value = tag === null ? null : rootAttribute(tag, name);
  (state.peeked ??= new Map()).set(name, value);
  return value;
}

/** How many bytes the root start tag is looked for in. */
const PEEK_BYTES = 8192;
const LT = 0x3c;
const GT = 0x3e;
const QUESTION = 0x3f;
const BANG = 0x21;
const DASH = 0x2d;

/**
 * The root start tag of UTF-8 XML `bytes`, from `<` to `>`, decoded; null when it is not in the
 * first `PEEK_BYTES` or the bytes are not UTF-8. Skips the declaration, processing
 * instructions, and comments before it.
 */
function rootStartTag(bytes: Uint8Array): string | null {
  const end = Math.min(bytes.length, PEEK_BYTES);
  let at = 0;
  if (bytes[0] === 0xff || bytes[0] === 0xfe) return null; // UTF-16
  if (bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf) at = 3;
  for (;;) {
    while (at < end && bytes[at] !== LT) at += 1;
    if (at + 1 >= end) return null;
    const next = bytes[at + 1];
    if (next === QUESTION) {
      while (at + 1 < end && !(bytes[at] === QUESTION && bytes[at + 1] === GT)) at += 1;
      at += 2;
      continue;
    }
    if (next === BANG) {
      if (bytes[at + 2] !== DASH || bytes[at + 3] !== DASH) return null; // a DOCTYPE: refused later
      at += 4;
      while (
        at + 2 < end &&
        !(bytes[at] === DASH && bytes[at + 1] === DASH && bytes[at + 2] === GT)
      )
        at += 1;
      at += 3;
      continue;
    }
    break;
  }
  // The tag runs to the first `>` outside a quoted value.
  let quote = 0;
  for (let i = at + 1; i < end; i += 1) {
    const b = bytes[i]!;
    if (quote) {
      if (b === quote) quote = 0;
    } else if (b === 0x22 || b === 0x27) quote = b;
    else if (b === GT) return utf8.decode(bytes.subarray(at, i + 1));
  }
  return null;
}

const utf8 = new TextDecoder('utf-8', { fatal: false });

/** The attribute `name` of start tag `tag`, as `peekDeferredRootAttribute`. */
function rootAttribute(tag: string, name: string): string | undefined | null {
  let at = 1;
  while (at < tag.length && !/[\s/>]/.test(tag[at]!)) at += 1;
  let found: string | undefined;
  for (;;) {
    while (at < tag.length && /\s/.test(tag[at]!)) at += 1;
    if (at >= tag.length) return null;
    if (tag[at] === '>' || tag.startsWith('/>', at)) return found;
    const eq = tag.indexOf('=', at);
    if (eq < 0) return null;
    const attrName = tag.slice(at, eq).trim();
    let q = eq + 1;
    while (q < tag.length && /\s/.test(tag[q]!)) q += 1;
    const quote = tag[q];
    if (quote !== '"' && quote !== "'") return null;
    const close = tag.indexOf(quote, q + 1);
    if (close < 0) return null;
    const value = tag.slice(q + 1, close);
    if (attrName === name) {
      if (value.includes('&') || value.includes('<')) return null;
      found = value;
    }
    at = close + 1;
  }
}

/**
 * How many deferred parts have failed to parse in this page: a change means a new refusal.
 * @internal
 */
export function deferredRejectionCount(): number {
  return rejectionCount;
}

/** The bytes of the deferred parts of `pkg` that have not been parsed yet, uncompressed. @internal */
export function pendingDeferredBytes(pkg: PptxPackage): number {
  let total = 0;
  for (const part of pkg.parts.values()) {
    const state = states.get(part);
    if (state && state.xml === null && state.bytes) total += state.bytes.length;
  }
  return total;
}

/** @internal How many deferred parts have parsed in this page, for tests and the bench. */
export function deferredParseCount(): number {
  return parsedCount;
}
