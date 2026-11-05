// Validation of op payloads. Ops are plain JSON from a command, a clipboard paste, or a
// collaboration peer, so nothing here trusts the TypeScript types: every field is checked at run
// time, and every failure is a typed rejection, never a TypeError.
//
// Names must be XML NCNames, namespace URIs and declarations valid XML text, and every binding
// one the Namespaces spec allows. Without this a local name such as `x="1"><p:evil/><y z` is
// written verbatim, and a pasted `xmlns:a` attribute rebinds DrawingML in the saved file.

import { isValidXmlText } from '@eigenpal/ooxml-core';
import { isNcName, namespaceBindingError, qualifiedNameError } from '../package/xml-names.ts';
import type { NameRuleViolation } from '../package/xml-names.ts';
import type { OpRejection } from './ops.ts';

type Rejection = OpRejection;

const isObject = (v: unknown): v is Record<string, unknown> =>
  typeof v === 'object' && v !== null && !Array.isArray(v);

const malformed = (detail: string): Rejection => ({ code: 'malformed-op', detail });

const fromRule = (v: NameRuleViolation, name: string, uri: string): Rejection =>
  v.rule === 'reserved'
    ? { code: 'reserved-name', name, detail: v.detail }
    : { code: 'invalid-namespace', uri, detail: v.detail };

function checkUri(uri: unknown, what: string): Rejection | null {
  if (typeof uri !== 'string') return malformed(`${what} namespace is not a string`);
  if (!isValidXmlText(uri))
    return { code: 'invalid-namespace', uri, detail: 'not valid XML 1.0 text' };
  return null;
}

/**
 * Check an expanded name `{ns, local, prefix?}` for an element or an attribute.
 * `prefix` may be absent or empty (no preference).
 */
export function checkQualifiedName(name: unknown, isAttribute: boolean): Rejection | null {
  const what = isAttribute ? 'attribute' : 'element';
  if (!isObject(name)) return malformed(`${what} name is not an object`);
  const { ns, local, prefix } = name;
  const bad = checkUri(ns, what);
  if (bad) return bad;
  if (typeof local !== 'string') return malformed(`${what} local name is not a string`);
  if (!isNcName(local)) return { code: 'invalid-name', name: local, detail: `${what} local name` };
  if (prefix !== undefined) {
    if (typeof prefix !== 'string') return malformed(`${what} prefix is not a string`);
    if (prefix !== '' && !isNcName(prefix))
      return { code: 'invalid-name', name: prefix, detail: `${what} prefix` };
  }
  const rule = qualifiedNameError(ns as string, local, prefix as string | undefined, isAttribute);
  return rule ? fromRule(rule, prefix ? `${prefix}:${local}` : local, ns as string) : null;
}

function checkNsDecls(decls: unknown, elementNs: string): Rejection | null {
  if (decls === undefined) return null;
  if (!Array.isArray(decls)) return malformed('nsDecls is not an array');
  const prefixes = new Set<string>();
  for (const d of decls) {
    if (!isObject(d)) return malformed('a namespace declaration is not an object');
    const { prefix, uri } = d;
    if (typeof prefix !== 'string') return malformed('a declared prefix is not a string');
    if (prefix !== '' && !isNcName(prefix))
      return { code: 'invalid-name', name: prefix, detail: 'declared prefix' };
    const bad = checkUri(uri, 'declared');
    if (bad) return bad;
    const rule = namespaceBindingError(prefix, uri as string);
    if (rule) return fromRule(rule, prefix ? `xmlns:${prefix}` : 'xmlns', uri as string);
    if (prefixes.has(prefix)) return { code: 'duplicate-name', name: prefix || 'xmlns' };
    prefixes.add(prefix);
    // An element in no namespace cannot be written under its own non-empty default namespace.
    if (prefix === '' && uri !== '' && elementNs === '')
      return { code: 'invalid-namespace', uri: uri as string, detail: 'element has no namespace' };
  }
  return null;
}

function checkAttrs(attrs: unknown): Rejection | null {
  if (attrs === undefined) return null;
  if (!Array.isArray(attrs)) return malformed('attrs is not an array');
  const seen = new Set<string>();
  for (const a of attrs) {
    const bad = checkQualifiedName(a, true);
    if (bad) return bad;
    const { ns, local, value } = a as Record<string, unknown>;
    if (typeof value !== 'string') return malformed('attribute value is not a string');
    if (!isValidXmlText(value)) return { code: 'invalid-text', detail: 'attribute value' };
    const key = `{${ns as string}}${local as string}`;
    if (seen.has(key)) return { code: 'duplicate-name', name: local as string };
    seen.add(key);
  }
  return null;
}

/**
 * Check a node spec from untrusted JSON. Returns the height of the element tree it describes
 * (a leaf element is 1), or the rejection. `maxHeight` is how many levels the parent leaves
 * below it, so an op cannot build a tree deeper than the reader accepts.
 */
export function checkSpec(spec: unknown, maxHeight: number): Rejection | number {
  if (maxHeight < 1) return { code: 'too-deep', detail: 'the parent is at the depth limit' };
  const bad = checkQualifiedName(spec, false);
  if (bad) return bad;
  const s = spec as Record<string, unknown>;
  const declBad = checkNsDecls(s.nsDecls, s.ns as string);
  if (declBad) return declBad;
  const attrBad = checkAttrs(s.attrs);
  if (attrBad) return attrBad;
  if (s.children === undefined) return 1;
  if (!Array.isArray(s.children)) return malformed('children is not an array');
  let height = 1;
  for (const child of s.children) {
    if (isObject(child) && 'text' in child) {
      if (typeof child.text !== 'string') return malformed('text is not a string');
      if (!isValidXmlText(child.text)) return { code: 'invalid-text', detail: 'text' };
      continue;
    }
    if (maxHeight < 2)
      return { code: 'too-deep', detail: 'the node would be deeper than the reader accepts' };
    const h = checkSpec(child, maxHeight - 1);
    if (typeof h !== 'number') return h;
    if (h + 1 > height) height = h + 1;
  }
  return height;
}

const OP_FIELDS: Readonly<Record<string, readonly string[]>> = Object.freeze({
  setAttribute: ['part', 'node'],
  setText: ['part', 'node'],
  insertNode: ['part', 'parent'],
  removeNode: ['part', 'node'],
  moveNode: ['part', 'node', 'parent'],
  addPart: ['part'],
  removePart: ['part'],
});

/** Check the shape of an op: a known kind with string addresses and the right field types. */
export function checkOpShape(op: unknown): Rejection | null {
  if (!isObject(op)) return malformed('op is not an object');
  const fields = typeof op.op === 'string' ? OP_FIELDS[op.op] : undefined;
  if (!fields || !Object.hasOwn(OP_FIELDS, op.op as string))
    return malformed(`unknown op ${String(op.op)}`);
  for (const f of fields) if (typeof op[f] !== 'string') return malformed(`${f} is not a string`);
  switch (op.op) {
    case 'setAttribute':
      if (op.value !== null && typeof op.value !== 'string')
        return malformed('value is not a string or null');
      return checkQualifiedName(op.name, true);
    case 'setText':
      return typeof op.text === 'string' ? null : malformed('text is not a string');
    case 'insertNode':
    case 'moveNode':
      return typeof op.index === 'number' ? null : malformed('index is not a number');
    case 'addPart':
      if (op.xml !== undefined && typeof op.xml !== 'string')
        return malformed('xml is not a string');
      if (op.base64 !== undefined && typeof op.base64 !== 'string')
        return malformed('base64 is not a string');
      return null;
  }
  return null;
}
