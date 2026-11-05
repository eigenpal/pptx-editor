// Serialize a canonical tree back to XML, and fingerprint it.
//
// Modeled XML parts re-emit NORMALIZED (the same contract as DOCX's ooxml-serialize.ts):
// authored prefixes and namespace declarations are kept, insignificant whitespace is not.
// Every string is validated and escaped on the way out; a value that is not valid XML 1.0
// text fails the save rather than writing a corrupt part. Names and namespace bindings are
// validated too: a name that is not an NCName, or a binding the Namespaces spec forbids, fails
// the save rather than writing markup a reader would parse differently.

import { escapeXmlAttributeChecked, escapeXmlChecked } from '@eigenpal/ooxml-core';
import { NS, conventionalPrefix } from './namespaces.ts';
import { sha256FontBytes } from '@eigenpal/ooxml-core';
import { isNcName, namespaceBindingError, qualifiedNameError } from './xml-names.ts';
import type { Element, Node, XmlPart } from './tree.ts';

const DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';

/** Serialize one part to an XML string with the standard OOXML declaration. */
export function serializeXmlPart(part: XmlPart): string {
  const out: string[] = [DECLARATION];
  emit(part.root, new Map([['xml', NS.xml]]), out);
  return out.join('');
}

function fail(what: string, why: string): never {
  throw new Error(`cannot serialize ${what}: ${why}`);
}

function prefixFor(
  ns: string,
  authored: string | undefined,
  scope: ReadonlyMap<string, string>,
  isAttr: boolean
): string | null {
  if (ns === '') return '';
  // An attribute in a namespace always needs a prefix: an unprefixed attribute has no namespace.
  if (authored && scope.get(authored) === ns) return authored;
  if (!isAttr && scope.get('') === ns) return '';
  for (const [prefix, uri] of scope) if (uri === ns && prefix !== '') return prefix;
  return null;
}

/**
 * A prefix to declare for `ns` that this scope does not use yet. Re-binding a prefix already in
 * scope would change the meaning of the element's own name or of a sibling attribute, so the
 * authored and conventional prefixes are taken only when free, and a generated one never
 * collides.
 */
function freshPrefix(ns: string, authored: string | undefined, scope: Map<string, string>): string {
  for (const candidate of [authored, conventionalPrefix(ns)])
    if (candidate && candidate !== 'xml' && candidate !== 'xmlns' && !scope.has(candidate))
      return candidate;
  for (let n = scope.size; ; n += 1) if (!scope.has(`ns${n}`)) return `ns${n}`;
}

function emit(el: Element, parentScope: ReadonlyMap<string, string>, out: string[]): void {
  if (!isNcName(el.local)) fail('element', `${JSON.stringify(el.local)} is not an XML name`);
  if (el.prefix && !isNcName(el.prefix))
    fail('element', `${JSON.stringify(el.prefix)} is not an XML prefix`);
  const elementError = qualifiedNameError(el.ns, el.local, el.prefix, false);
  if (elementError) fail(`element ${el.local}`, elementError.detail);
  const scope = new Map(parentScope);
  const declared = new Set<string>();
  let open = '';
  for (const d of el.nsDecls) {
    if (d.prefix !== '' && !isNcName(d.prefix))
      fail('namespace declaration', `${JSON.stringify(d.prefix)} is not an XML prefix`);
    const bindingError = namespaceBindingError(d.prefix, d.uri);
    if (bindingError) fail(`namespace declaration ${d.prefix}`, bindingError.detail);
    if (declared.has(d.prefix)) fail('namespace declaration', `${d.prefix} is declared twice`);
    declared.add(d.prefix);
    scope.set(d.prefix, d.uri);
    open += d.prefix
      ? ` xmlns:${d.prefix}="${escapeXmlAttributeChecked(d.uri, 'namespace')}"`
      : ` xmlns="${escapeXmlAttributeChecked(d.uri, 'namespace')}"`;
  }
  const extraDecls: string[] = [];
  // An element in no namespace cannot carry a prefix, so a default namespace in scope would
  // capture it on re-read. Undeclare the default namespace for this element.
  if (el.ns === '' && (scope.get('') ?? '') !== '') {
    if (declared.has('')) fail(`element ${el.local}`, 'no namespace under its own default');
    scope.set('', '');
    extraDecls.push(' xmlns=""');
  }
  const bind = (ns: string, authored: string | undefined, isAttr: boolean): string => {
    const found = prefixFor(ns, authored, scope, isAttr);
    if (found !== null) return found;
    // A node created by an edit, in a namespace this scope does not bind: declare it here.
    const prefix = freshPrefix(ns, authored, scope);
    scope.set(prefix, ns);
    extraDecls.push(` xmlns:${prefix}="${escapeXmlAttributeChecked(ns, 'namespace')}"`);
    return prefix;
  };
  const elPrefix = bind(el.ns, el.prefix, false);
  const name = elPrefix ? `${elPrefix}:${el.local}` : el.local;
  const attrText: string[] = [];
  const seen = new Set<string>();
  for (const a of el.attrs) {
    if (!isNcName(a.local)) fail('attribute', `${JSON.stringify(a.local)} is not an XML name`);
    if (a.prefix && !isNcName(a.prefix))
      fail('attribute', `${JSON.stringify(a.prefix)} is not an XML prefix`);
    const attrError = qualifiedNameError(a.ns, a.local, a.prefix, true);
    if (attrError) fail(`attribute ${a.local}`, attrError.detail);
    const expanded = `{${a.ns}}${a.local}`;
    if (seen.has(expanded)) fail(`attribute ${a.local}`, 'it appears twice on one element');
    seen.add(expanded);
    const ap = a.ns === '' ? '' : bind(a.ns, a.prefix, true);
    const an = ap ? `${ap}:${a.local}` : a.local;
    attrText.push(` ${an}="${escapeXmlAttributeChecked(a.value, `attribute ${an}`)}"`);
  }
  open = `<${name}${open}${extraDecls.join('')}${attrText.join('')}`;
  if (el.children.length === 0) {
    out.push(`${open}/>`);
    return;
  }
  out.push(`${open}>`);
  for (const child of el.children) {
    if (child.type === 'text') out.push(escapeXmlChecked(child.value, 'text'));
    else emit(child, scope, out);
  }
  out.push(`</${name}>`);
}

/**
 * The canonical form of a subtree: expanded names, attributes sorted by expanded name, no
 * prefixes, no namespace declarations, no ids. Two trees with the same canonical form are the
 * same document content however a writer spelled them.
 */
export function canonicalForm(node: Node): string {
  if (node.type === 'text') return `"${JSON.stringify(node.value)}`;
  const attrs = node.attrs.map((a) => `{${a.ns}}${a.local}=${JSON.stringify(a.value)}`).sort();
  const parts: string[] = [`<{${node.ns}}${node.local}`];
  for (const a of attrs) parts.push(a);
  parts.push('>');
  for (const child of node.children) parts.push(canonicalForm(child));
  parts.push('/>');
  return parts.join(' ');
}

const encoder = new TextEncoder();

/** SHA-256 of the canonical form: the structural fingerprint of one part. */
export function canonicalFingerprint(part: XmlPart): string {
  return sha256FontBytes(encoder.encode(canonicalForm(part.root)));
}

/**
 * SHA-256 of raw bytes, for binary parts. Uses the platform hash where the runtime has a
 * synchronous one (Bun, Node.js) and the pure-JavaScript one elsewhere; both give the same
 * digest.
 */
export function bytesFingerprint(bytes: Uint8Array): string {
  return sha256FontBytes(bytes);
}
