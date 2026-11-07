// Small builders over immutable trees, for the diagram edits that rewrite the data model and the
// saved drawing as whole parts (`diagram-structure.ts`, `diagram-relayout.ts`). A built element
// has no node id: the part is serialized and written as bytes, and the reader gives ids again.

import { NS, type Attr, type Element, type Node } from '../../store/index.ts';

/** A new element. Attributes are unqualified; values are escaped when the part is written. */
export function el(
  ns: string,
  prefix: string,
  local: string,
  attrs: Readonly<Record<string, string>> = {},
  kids: readonly Node[] = []
): Element {
  return {
    type: 'element',
    id: '',
    kind: 'generic',
    ns,
    local,
    prefix,
    nsDecls: [],
    attrs: Object.entries(attrs).map(([local, value]): Attr => ({ ns: '', local, value })),
    children: kids,
  };
}

export const dgm = (local: string, attrs?: Record<string, string>, kids?: readonly Node[]) =>
  el(NS.dgm, 'dgm', local, attrs, kids);
export const dml = (local: string, attrs?: Record<string, string>, kids?: readonly Node[]) =>
  el(NS.a, 'a', local, attrs, kids);

/** `el` with the unqualified attribute `local` set to `value`, or removed for null. */
export function withAttr(e: Element, local: string, value: string | null): Element {
  const rest = e.attrs.filter((a) => a.ns !== '' || a.local !== local);
  const at = e.attrs.findIndex((a) => a.ns === '' && a.local === local);
  if (value === null) return at < 0 ? e : { ...e, attrs: rest };
  const attr: Attr = { ns: '', local, value };
  if (at < 0) return { ...e, attrs: [...e.attrs, attr] };
  return { ...e, attrs: e.attrs.map((a, i) => (i === at ? attr : a)) };
}

/** `parent` with its children replaced. */
export const withChildren = (parent: Element, kids: readonly Node[]): Element => ({
  ...parent,
  children: kids,
});

/** `parent` with each child element mapped; text nodes stay. */
export function mapElements(parent: Element, fn: (e: Element) => Element | null): Element {
  const kids: Node[] = [];
  for (const c of parent.children) {
    if (c.type !== 'element') {
      kids.push(c);
      continue;
    }
    const next = fn(c);
    if (next) kids.push(next);
  }
  return withChildren(parent, kids);
}

/** Replace the first descendant reached by the child names `path` with `fn` of it. */
export function updatePath(
  root: Element,
  path: readonly string[],
  fn: (e: Element) => Element
): Element {
  const [head, ...rest] = path;
  if (head === undefined) return fn(root);
  const [prefix, local] = head.split(':');
  const ns = prefix === 'a' ? NS.a : prefix === 'dsp' ? NS.dsp : NS.dgm;
  let done = false;
  return mapElements(root, (c) => {
    if (done || c.ns !== ns || c.local !== local) return c;
    done = true;
    return updatePath(c, rest, fn);
  });
}
