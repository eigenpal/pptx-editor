// Package-level edit helpers shared by shape and slide commands: relationships, content types,
// shape ids, and new part names. All compile to ops on an OpBatch.

import {
  attr,
  children,
  CONTENT_TYPES_PART,
  elements,
  find,
  NS,
  REL,
  relsPartName,
  xmlPartOf,
  type Element,
  type NodeSpec,
} from '../../store/index.ts';
import { partNameKey } from '@eigenpal/ooxml-core';
import { escapeXmlAttributeChecked, escapeXmlChecked } from '@eigenpal/ooxml-core';
import { CommandRefusal, type OpBatch } from './op-batch.ts';
import {
  ACTOR_ID_STRIPE,
  actorStripe,
  MAX_RELATIONSHIP_NUMBER,
  relationshipIdFromNumber,
  relationshipNumberOf,
} from './actor-scoped-ids.ts';

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';

/** An attribute value for an XML template, validated and escaped. Refuses invalid characters. */
export function xmlAttr(value: string, what: string): string {
  try {
    return escapeXmlAttributeChecked(value, what);
  } catch {
    throw new CommandRefusal('errors.unsupported', what);
  }
}

/** A field GUID in braces (`a:fld/@id`, `ST_Guid`). */
export function fieldId(): string {
  const hex = (n: number) =>
    Array.from({ length: n }, () => Math.floor(Math.random() * 16).toString(16)).join('');
  return `{${hex(8)}-${hex(4)}-4${hex(3)}-${'89AB'[Math.floor(Math.random() * 4)]}${hex(3)}-${hex(12)}}`.toUpperCase();
}

/** Character data for an XML template, validated and escaped. Refuses invalid characters. */
export function xmlText(value: string, what: string): string {
  try {
    return escapeXmlChecked(value, what);
  } catch {
    throw new CommandRefusal('errors.unsupported', what);
  }
}

/**
 * `ST_CoordinateUnqualified` (ECMA-376 Part 1, `dml-main.xsd`): the range of `a:off`, `a:chOff`
 * x and y. `ST_PositiveCoordinate` (`a:ext`, `a:chExt` cx and cy) runs from 0 to the same
 * maximum. Inside it, an integer never prints in exponent notation.
 */
export const MIN_COORDINATE = -27_273_042_329_600;
export const MAX_COORDINATE = 27_273_042_316_900;

/** An `ST_Coordinate` attribute value (EMU, rounded). Refuses a value out of the range. */
export function coordinate(value: number, what: string): string {
  const v = Math.round(value);
  if (!Number.isFinite(v) || v < MIN_COORDINATE || v > MAX_COORDINATE)
    throw new CommandRefusal('errors.coordinateRange', what);
  return String(v + 0);
}

/** An `ST_PositiveCoordinate` attribute value (EMU, rounded). Refuses a value out of the range. */
export function positiveCoordinate(value: number, what: string): string {
  const v = Math.round(value);
  if (!Number.isFinite(v) || v < 0 || v > MAX_COORDINATE)
    throw new CommandRefusal('errors.coordinateRange', what);
  return String(v + 0);
}

/** A relative target from one part to another (`../media/image3.png`). */
export function relativeTarget(from: string, to: string): string {
  const a = from.split('/').slice(1, -1);
  const b = to.split('/').slice(1);
  let i = 0;
  while (i < a.length && i < b.length - 1 && a[i] === b[i]) i += 1;
  return [...a.slice(i).map(() => '..'), ...b.slice(i)].join('/');
}

/** Add a relationship from `source` to `target` and return its new id. Creates the rels part. */
export function addRelationship(
  batch: OpBatch,
  source: string,
  type: string,
  target: string
): string {
  return addRelationships(batch, source, type, [target])[0]!;
}

/**
 * Add a relationship of `type` from `source` to each of `targets`, in order, and return their
 * new ids. The ids in use are read once for all of them. Creates the rels part.
 */
export function addRelationships(
  batch: OpBatch,
  source: string,
  type: string,
  targets: readonly string[]
): string[] {
  const relsName = relsPartName(source);
  if (!xmlPartOf(batch.pkg, relsName))
    batch.push({
      op: 'addPart',
      part: relsName,
      xml: `${XML_DECL}<Relationships xmlns="${xmlAttr(NS.rel, 'namespace')}"/>`,
    });
  const rels = xmlPartOf(batch.pkg, relsName)!;
  const used = new Set<string>();
  const numbers = new Set<number>();
  let max = 0;
  const take = (value: string | undefined): void => {
    if (!value) return;
    used.add(value);
    // A file's `rId` too large to count from (past nine digits) is skipped.
    const number = relationshipNumberOf(value);
    if (number !== null) {
      numbers.add(number);
      max = Math.max(max, number);
    }
  };
  for (const r of elements(rels.root)) take(attr(r, 'Id'));
  // An id that the source part still names stays taken, also when its relationship is gone: a
  // slide that one collaborator deleted while another moved it comes back in `p:sldIdLst` with
  // its old `r:id`, and a new relationship with that id would bring it back.
  const stack = [xmlPartOf(batch.pkg, source)?.root];
  for (let el = stack.pop(); el; el = stack.pop()) {
    for (const a of el.attrs) if (a.ns === NS.r) take(a.value);
    for (const c of el.children) if (c.type === 'element') stack.push(c);
  }
  // In a collaboration session the number comes from the actor's stripe (DOCX Editor's
  // actor-scoped ids), so two peers do not mint one; alone it is one past the largest.
  const next = batch.allocationActor
    ? // The same cursor as every other kind; relationship numbers start at 1, as a file's own do.
      stripedIds(batch.allocationActor, numbers, 1, MAX_RELATIONSHIP_NUMBER, batch.mintedOf('rId'))
    : null;
  let n = max;
  const index = rels.root.children.length;
  return targets.map((target, i) => {
    let id: string;
    if (next) id = relationshipIdFromNumber(next());
    else {
      n += 1;
      while (used.has(relationshipIdFromNumber(n))) n += 1;
      id = relationshipIdFromNumber(n);
    }
    used.add(id);
    batch.push({
      op: 'insertNode',
      part: relsName,
      parent: rels.root.id,
      index: index + i,
      spec: {
        ns: NS.rel,
        local: 'Relationship',
        attrs: [
          { ns: '', local: 'Id', value: id },
          { ns: '', local: 'Type', value: type },
          { ns: '', local: 'Target', value: relativeTarget(source, target) },
        ],
      },
    });
    return id;
  });
}

/** Add an external relationship (a hyperlink); its target is recorded, never fetched. */
export function addExternalRelationship(
  batch: OpBatch,
  source: string,
  type: string,
  target: string
): string {
  const id = addRelationship(batch, source, type, '/placeholder');
  const rels = xmlPartOf(batch.pkg, relsPartName(source))!;
  const rel = elements(rels.root).find((r) => attr(r, 'Id') === id)!;
  for (const [name, value] of [
    ['Target', target],
    ['TargetMode', 'External'],
  ] as const)
    batch.push({
      op: 'setAttribute',
      part: relsPartName(source),
      node: rel.id,
      name: { ns: '', local: name },
      value,
    });
  return id;
}

/** Remove a relationship by id. */
export function removeRelationship(batch: OpBatch, source: string, id: string): void {
  const rels = xmlPartOf(batch.pkg, relsPartName(source));
  const rel = rels ? elements(rels.root).find((r) => attr(r, 'Id') === id) : undefined;
  if (rel) batch.push({ op: 'removeNode', part: relsPartName(source), node: rel.id });
}

/**
 * Remove the nodes `ids` of `part`, last first: a removal shifts only the siblings after it,
 * so none of the others moves.
 */
function removeLastFirst(batch: OpBatch, part: string, ids: readonly string[]): void {
  for (let i = ids.length - 1; i >= 0; i -= 1) batch.push({ op: 'removeNode', part, node: ids[i] });
}

/** Remove relationships of `source` by id, in one pass over its relationships part. */
export function removeRelationships(batch: OpBatch, source: string, ids: readonly string[]): void {
  const gone = new Set(ids);
  const rels = xmlPartOf(batch.pkg, relsPartName(source));
  const nodes: string[] = [];
  // The first relationship with an id goes, as `removeRelationship` removes it.
  for (const rel of elements(rels?.root))
    if (gone.delete(attr(rel, 'Id') ?? '')) nodes.push(rel.id);
  removeLastFirst(batch, relsPartName(source), nodes);
}

/** Whether an element of `root` names `id` in an `r:*` attribute. Bounded at 256 levels. */
function namesRelationship(root: Element, id: string): boolean {
  const stack: { el: Element; depth: number }[] = [{ el: root, depth: 0 }];
  while (stack.length > 0) {
    const { el, depth } = stack.pop()!;
    if (el.attrs.some((a) => a.ns === NS.r && a.value === id)) return true;
    if (depth < 256) for (const c of elements(el)) stack.push({ el: c, depth: depth + 1 });
  }
  return false;
}

/**
 * Remove the hyperlink relationship `id` of `source` when nothing in the part names it any
 * more. A relationship of another type than `types`, or one that an element still names, stays.
 */
export function removeUnusedHyperlink(
  batch: OpBatch,
  source: string,
  id: string,
  types: readonly string[] = [REL.hyperlink]
): void {
  const rels = xmlPartOf(batch.pkg, relsPartName(source));
  const rel = rels ? elements(rels.root).find((r) => attr(r, 'Id') === id) : undefined;
  if (!rel || !types.includes(attr(rel, 'Type') ?? '')) return;
  const xml = xmlPartOf(batch.pkg, source);
  if (xml && namesRelationship(xml.root, id)) return;
  batch.push({ op: 'removeNode', part: relsPartName(source), node: rel.id });
}

/** Add a content-type override for a part. */
export function addOverride(batch: OpBatch, partName: string, contentType: string): void {
  addOverrides(batch, [[partName, contentType]]);
}

/** Add content-type overrides, as `[partName, contentType]`, in one read of the content types. */
export function addOverrides(
  batch: OpBatch,
  overrides: readonly (readonly [string, string])[]
): void {
  if (overrides.length === 0) return;
  const ct = xmlPartOf(batch.pkg, CONTENT_TYPES_PART);
  if (!ct) throw new CommandRefusal('errors.malformed');
  const index = ct.root.children.length;
  overrides.forEach(([partName, contentType], i) =>
    batch.push({
      op: 'insertNode',
      part: CONTENT_TYPES_PART,
      parent: ct.root.id,
      index: index + i,
      spec: {
        ns: NS.ct,
        local: 'Override',
        attrs: [
          { ns: '', local: 'PartName', value: partName },
          { ns: '', local: 'ContentType', value: contentType },
        ],
      },
    })
  );
}

/**
 * Ensure a content type for a new media part: the `Default` of its extension, as a deck usually
 * holds it. In a collaboration session a new extension gets an `Override` for `partName`
 * instead. An `Override` belongs to one part, so an undo that removes one peer's picture removes
 * only that picture's entry.
 */
export function ensureContentType(
  batch: OpBatch,
  extension: string,
  contentType: string,
  partName?: string
): void {
  const ct = xmlPartOf(batch.pkg, CONTENT_TYPES_PART);
  if (!ct) throw new CommandRefusal('errors.malformed');
  const exists = elements(ct.root).some(
    (e) =>
      e.local === 'Default' &&
      (attr(e, 'Extension') ?? '').toLowerCase() === extension.toLowerCase()
  );
  if (exists) return;
  if (batch.allocationActor && partName) return addOverride(batch, partName, contentType);
  batch.push({
    op: 'insertNode',
    part: CONTENT_TYPES_PART,
    parent: ct.root.id,
    index: 0,
    spec: {
      ns: NS.ct,
      local: 'Default',
      attrs: [
        { ns: '', local: 'Extension', value: extension },
        { ns: '', local: 'ContentType', value: contentType },
      ],
    },
  });
}

/**
 * Remove the content-type overrides of `partNames`, in one pass over the content types: the
 * first override of each part goes, as `removeOverride` removes it.
 */
export function removeOverrides(batch: OpBatch, partNames: readonly string[]): void {
  const ct = xmlPartOf(batch.pkg, CONTENT_TYPES_PART);
  const gone = new Set(partNames.map(partNameKey));
  const nodes: string[] = [];
  for (const e of elements(ct?.root))
    if (e.local === 'Override' && gone.delete(partNameKey(attr(e, 'PartName') ?? '')))
      nodes.push(e.id);
  removeLastFirst(batch, CONTENT_TYPES_PART, nodes);
}

/** Remove a content-type override. Part names compare the OPC way (ASCII case-insensitive). */
export function removeOverride(batch: OpBatch, partName: string): void {
  const ct = xmlPartOf(batch.pkg, CONTENT_TYPES_PART);
  const key = partNameKey(partName);
  const el = ct
    ? elements(ct.root).find(
        (e) => e.local === 'Override' && partNameKey(attr(e, 'PartName') ?? '') === key
      )
    : undefined;
  if (el) batch.push({ op: 'removeNode', part: CONTENT_TYPES_PART, node: el.id });
}

/**
 * The first free part name matching `prefix<n>suffix`. Names compare the OPC way (ASCII
 * case-insensitive), so `/ppt/media/Image1.png` takes `image1.png`. A name the session held
 * before (a deleted slide) stays taken, and so does a name whose relationships part exists.
 */
export function freePartName(batch: OpBatch, prefix: string, suffix: string): string {
  const present = new Set<string>();
  for (const name of batch.pkg.parts.keys()) present.add(partNameKey(name));
  const taken = (name: string): boolean => {
    const key = partNameKey(name);
    return present.has(key) || batch.reservedPartNames.has(key);
  };
  // Alone 1, 2, 3, ...; in a collaboration session the actor's full stripe, not DOCX Editor's
  // smaller part-number stripe (4096): two peers that share a part name lose a slide or a
  // picture, and with 4096 classes twenty peers share one in about 4.6 % of rooms.
  const striped = batch.allocationActor
    ? stripedIds(batch.allocationActor, new Set(), 1, MAX_MINTED_ID, batch.mintedOf('part'))
    : null;
  let dense = 0;
  for (let tries = 0; tries < 100_000; tries += 1) {
    const n = striped ? striped() : ++dense;
    const name = `${prefix}${n}${suffix}`;
    if (!taken(name) && !taken(relsPartName(name))) return name;
  }
  throw new CommandRefusal('errors.tooLarge');
}

/**
 * The part namer of a command that adds many parts: `freePartName` with the names in use read
 * once. Each name it gives counts as taken for the next. In a collaboration session it is
 * `freePartName` itself, whose names come from the actor's stripe.
 */
export function partNamer(batch: OpBatch): (prefix: string, suffix: string) => string {
  if (batch.allocationActor) return (prefix, suffix) => freePartName(batch, prefix, suffix);
  const present = new Set<string>();
  for (const name of batch.pkg.parts.keys()) present.add(partNameKey(name));
  const taken = (name: string): boolean => {
    const key = partNameKey(name);
    return present.has(key) || batch.reservedPartNames.has(key);
  };
  // The lowest number that may be free, for each pattern: the numbers below it are taken.
  const from = new Map<string, number>();
  return (prefix, suffix) => {
    const pattern = `${prefix}\u0000${suffix}`;
    for (let n = from.get(pattern) ?? 1, tries = 0; tries < 100_000; n += 1, tries += 1) {
      const name = `${prefix}${n}${suffix}`;
      if (taken(name) || taken(relsPartName(name))) continue;
      from.set(pattern, n + 1);
      present.add(partNameKey(name));
      return name;
    }
    throw new CommandRefusal('errors.tooLarge');
  };
}

/** `xsd:unsignedInt`: the largest value of `ST_DrawingElementId`, `ST_Index`, and their kin. */
export const MAX_UNSIGNED_INT = 4_294_967_295;

/**
 * The largest id that a new element gets: 2^31 - 1. ECMA-376 Part 1, 20.1.10.21
 * `ST_DrawingElementId` and `pml.xsd` `ST_Index` are `xsd:unsignedInt` (to 4294967295), and
 * `ST_SlideId` stops at 2147483647. New ids stay at or below the smaller bound, so a reader that
 * stores an id in a signed 32-bit integer reads every id that the editor writes.
 */
export const MAX_MINTED_ID = 2_147_483_647;

/** An `xsd:unsignedInt` in its lexical form, after `whiteSpace="collapse"`. */
const UNSIGNED_INT = /^[ \t\r\n]*(?:\+?(\d+)|-(0+))[ \t\r\n]*$/;

/**
 * An id attribute from the file as a number from `min` to `max`, read as `xsd:unsignedInt`
 * (XML Schema Part 2, 3.3.22): surrounding whitespace, a `+` sign, and leading zeros are
 * allowed, so " 7", "+7", and "007" are id 7, as a schema-aware reader takes them. Any other
 * value (a minus sign, an exponent, hex, a fraction, a number past `max`) reads as undefined: it
 * is not a schema id, so it sets no maximum. Every value from a file is attacker-controlled
 * (AGENTS.md "Security"); the pattern runs in linear time on any input.
 */
export function unsignedIntId(
  raw: string | undefined,
  min = 0,
  max = MAX_UNSIGNED_INT
): number | undefined {
  const m = raw === undefined ? null : UNSIGNED_INT.exec(raw);
  if (!m) return undefined;
  const digits = (m[1] ?? '0').replace(/^0+(?=\d)/, '');
  if (digits.length > 10) return undefined;
  const n = Number(digits);
  return n >= min && n <= max ? n : undefined;
}

/**
 * A source of ids from `min` to `max` that `used` does not hold: one past the largest of
 * `used` (and of `after`), counting up. When that passes `max`, the lowest free ids from `min`.
 * Each id comes once. Refuses (`errors.tooLarge`) when no id is left. The work is linear in the
 * size of `used`: the scan for low ids steps only over ids that `used` holds.
 */
export function freeIds(
  used: ReadonlySet<number>,
  min: number,
  max: number,
  after = min - 1
): () => number {
  let highest = Math.min(after, max);
  for (const n of used) if (n > highest && n <= max) highest = n;
  const start = highest + 1;
  let high = start;
  let low = min;
  return () => {
    if (high <= max) return high++;
    // Every id from `start` to `max` is given; below `start`, take the ids `used` does not hold.
    while (low < start && used.has(low)) low += 1;
    if (low < start) return low++;
    throw new CommandRefusal('errors.tooLarge');
  };
}

/**
 * A source of ids in the actor's stripe (`n ≡ actorStripe(actor) (mod ACTOR_ID_STRIPE)`) from
 * `min` to `max`, skipping `used` and `minted`. It is DOCX Editor's `stripedDecimalIdSequence`
 * with a floor, because PresentationML numbers have floors (slide ids from 256, shape ids from
 * 1). Refuses (`errors.tooLarge`) past `max`: a wrap or a dense fallback would hand back a
 * number that another replica can mint.
 *
 * `minted` is the session's record for one id kind (`OpBatch.mintedOf`): every number this
 * replica handed out, and every number of its stripe that the opened file holds. Only this
 * replica mints in its stripe, so a number in the record is never given again, also after a
 * peer deleted the shape or slide that held it and that peer's undo brings it back.
 */
export function stripedIds(
  actor: string,
  used: ReadonlySet<number>,
  min: number,
  max: number,
  minted?: Set<number>
): () => number {
  const stripe = actorStripe(actor);
  let n = min + ((stripe - (min % ACTOR_ID_STRIPE) + ACTOR_ID_STRIPE) % ACTOR_ID_STRIPE);
  return () => {
    while (n <= max && (used.has(n) || minted?.has(n))) n += ACTOR_ID_STRIPE;
    if (n > max) throw new CommandRefusal('errors.tooLarge');
    const out = n;
    minted?.add(out);
    n += ACTOR_ID_STRIPE;
    return out;
  };
}

/**
 * The `p:cNvPr/@id` values under `root`, as numbers. A value that is not an
 * `ST_DrawingElementId` (`unsignedIntId`) is left out.
 */
export function drawingElementIds(root: Element): Set<number> {
  const used = new Set<number>();
  const walk = (el: Element, depth: number): void => {
    if (depth > 64) return;
    if (el.local === 'cNvPr') {
      const id = unsignedIntId(attr(el, 'id'));
      if (id !== undefined) used.add(id);
    }
    for (const c of elements(el)) walk(c, depth + 1);
  };
  walk(root, 0);
  return used;
}

/**
 * A source of ids for new shapes on a slide, unique among the `p:cNvPr/@id` values of its
 * shape tree (the `shape-id` rule of `part-checks.ts`). One past the largest id, counting up;
 * past `MAX_MINTED_ID`, the lowest free ids. Id 1 is the shape tree's own, so the count starts
 * after it. In a collaboration session the ids come from the actor's stripe (`stripedIds`), so
 * two peers that insert shapes at once never give two shapes one id.
 */
export function shapeIds(batch: OpBatch, root: Element): () => number {
  const used = drawingElementIds(root);
  if (batch.allocationActor)
    return stripedIds(batch.allocationActor, used, 1, MAX_MINTED_ID, batch.mintedOf('shapeId'));
  return freeIds(used, 1, MAX_MINTED_ID, 1);
}

/** The id for one new shape on a slide (`shapeIds`). */
export function nextShapeId(batch: OpBatch, root: Element): number {
  return shapeIds(batch, root)();
}

/** Connector ends: `a:stCxn` and `a:endCxn` name the `p:cNvPr/@id` of the shape they glue to. */
const CONNECTION_ENDS = new Set(['stCxn', 'endCxn']);

/**
 * Give copied shapes new `p:cNvPr` ids from `first` on, in document order, and rewrite the
 * connector ends that point at a copied shape. An end whose shape was not copied is removed,
 * so the copy is not glued to a shape it was not copied with.
 */
export function renumberShapes(
  specs: readonly NodeSpec[],
  first: number | (() => number)
): { readonly specs: NodeSpec[]; readonly next: number } {
  // A number counts up from it; a source (`shapeIds`) gives each id, also outside a run, and
  // then `next` is only one past the last id given.
  let counter = typeof first === 'number' ? first : 0;
  const take = typeof first === 'number' ? () => counter++ : first;
  const fresh: number[] = [];
  const isCNvPr = (n: NodeSpec): boolean => n.ns === NS.p && n.local === 'cNvPr';
  const idOf = (n: NodeSpec): string | undefined =>
    n.attrs?.find((a) => a.ns === '' && a.local === 'id')?.value;
  // Old ids compare as `xsd:unsignedInt` (`unsignedIntId`), so an end that names "0005" stays
  // glued to the shape whose id is "5".
  const key = (id: string | undefined): string | undefined =>
    id === undefined ? undefined : String(unsignedIntId(id) ?? id);
  const map = new Map<string, string>();
  const collect = (n: NodeSpec, depth: number): void => {
    if (depth > 64) return;
    if (isCNvPr(n)) {
      const old = key(idOf(n));
      const id = take();
      fresh.push(id);
      if (old !== undefined && !map.has(old)) map.set(old, String(id));
    }
    for (const c of n.children ?? []) if (!('text' in c)) collect(c, depth + 1);
  };
  for (const s of specs) collect(s, 0);
  let next = 0;
  const rewrite = (n: NodeSpec, depth: number): NodeSpec | null => {
    if (depth > 64) return n;
    let attrs = n.attrs;
    if (isCNvPr(n))
      attrs = (attrs ?? []).map((a) =>
        a.ns === '' && a.local === 'id' ? { ...a, value: String(fresh[next++]) } : a
      );
    else if (n.ns === NS.a && CONNECTION_ENDS.has(n.local)) {
      const target = map.get(key(idOf(n)) ?? '');
      if (target === undefined) return null;
      attrs = (attrs ?? []).map((a) =>
        a.ns === '' && a.local === 'id' ? { ...a, value: target } : a
      );
    }
    const kids: (NodeSpec | { text: string })[] = [];
    for (const c of n.children ?? []) {
      if ('text' in c) kids.push(c);
      else {
        const k = rewrite(c, depth + 1);
        if (k) kids.push(k);
      }
    }
    return { ...n, ...(attrs ? { attrs } : {}), children: kids };
  };
  const out = specs.map((s) => rewrite(s, 0) ?? s);
  return { specs: out, next: (fresh.at(-1) ?? (typeof first === 'number' ? first - 1 : 0)) + 1 };
}

/** The slide's `p:spTree`. */
export function spTreeOf(batch: OpBatch, slidePart: string): Element {
  const tree = find(xmlPartOf(batch.pkg, slidePart)?.root, 'p:spTree');
  if (!tree) throw new CommandRefusal('errors.malformed');
  return tree;
}

/** The parent element of a node (a shape's tree or group). */
export function parentOf(batch: OpBatch, part: string, id: string): Element {
  const parentId = batch.parentId(part, id);
  if (!parentId) throw new CommandRefusal('errors.unsupported', `missing node ${id}`);
  return batch.element(part, parentId);
}

/** Content children of a shape tree (shapes, pictures, groups, frames). */
export function shapeChildren(tree: Element): Element[] {
  return elements(tree).filter(
    (e) => !(e.ns === NS.p && (e.local === 'nvGrpSpPr' || e.local === 'grpSpPr'))
  );
}

export { children, XML_DECL };
