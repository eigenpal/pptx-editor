// PresentationML checks of the parts a transaction touched, run once at the commit boundary.
//
// After DOCX Editor's `TreeDocumentStore.commitTransaction` (store/store/tree-store.ts): the
// transaction records the parts it rewrote, and the commit validates those and no others, so
// the cost follows the size of the edit, not the size of the deck. An invalid result abandons
// the whole transaction: no revision, no history entry, no change event.
//
// The checks compare the package before and after the transaction and refuse only a violation
// that the transaction INTRODUCED. Reading stays lenient: a deck that arrives with a broken
// part opens and stays editable, and an edit that leaves an old violation as it was commits.
// Each violation has a stable key (part, rule, and the node id or relationship id it is
// about), so a violation the base package already had is never charged to the edit.
//
// Rules:
// - `table-grid`: every `a:tr` holds one `a:tc` per `a:gridCol` (ECMA-376 Part 1, 21.1.3.16 tc;
//   `CT_TableRow`, `CT_TableGrid`).
// - `table-merge`: `gridSpan` and `rowSpan` stay inside the grid, blocks do not overlap, every
//   cell a merged block covers carries `hMerge` (off the block's first column) and `vMerge`
//   (off its first row), and every merge marker lies inside a block (21.1.3.16 tc,
//   `CT_TableCell`). The blocks are the ones the cascade draws (`table-grid.ts`), and each
//   violation is keyed by its cell: a file's orphan marker stays editable, and an edit that
//   makes a new one is refused.
// - `shape-id`: `p:cNvPr/@id` (`CT_NonVisualDrawingProps`, `ST_DrawingElementId`) is unique
//   among the shapes of one slide, layout, or master.
// - `slide-list`: every `p:sldId`, `p:sldMasterId`, and `p:notesMasterId` of the presentation
//   part names, by its required `r:id`, an internal relationship whose target part exists
//   (19.2.1.33 sldId; `CT_SlideIdListEntry`, `CT_SlideMasterIdListEntry`,
//   `CT_NotesMasterIdListEntry`).
// - `relationship-target`: every internal relationship of a touched `.rels` part targets a part
//   of the package, and no relationship anywhere targets a part the transaction removed
//   (ECMA-376 Part 2, Open Packaging Conventions, Relationships). External targets are
//   recorded only, never resolved or fetched.
// - `content-type`: every part has a content type, by `Override` or by `Default` extension
//   (ECMA-376 Part 2, Open Packaging Conventions, Content Types stream). When the transaction
//   edits `[Content_Types].xml`, only the parts whose `Override` or `Default` it removed can
//   lose their type, so only those are read.

import { asciiFold, normalizePartName, partNameKey } from '@eigenpal/ooxml-core';
import {
  CONTENT_TYPES_PART,
  contentTypeOf,
  equivalentPartName,
  mainPresentationPart,
  relationshipsOf,
  relsPartName,
  type PptxPackage,
} from '../package/package.ts';
import { NS } from '../package/namespaces.ts';
import { attr, child, children, elements } from '../package/query.ts';
import { resolveTableGrid, tableMergeFaults } from '../package/table-grid.ts';
import type { Element, XmlPart } from '../package/tree.ts';
import type { PackageInvariantRule, TreeInvariantRule } from './ops.ts';

interface ViolationBase {
  readonly part: string;
  readonly detail: string;
  /** Stable across edits that do not change what the violation is about. */
  readonly key: string;
}

/** One violation: of a part's own tree, or of a rule that reads more than one part. */
export type PartViolation =
  | (ViolationBase & { readonly kind: 'tree'; readonly rule: TreeInvariantRule })
  | (ViolationBase & { readonly kind: 'package'; readonly rule: PackageInvariantRule });

/** Groups nest at most 32 deep (AGENTS.md resource limits); the XML reader caps depth at 256. */
const MAX_WALK_DEPTH = 256;

const isP = (el: Element, local: string): boolean => el.ns === NS.p && el.local === local;
const isA = (el: Element, local: string): boolean => el.ns === NS.a && el.local === local;

// ---------------------------------------------------------------- tables

function tablesIn(el: Element, depth: number, out: Element[]): void {
  if (depth > MAX_WALK_DEPTH) return;
  for (const c of el.children) {
    if (c.type !== 'element') continue;
    if (isA(c, 'tbl')) out.push(c);
    else tablesIn(c, depth + 1, out);
  }
}

function checkTable(part: string, tbl: Element, out: PartViolation[]): void {
  const grid = children(child(tbl, 'a:tblGrid'), 'a:gridCol').length;
  const rows = children(tbl, 'a:tr').map((tr) => children(tr, 'a:tc'));
  // ECMA-376 Part 1, 21.1.3.16 tc: one `a:tc` per grid column in each row; a merged cell is
  // still a `tc`, marked with `hMerge` or `vMerge`.
  for (let r = 0; r < rows.length; r += 1)
    if (rows[r].length !== grid) {
      const n = rows[r].length;
      out.push({
        kind: 'tree',
        rule: 'table-grid',
        part,
        detail: `row ${r} has ${n} a:tc for ${grid} a:gridCol (one a:tc per column)`,
        key: `${part}\ttable-grid\t${tbl.id}`,
      });
      return;
    }
  // The blocks as the cascade draws them (`resolveTableGrid`). Each violation is keyed by the
  // cell it is about, so an edit that leaves a file's orphan marker as it was commits, and an
  // edit that makes a new one is refused.
  const nR = rows.length;
  const g = resolveTableGrid(rows, grid);
  const bad = (tc: Element, what: string, detail: string): void => {
    out.push({
      kind: 'tree',
      rule: 'table-merge',
      part,
      detail,
      key: `${part}\ttable-merge\t${tc.id}\t${what}`,
    });
  };
  for (const f of tableMergeFaults(rows, g)) {
    const tc = rows[f.row][f.col];
    const at = `cell (row ${f.row}, column ${f.col})`;
    if (f.fault === 'past-grid') bad(tc, f.fault, `${at} spans past the ${nR} x ${grid} grid`);
    else if (f.fault === 'overlap') bad(tc, f.fault, `${at} spans into another merged block`);
    else if (f.fault === 'orphan')
      bad(tc, f.fault, `${at} has hMerge or vMerge outside a merged block`);
    else if (f.fault === 'unmarked')
      bad(tc, 'marker', `${at} lies in a merged block but has no hMerge or vMerge`);
    else bad(tc, 'marker', `${at} has hMerge or vMerge that disagree with its merged block`);
  }
}

// ---------------------------------------------------------------- shape ids

/** The `p:cNvPr` of a shape-tree member: `p:nvSpPr`, `p:nvPicPr`, and the like. */
function cNvPrOf(shape: Element): Element | undefined {
  for (const c of shape.children)
    if (c.type === 'element' && c.ns === NS.p && c.local.startsWith('nv'))
      return child(c, 'p:cNvPr');
  return undefined;
}

const SHAPES = new Set(['sp', 'pic', 'cxnSp', 'graphicFrame', 'contentPart']);

/**
 * The shape ids of a shape tree, in document order. Only shape-tree members count: a picture
 * inside `p:oleObj` is not a shape of the slide. Of `mc:AlternateContent`, only the first
 * `mc:Choice` counts, because its `mc:Fallback` repeats the same shapes with the same ids.
 */
function shapeIds(tree: Element, depth: number, out: Element[]): void {
  if (depth > MAX_WALK_DEPTH) return;
  for (const c of tree.children) {
    if (c.type !== 'element') continue;
    if (c.ns === NS.p) {
      if (c.local === 'nvGrpSpPr') {
        const id = child(c, 'p:cNvPr');
        if (id) out.push(id);
      } else if (SHAPES.has(c.local)) {
        const id = cNvPrOf(c);
        if (id) out.push(id);
      } else if (c.local === 'grpSp') shapeIds(c, depth + 1, out);
    } else if (c.ns === NS.mc && c.local === 'AlternateContent') {
      const choice = elements(c).find((e) => e.ns === NS.mc && e.local === 'Choice');
      if (choice) shapeIds(choice, depth + 1, out);
    }
  }
}

const SHAPE_ID_ROOTS = new Set(['sld', 'sldLayout', 'sldMaster']);

function checkShapeIds(xml: XmlPart, out: PartViolation[]): void {
  const root = xml.root;
  if (root.ns !== NS.p || !SHAPE_ID_ROOTS.has(root.local)) return;
  const tree = child(child(root, 'p:cSld'), 'p:spTree');
  if (!tree) return;
  const ids: Element[] = [];
  shapeIds(tree, 0, ids);
  const seen = new Set<string>();
  const reported = new Set<string>();
  for (const el of ids) {
    const id = attr(el, 'id');
    if (id === undefined) continue;
    if (seen.has(id) && !reported.has(id)) {
      reported.add(id);
      out.push({
        kind: 'tree',
        rule: 'shape-id',
        part: xml.partName,
        detail: `p:cNvPr id="${id}" is used by more than one shape`,
        key: `${xml.partName}\tshape-id\t${id}`,
      });
    }
    seen.add(id);
  }
}

const treeCache = new WeakMap<XmlPart, readonly PartViolation[]>();

/** The violations of one part's own tree. Memoized per part object, so a commit pays once. */
export function treeViolations(xml: XmlPart): readonly PartViolation[] {
  const hit = treeCache.get(xml);
  if (hit) return hit;
  const out: PartViolation[] = [];
  if (xml.root.ns === NS.p) {
    const tables: Element[] = [];
    tablesIn(xml.root, 0, tables);
    for (const tbl of tables) checkTable(xml.partName, tbl, out);
    checkShapeIds(xml, out);
  }
  treeCache.set(xml, out);
  return out;
}

// ---------------------------------------------------------------- package rules

/** The source part of a `.rels` part (`/` for the package), or undefined. */
function relsSource(partName: string): string | undefined {
  if (partName === '/_rels/.rels') return '/';
  const m = /^(.*)\/_rels\/([^/]+)\.rels$/.exec(partName);
  return m ? `${m[1]}/${m[2]}` : undefined;
}

const exists = (pkg: PptxPackage, partName: string): boolean =>
  equivalentPartName(pkg, partName) !== undefined;

function checkRelationships(pkg: PptxPackage, relsPart: string, out: PartViolation[]): void {
  const source = relsSource(relsPart);
  if (source === undefined || relsPartName(source) !== relsPart) return;
  for (const rel of relationshipsOf(pkg, source)) {
    if (rel.external || !rel.target || exists(pkg, rel.target)) continue;
    out.push({
      kind: 'package',
      rule: 'relationship-target',
      part: relsPart,
      detail: `relationship ${rel.id} targets ${rel.target}, which is not in the package`,
      key: `${relsPart}\trelationship-target\t${rel.id}\t${partNameKey(rel.target)}`,
    });
  }
}

/** The id lists of `p:presentation` whose entries name a part by `r:id`. */
const ID_LISTS: readonly (readonly [string, string])[] = [
  ['sldMasterIdLst', 'sldMasterId'],
  ['notesMasterIdLst', 'notesMasterId'],
  ['sldIdLst', 'sldId'],
];

function checkSlideList(pkg: PptxPackage, main: string, out: PartViolation[]): void {
  const part = pkg.parts.get(main);
  if (!part || part.kind !== 'xml') return;
  const live = new Map<string, string>();
  for (const rel of relationshipsOf(pkg, main))
    if (!rel.external && rel.target && exists(pkg, rel.target)) live.set(rel.id, rel.target);
  for (const [list, entry] of ID_LISTS)
    for (const l of part.xml.root.children) {
      if (l.type !== 'element' || !isP(l, list)) continue;
      for (const e of l.children) {
        if (e.type !== 'element' || !isP(e, entry)) continue;
        const rid = attr(e, 'r:id') ?? '';
        if (live.has(rid)) continue;
        out.push({
          kind: 'package',
          rule: 'slide-list',
          part: main,
          detail: `p:${entry} r:id="${rid}" names no relationship to a part of the package`,
          key: `${main}\tslide-list\t${entry}\t${rid}`,
        });
      }
    }
}

function checkContentType(pkg: PptxPackage, partName: string, out: PartViolation[]): void {
  if (partName === CONTENT_TYPES_PART || !pkg.parts.has(partName)) return;
  if (contentTypeOf(pkg, partName) !== undefined) return;
  out.push({
    kind: 'package',
    rule: 'content-type',
    part: partName,
    detail: `${partName} has no content type (no Override for it, no Default for its extension)`,
    key: `${partName}\tcontent-type`,
  });
}

/** The `Default` and `Override` elements of the content types part of a package. */
const declarations = (pkg: PptxPackage): Element[] => {
  const part = pkg.parts.get(CONTENT_TYPES_PART);
  return part?.kind === 'xml' ? elements(part.xml.root).filter((e) => e.ns === NS.ct) : [];
};

/**
 * The parts that can have lost their content type: the touched parts, and, when the content
 * types part changed, the parts that a `Default` or `Override` the transaction removed or
 * changed typed. The tree is persistent, so an element the ops did not touch is the same
 * object in both packages: only the others are read. Only a changed `Default` reads the part
 * names of the whole package, and only their extensions.
 */
function typeCandidates(
  before: PptxPackage,
  after: PptxPackage,
  touched: ReadonlySet<string>
): Set<string> {
  const out = new Set(touched);
  if (!touched.has(CONTENT_TYPES_PART)) return out;
  const kept = new Set(declarations(after));
  const lost: string[] = [];
  for (const el of declarations(before)) {
    if (kept.has(el)) continue;
    if (el.local === 'Default') {
      const ext = attr(el, 'Extension');
      if (ext) lost.push(asciiFold(ext));
    } else if (el.local === 'Override') {
      const name = attr(el, 'PartName');
      const norm = name ? normalizePartName(name) : undefined;
      const part = norm?.ok ? equivalentPartName(after, norm.partName) : undefined;
      if (part) out.add(part);
    }
  }
  if (lost.length > 0)
    for (const name of after.parts.keys()) {
      const dot = name.lastIndexOf('.');
      if (dot > name.lastIndexOf('/') && lost.includes(asciiFold(name.slice(dot + 1))))
        out.add(name);
    }
  return out;
}

/** What a commit must check. */
interface Scope {
  /** The parts whose object differs between the two packages. */
  readonly touched: ReadonlySet<string>;
  /** The parts to check for a content type. */
  readonly typed: ReadonlySet<string>;
  /** Whether the transaction removed a part. */
  readonly removed: boolean;
}

function violations(pkg: PptxPackage, scope: Scope): PartViolation[] {
  const out: PartViolation[] = [];
  for (const name of scope.touched) {
    const part = pkg.parts.get(name);
    if (!part) continue;
    if (part.kind === 'xml') for (const v of treeViolations(part.xml)) out.push(v);
    if (name.endsWith('.rels')) checkRelationships(pkg, name, out);
  }
  for (const name of scope.typed) checkContentType(pkg, name, out);
  // The presentation part lists parts by relationship: check it when it, its relationships,
  // or the part set changed.
  const main = mainPresentationPart(pkg);
  if (main && (scope.removed || scope.touched.has(main) || scope.touched.has(relsPartName(main))))
    checkSlideList(pkg, main, out);
  return out;
}

/**
 * A relationship of an untouched `.rels` part whose target the transaction removed. Every such
 * relationship is new: its target existed before the transaction. Runs only when a part was
 * removed, which is a structural edit (delete a slide or a picture), never typing.
 */
function danglingAfterRemoval(
  after: PptxPackage,
  touched: ReadonlySet<string>,
  removed: ReadonlySet<string>
): PartViolation | null {
  for (const relsPart of after.parts.keys()) {
    if (touched.has(relsPart) || !relsPart.endsWith('.rels')) continue;
    const source = relsSource(relsPart);
    if (source === undefined || relsPartName(source) !== relsPart) continue;
    if (source !== '/' && !after.parts.has(source)) continue;
    for (const rel of relationshipsOf(after, source)) {
      if (rel.external || !rel.target || !removed.has(partNameKey(rel.target))) continue;
      if (exists(after, rel.target)) continue;
      return {
        kind: 'package',
        rule: 'relationship-target',
        part: relsPart,
        detail: `relationship ${rel.id} targets ${rel.target}, which the transaction removed`,
        key: `${relsPart}\trelationship-target\t${rel.id}\t${partNameKey(rel.target)}`,
      };
    }
  }
  return null;
}

/**
 * The first violation that `after` has in the touched parts and `before` did not have, or null.
 * `touched` names every part whose object differs between the two packages.
 */
export function introducedViolation(
  before: PptxPackage,
  after: PptxPackage,
  touched: readonly string[]
): PartViolation | null {
  if (touched.length === 0) return null;
  const set = new Set(touched);
  const removed = new Set<string>();
  for (const name of touched)
    if (before.parts.has(name) && !after.parts.has(name)) removed.add(partNameKey(name));
  const scope: Scope = {
    touched: set,
    typed: typeCandidates(before, after, set),
    removed: removed.size > 0,
  };
  const found = violations(after, scope);
  if (found.length > 0) {
    const old = new Set(violations(before, scope).map((v) => v.key));
    const fresh = found.find((v) => !old.has(v.key) && !copiesDangling(before, v));
    if (fresh) return fresh;
  }
  return removed.size > 0 ? danglingAfterRemoval(after, set, removed) : null;
}

/**
 * Whether `v` is a relationship to a part that a relationship of `before` already names and
 * that `before` lacks too: a copy of a slide with a missing picture or chart copies that
 * relationship, and the copy is no worse than the file.
 */
function copiesDangling(before: PptxPackage, v: PartViolation): boolean {
  if (v.rule !== 'relationship-target') return false;
  const target = v.key.slice(v.key.lastIndexOf('\t') + 1);
  for (const name of before.parts.keys()) {
    const source = relsSource(name);
    if (source === undefined || relsPartName(source) !== name) continue;
    for (const rel of relationshipsOf(before, source))
      if (!rel.external && rel.target && partNameKey(rel.target) === target)
        return !exists(before, rel.target);
  }
  return false;
}

/** Every violation of a whole package, for tests and diagnostics. @internal */
export function packageViolations(pkg: PptxPackage): PartViolation[] {
  const all = new Set(pkg.parts.keys());
  return violations(pkg, { touched: all, typed: all, removed: true });
}
