/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Children that OOXML allows once, written by concurrent peers, adapted from DOCX Editor's
// `packages/pro/src/collaboration/document/singleton-children.ts` at commit `5ffe967`. DOCX
// Editor names its own node kinds; this table names DrawingML and PresentationML elements. The
// functions work on replica specs, not id lists, so the exported API differs: `withCompanions`
// for `partitionWithMerges`, `groupOf` and `inSchemaOrder` for `singletonLayout` and
// `partitionChildIds`.
//
// Two peers that both give one shape its first `a:effectLst`, or both replace its fill, each
// write a valid part; merged, the shape holds two of an element its schema allows once, and
// a presentation app repairs the file. Two peers that add different children at one place (a
// reflection and a soft edge) get the order of their Yjs client ids, not the schema's.
//
// One table per parent gives the parent's child sequence (ECMA-376 Part 1 types) as groups:
// a group of one name is a singleton, a group of several is a choice (one fill), and `*` is
// the parent's repeating content (paragraphs, runs, shapes). OPC entries group by their key
// (`KEYED`). Every node keeps a minted key; the rule acts when the mirror reads:
//
// - The first copy of a group shows, and the later copies stay in the room as its companions
//   (`withCompanions`), so an undo of the first finds the next. A property container
//   (`MERGED_KINDS`) merges its companions child by child; `mirror.ts` writes each value to the
//   copy that holds it, so each author's undo reverts only that author's properties.
// - The children come back in the table's order (`inSchemaOrder`).

import {
  asciiFold,
  type ReplicaElementSpec,
  type ReplicaTextSpec,
} from '@pptx-editor.dev/core/store';
import { ATTR_SEP, ROOM_LIMITS } from './schema.ts';

const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const NS_P = 'http://schemas.openxmlformats.org/presentationml/2006/main';

/** One position of a child sequence: its slot name and the element names that fill it. */
type Group = readonly [slot: string, names: readonly string[]];

/** A position one element fills. */
const one = (name: string): Group => [name, [name]];
/** A position one of several elements fills. */
const choice = (slot: string, names: readonly string[]): Group => [slot, names];
/** The parent's repeating content: no singleton, and any child the table does not name. */
const CONTENT: Group = ['*', []];

/** EG_FillProperties. */
const FILL = choice('fill', ['noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill']);
/** EG_LineFillProperties. */
const LINE_FILL = choice('fill', ['noFill', 'solidFill', 'gradFill', 'pattFill']);
/** EG_EffectProperties. */
const EFFECT = choice('effect', ['effectLst', 'effectDag']);
/** EG_ColorChoice, the one child of CT_Color, CT_SolidColorFillProperties, and the effects. */
const COLOR = [
  choice('color', ['scrgbClr', 'srgbClr', 'hslClr', 'sysClr', 'schemeClr', 'prstClr']),
];

/** CT_ShapeProperties (CT_GroupShapeProperties is a subsequence). */
const SHAPE_PROPERTIES = [
  one('xfrm'),
  choice('geometry', ['custGeom', 'prstGeom']),
  FILL,
  one('ln'),
  EFFECT,
  one('scene3d'),
  one('sp3d'),
  one('extLst'),
];
/** CT_LineProperties: `a:ln` and the table cell borders. */
const LINE = [
  LINE_FILL,
  choice('dash', ['prstDash', 'custDash']),
  choice('join', ['round', 'bevel', 'miter']),
  one('headEnd'),
  one('tailEnd'),
  one('extLst'),
];
/** CT_TextCharacterProperties: `a:rPr`, `a:defRPr`, `a:endParaRPr`. */
const CHARACTER = [
  one('ln'),
  FILL,
  EFFECT,
  one('highlight'),
  choice('uLn', ['uLnTx', 'uLn']),
  choice('uFill', ['uFillTx', 'uFill']),
  one('latin'),
  one('ea'),
  one('cs'),
  one('sym'),
  one('hlinkClick'),
  one('hlinkMouseOver'),
  one('rtl'),
  one('extLst'),
];
/** The non-visual property groups: CT_ShapeNonVisual and its siblings. */
const NON_VISUAL = (locks: string) => [one('cNvPr'), one(locks), one('nvPr')];
/** The `p:cNv…Pr` elements: their locks, then an extension list. */
const LOCKS = (name: string) => [one(name), one('extLst')];
const TABLE_BORDERS = ['lnL', 'lnR', 'lnT', 'lnB', 'lnTlToBr', 'lnBlToTr'];

/** Parent local name → its child sequence. */
const CHILDREN: Readonly<Record<string, readonly Group[]>> = Object.assign(
  Object.create(null) as object,
  {
    // CT_Slide, CT_CommonSlideData, CT_Background, CT_BackgroundProperties.
    sld: [one('cSld'), one('clrMapOvr'), one('transition'), one('timing'), one('extLst')],
    cSld: [one('bg'), one('spTree'), one('custDataLst'), one('controls'), one('extLst')],
    bg: [choice('background', ['bgPr', 'bgRef'])],
    bgPr: [FILL, EFFECT, one('extLst')],
    spPr: SHAPE_PROPERTIES,
    grpSpPr: SHAPE_PROPERTIES,
    ln: LINE,
    ...Object.fromEntries(TABLE_BORDERS.map((b) => [b, LINE])),
    // CT_EffectList: each effect once.
    effectLst: [
      one('blur'),
      one('fillOverlay'),
      one('glow'),
      one('innerShdw'),
      one('outerShdw'),
      one('prstShdw'),
      one('reflection'),
      one('softEdge'),
    ],
    ...Object.fromEntries(
      [
        'solidFill',
        'outerShdw',
        'innerShdw',
        'prstShdw',
        'glow',
        'highlight',
        'buClr',
        'lnRef',
        'fillRef',
        'effectRef',
        'fontRef',
      ].map((n) => [n, COLOR])
    ),
    // CT_ShapeStyle.
    style: [one('lnRef'), one('fillRef'), one('effectRef'), one('fontRef')],
    // CT_Shape, CT_Picture, CT_Connector, CT_GroupShape, CT_GraphicalObjectFrame.
    sp: [one('nvSpPr'), one('spPr'), one('style'), one('txBody'), one('extLst')],
    pic: [one('nvPicPr'), one('blipFill'), one('spPr'), one('style'), one('extLst')],
    cxnSp: [one('nvCxnSpPr'), one('spPr'), one('style'), one('extLst')],
    grpSp: [one('nvGrpSpPr'), one('grpSpPr'), CONTENT, one('extLst')],
    graphicFrame: [one('nvGraphicFramePr'), one('xfrm'), one('graphic'), one('extLst')],
    nvSpPr: NON_VISUAL('cNvSpPr'),
    nvPicPr: NON_VISUAL('cNvPicPr'),
    nvCxnSpPr: NON_VISUAL('cNvCxnSpPr'),
    nvGrpSpPr: NON_VISUAL('cNvGrpSpPr'),
    nvGraphicFramePr: NON_VISUAL('cNvGraphicFramePr'),
    // CT_NonVisualDrawingProps and CT_ApplicationNonVisualDrawingProps.
    cNvPr: [one('hlinkClick'), one('hlinkHover'), one('extLst')],
    nvPr: [
      one('ph'),
      choice('media', ['audioCd', 'wavAudioFile', 'audioFile', 'videoFile', 'quickTimeFile']),
      one('custDataLst'),
      one('extLst'),
    ],
    cNvSpPr: LOCKS('spLocks'),
    cNvPicPr: LOCKS('picLocks'),
    cNvGrpSpPr: LOCKS('grpSpLocks'),
    cNvGraphicFramePr: LOCKS('graphicFrameLocks'),
    cNvCxnSpPr: [one('cxnSpLocks'), one('stCxn'), one('endCxn'), one('extLst')],
    // CT_Transform2D, CT_GroupTransform2D.
    xfrm: [one('off'), one('ext'), one('chOff'), one('chExt')],
    // CT_BlipFillProperties (`p:blipFill`, `a:blipFill`): the image, its crop, its fill mode.
    blipFill: [one('blip'), one('srcRect'), choice('fillMode', ['tile', 'stretch'])],
    // CT_TextBody, CT_TextBodyProperties, CT_TextParagraph, and the text runs.
    txBody: [one('bodyPr'), one('lstStyle'), CONTENT],
    bodyPr: [
      one('prstTxWarp'),
      choice('autofit', ['noAutofit', 'normAutofit', 'spAutoFit']),
      one('scene3d'),
      choice('text3d', ['sp3d', 'flatTx']),
      one('extLst'),
    ],
    p: [one('pPr'), CONTENT, one('endParaRPr')],
    r: [one('rPr'), one('t')],
    fld: [one('rPr'), one('pPr'), one('t')],
    br: [one('rPr')],
    // CT_TextParagraphProperties.
    pPr: [
      one('lnSpc'),
      one('spcBef'),
      one('spcAft'),
      choice('buClr', ['buClrTx', 'buClr']),
      choice('buSz', ['buSzTx', 'buSzPct', 'buSzPts']),
      choice('buFont', ['buFontTx', 'buFont']),
      choice('bullet', ['buNone', 'buAutoNum', 'buChar', 'buBlip']),
      one('tabLst'),
      one('defRPr'),
      one('extLst'),
    ],
    rPr: CHARACTER,
    defRPr: CHARACTER,
    endParaRPr: CHARACTER,
    // CT_Table, CT_TableProperties, CT_TableCell, CT_TableCellProperties.
    tbl: [one('tblPr'), one('tblGrid'), CONTENT],
    tblPr: [FILL, EFFECT, choice('tableStyle', ['tableStyle', 'tableStyleId']), one('extLst')],
    tc: [one('txBody'), one('tcPr'), one('extLst')],
    tcPr: [...TABLE_BORDERS.map(one), one('cell3D'), FILL, one('headers'), one('extLst')],
  }
);

type Name = { readonly ns: string; readonly local: string };

const ooxml = (n: Name): boolean => n.ns === NS_A || n.ns === NS_P;

/**
 * Parents that exist only in PresentationML. Their DrawingML namesakes have other sequences
 * (`a:graphicFrame`, CT_GvmlGraphicalObjectFrame, puts `a:graphic` before `a:xfrm`).
 */
const PRESENTATION_ONLY = new Set([
  'sld',
  'cSld',
  'bg',
  'bgPr',
  'sp',
  'pic',
  'cxnSp',
  'grpSp',
  'graphicFrame',
  'nvSpPr',
  'nvPicPr',
  'nvCxnSpPr',
  'nvGrpSpPr',
  'nvGraphicFramePr',
  'nvPr',
  'style',
]);

/** The child sequence of a parent, or undefined for a parent the table does not list. */
function childrenOf(parent: Name): readonly Group[] | undefined {
  if (!ooxml(parent) || (PRESENTATION_ONLY.has(parent.local) && parent.ns !== NS_P))
    return undefined;
  return CHILDREN[parent.local];
}

/** The position of a named child in its parent's sequence, or -1. */
function positionOf(groups: readonly Group[], child: Name): number {
  return ooxml(child) ? groups.findIndex(([, names]) => names.includes(child.local)) : -1;
}

/**
 * The property containers whose concurrent copies merge child by child, after DOCX Editor's
 * `MERGED_KINDS` (`w:pPr`, `w:rPr`, `w:tblPr`): paragraph, run, and text body properties,
 * table and cell properties, and line properties.
 */
const MERGED_KINDS = new Set([
  'pPr',
  'rPr',
  'defRPr',
  'endParaRPr',
  'bodyPr',
  'tblPr',
  'tcPr',
  // CT_LineProperties: an outline and the cell borders. Its width, cap, and compound type are
  // attributes, and its fill and dash are choice slots that keep the first copy's child.
  'ln',
  ...TABLE_BORDERS,
]);

/** Whether concurrent copies of a child merge on read (a property container). */
export function mergesCopies(child: Name): boolean {
  return ooxml(child) && MERGED_KINDS.has(child.local);
}

/**
 * The position of a child in its parent's sequence, by slot name (`fill` for any fill), or
 * undefined for repeating content and children the table does not name.
 */
export function positionSlotOf(parent: Name, child: Name): string | undefined {
  const groups = childrenOf(parent);
  const at = groups ? positionOf(groups, child) : -1;
  return at < 0 ? undefined : groups![at]![0];
}

/**
 * A table cell's text body. ECMA-376 Part 1, 21.1.3.16 tc allows one `a:txBody` or none, and two
 * peers that type into a cell with none each give it one. Both stay: a hidden copy would take
 * its author's text from every replica. Core draws and saves the cell with one body that holds
 * the paragraphs of both (`cellAsSaved`), and an edit of the cell or a table command joins them
 * in the store. The schema order still puts the body first (`inSchemaOrder`).
 */
const cellBody = (parent: Name, child: Name): boolean =>
  parent.ns === NS_A && parent.local === 'tc' && child.ns === NS_A && child.local === 'txBody';

/**
 * OPC entries that one key names once, as DOCX Editor builds `[Content_Types].xml` from maps
 * keyed by extension and by part name. ECMA-376 Part 2 compares an extension and a part name
 * without case (10.1.2.2), and a relationship `Id` (an `xsd:ID`, 9.3) as written. A redo after
 * an undo repair, or two peers that add parts, can write one entry twice.
 */
const NS_CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const NS_REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
type EntryKey = { readonly attr: string; readonly caseless: boolean };
const KEYED: Readonly<Record<string, Readonly<Record<string, EntryKey>>>> = Object.assign(
  Object.create(null) as object,
  {
    [`${NS_CT}|Types`]: {
      Default: { attr: 'Extension', caseless: true },
      Override: { attr: 'PartName', caseless: true },
    },
    [`${NS_REL}|Relationships`]: { Relationship: { attr: 'Id', caseless: false } },
  }
);

/** How concurrent copies of a child group (DOCX Editor's `groupKey`). */
export interface ChildGroup {
  /** The group's name among its siblings. */
  readonly slot: string;
  /** `merge`: a property container; `first`: the first copy shows whole; `key`: an OPC entry. */
  readonly rule: 'merge' | 'first' | 'key';
}

/**
 * The group of a child under its parent, or undefined for repeating content and children the
 * tables do not name. `attrOf` reads an attribute of the child, for an OPC entry's key.
 */
export function groupOf(
  parent: Name,
  child: Name,
  attrOf: (local: string) => string | undefined
): ChildGroup | undefined {
  const keyed = KEYED[`${parent.ns}|${parent.local}`];
  if (keyed && Object.prototype.hasOwnProperty.call(keyed, child.local)) {
    const { attr, caseless } = keyed[child.local]!;
    const value = attrOf(attr);
    if (value === undefined) return undefined;
    return { slot: `${child.local}:${caseless ? asciiFold(value) : value}`, rule: 'key' };
  }
  const slot = positionSlotOf(parent, child);
  if (slot === undefined || cellBody(parent, child)) return undefined;
  return { slot, rule: mergesCopies(child) ? 'merge' : 'first' };
}

/**
 * `children` in their parent's schema order: a stable sort by position. A child the table
 * does not name takes the content position (`*`) when the parent has one, else the position
 * of the child before it, so it stays after its neighbor. The same list comes back when it
 * is already in order or the parent is not listed.
 */
export function inSchemaOrder<T>(
  parent: Name,
  children: readonly T[],
  nameOf: (child: T) => Name | undefined
): readonly T[] {
  const groups = childrenOf(parent);
  if (!groups) return children;
  const content = groups.indexOf(CONTENT);
  let last = -1;
  let sorted = true;
  const ranked = children.map((c, i) => {
    const name = nameOf(c);
    const own = name ? positionOf(groups, name) : -1;
    const rank = own >= 0 ? own : content >= 0 ? content : last;
    if (rank < last) sorted = false;
    last = rank;
    return { c, rank, i };
  });
  if (sorted) return children;
  return ranked.sort((a, b) => a.rank - b.rank || a.i - b.i).map((r) => r.c);
}

// The read rule: copies of a group as a replica shows them.

type Spec = ReplicaElementSpec;
type ReplicaChild = ReplicaElementSpec | ReplicaTextSpec;
type Attr = NonNullable<Spec['attrs']>[number];

const attrKey = (a: { readonly ns: string; readonly local: string }): string =>
  `${a.ns}${ATTR_SEP}${a.local}`;

/** The client that minted a key: the part before its first `.`. */
export const clientOf = (key: string): string => key.slice(0, Math.max(0, key.indexOf('.')));

/** The mark of a key that the seed wrote: `_` is no base-36 digit, so no minted key has it. */
export const SEED_MARK = '_';

/** Whether the seed wrote a key (`client._N`): the opened file's own content. */
export const fromSeed = (key: string): boolean =>
  key.slice(key.indexOf('.') + 1).startsWith(SEED_MARK);

/**
 * How a replica shows the copy that won a concurrent group (DOCX Editor's companions): the
 * copies it hides, which stay in the room, and, for a property container, the companion that
 * holds each attribute or child it shows from a copy.
 */
export interface MergeView {
  /** A property container: its companions merge, and an edit goes to this client's copy. */
  merge: boolean;
  readonly companions: string[];
  /** Attribute key → the companion record that holds the shown value. */
  readonly attrFrom: Map<string, string>;
  /** Child key → the companion record that lists the child. */
  readonly childFrom: Map<string, string>;
  /**
   * Child key of a companion → its slot, for a child that does not show because the slot
   * shows another copy's child. It stays in its copy while the container fills the slot.
   */
  readonly shadowed: Map<string, string>;
}

function viewOf(views: Map<string, MergeView>, key: string): MergeView {
  let view = views.get(key);
  if (!view) {
    view = {
      merge: false,
      companions: [],
      attrFrom: new Map(),
      childFrom: new Map(),
      shadowed: new Map(),
    };
    views.set(key, view);
  }
  return view;
}

const attrValue = (spec: Spec, local: string): string | undefined =>
  spec.attrs?.find((a) => a.ns === '' && a.local === local)?.value;

/**
 * Every concurrent copy of one property container as one (after DOCX Editor's
 * `MERGED_KINDS`): the first copy's attributes and children, then each attribute and child
 * position that only a later copy has, in copy order. A child container that two copies hold
 * merges the same way. One pass over all copies, so the cost grows with their size, not with
 * its square. Returns null when a nested group holds too many copies (`withCompanions`).
 */
function mergeCopies(
  base: Spec,
  copies: readonly Spec[],
  views: Map<string, MergeView>
): Spec | null {
  const view = viewOf(views, base.key);
  view.merge = true;
  const attrs = new Map<string, Attr>();
  for (const a of base.attrs ?? []) attrs.set(attrKey(a), a);
  const slotOfKid = (c: ReplicaChild) =>
    'text' in c ? undefined : groupOf(base, c, (l) => attrValue(c, l))?.slot;
  const baseKids = base.children ?? [];
  const taken = new Set(baseKids.map(slotOfKid));
  const kids: ReplicaChild[] = [...baseKids];
  /** Child key → the copy whose record lists it. */
  const addedFrom = new Map<string, string>();
  for (const copy of copies) {
    view.companions.push(copy.key);
    for (const a of copy.attrs ?? []) {
      const k = attrKey(a);
      if (attrs.has(k)) continue;
      attrs.set(k, a);
      view.attrFrom.set(k, copy.key);
    }
    for (const c of copy.children ?? []) {
      if ('text' in c) continue;
      const slot = slotOfKid(c);
      if (slot === undefined) {
        // A child with no slot (an extension, `mc:AlternateContent`) is no singleton: it shows,
        // listed by its copy, so no write drops it from a peer's record.
        kids.push(c);
        addedFrom.set(c.key, copy.key);
        continue;
      }
      if (taken.has(slot) && !mergesCopies(c)) {
        // The slot shows another copy's child: this one stays in its copy, shadowed.
        view.shadowed.set(c.key, slot);
        continue;
      }
      taken.add(slot);
      kids.push(c);
      addedFrom.set(c.key, copy.key);
    }
  }
  const merged = withCompanions(base, kids, views);
  if (!merged) return null;
  // A child from a copy that stays shown is listed by the copy's record.
  for (const c of merged) {
    const from = 'text' in c ? undefined : addedFrom.get(c.key);
    if (from !== undefined) view.childFrom.set(c.key, from);
  }
  const children = inSchemaOrder(base, merged, (k) => ('text' in k ? undefined : k));
  const sorted = [...attrs.entries()].sort(([a], [b]) => (a < b ? -1 : a > b ? 1 : 0));
  return { ...base, attrs: sorted.map(([, a]) => a), children: [...children] };
}

/**
 * `kids` as a replica shows them, after DOCX Editor's `partitionWithMerges`: the first copy
 * of each group shows, and the later copies stay in the room as its companions. A property
 * container merges its companions; an OPC entry or another singleton shows the first whole.
 * Copies that the seed wrote are the opened file's own duplicates and stay as loaded, as in
 * DOCX Editor ("a group nobody wrote into"), OPC entries too.
 *
 * Returns null when one group holds more than `ROOM_LIMITS.maxCopiesPerGroup` concurrent copies:
 * only a hostile peer writes that many, and the mirror refuses the part.
 */
export function withCompanions(
  parent: Name,
  kids: readonly ReplicaChild[],
  views: Map<string, MergeView>
): ReplicaChild[] | null {
  const first = new Map<string, number>();
  const out: ReplicaChild[] = [];
  /** Index in `out` → the concurrent copies of its group. */
  const copiesAt = new Map<number, Spec[]>();
  for (const k of kids) {
    const group = 'text' in k ? undefined : groupOf(parent, k, (l) => attrValue(k, l));
    const at = group ? first.get(group.slot) : undefined;
    const winner = at === undefined ? undefined : (out[at] as Spec);
    const concurrent = winner !== undefined && !(fromSeed(winner.key) && fromSeed((k as Spec).key));
    if (group && at === undefined) first.set(group.slot, out.length);
    if (!concurrent) {
      out.push(k);
      continue;
    }
    let copies = copiesAt.get(at!);
    if (!copies) copiesAt.set(at!, (copies = []));
    if (copies.length >= ROOM_LIMITS.maxCopiesPerGroup) return null;
    copies.push(k as Spec);
  }
  for (const [at, copies] of copiesAt) {
    const winner = out[at] as Spec;
    if (groupOf(parent, winner, (l) => attrValue(winner, l))?.rule === 'merge') {
      const merged = mergeCopies(winner, copies, views);
      if (!merged) return null;
      out[at] = merged;
    } else {
      const view = viewOf(views, winner.key);
      for (const c of copies) view.companions.push(c.key);
    }
  }
  return out;
}
