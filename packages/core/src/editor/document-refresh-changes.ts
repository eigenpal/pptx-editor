// The changes of an accepted document refresh result, after DOCX Editor's
// document-refresh-changes.ts (docx-editor/packages/core/src/editor at b63cbea4a).
//
// A change comes from the update (`changes`) or, without them, from the review module, which
// compares the captured presentation with the returned one (`refreshChanges`). Each location
// is checked against the returned presentation: the slide, the shape on it, the paragraph, and
// the exact text at the offsets. A change that does not check out gets a status other than
// `available` and a diagnostic, and never a highlight. A computed change of a layout, a master,
// or a theme has no location of its own: it highlights the slides that show it.

import type {
  RefreshChange,
  RefreshChangeInput,
  RefreshLocation,
} from './document-refresh-types.ts';
import type {
  ComputedRefreshChange,
  RefreshChangeAspect,
  RefreshComputedChange,
  RefreshDiagnosticCode,
} from './document-refresh-types.ts';
import { paragraphsOf, paragraphText } from '../automation/commands/text-model.ts';
import {
  elements,
  presentationInfo,
  xmlPartOf,
  type Element,
  type PptxPackage,
} from '../store/index.ts';

/** Where a located change highlights. */
export interface RefreshTarget {
  readonly slideIndex: number;
  /** The slide part, or its notes part for `notes`. */
  readonly part: string;
  readonly notes?: boolean;
  /** The shape's element id in the open store. */
  readonly nodeId?: string;
  /** A text range of that shape, as the editor's text points. */
  readonly text?: { readonly paragraph: number; readonly start: number; readonly end: number };
}

/** One change of an accepted result, with where it highlights. @internal */
export interface LocatedChange {
  readonly change: RefreshChange;
  /** What makes the change new when it differs from the previous result's. */
  readonly fingerprint: string;
  /** The places it highlights; empty when it has none. */
  readonly targets: readonly RefreshTarget[];
}

const MIN_SLIDE_ID = 256;
const MAX_SLIDE_ID = 2147483647;

const isIndex = (v: unknown): v is number =>
  typeof v === 'number' && Number.isSafeInteger(v) && v >= 0;

/** The text of each paragraph of a shape's text body, as the editor counts offsets. @internal */
export function shapeParagraphTexts(shape: Element): string[] {
  const body = elements(shape).find((c) => c.local === 'txBody');
  return body ? paragraphsOf(body).map((p) => paragraphText(p)) : [];
}

/** The shapes of a part whose `p:cNvPr/@id` is `id`, in groups too. @internal */
export function shapesById(root: Element | undefined, id: number): Element[] {
  const out: Element[] = [];
  const visit = (el: Element, depth: number): void => {
    if (depth > 64) return;
    for (const child of elements(el)) {
      const nv = elements(child)[0];
      const cNvPr =
        nv && nv.local.startsWith('nv') ? elements(nv).find((c) => c.local === 'cNvPr') : undefined;
      const value = cNvPr?.attrs.find((a) => a.ns === '' && a.local === 'id')?.value;
      if (value !== undefined && Number(value) === id) out.push(child);
      if (child.local === 'spTree' || child.local === 'grpSp' || child.local === 'cSld')
        visit(child, depth + 1);
    }
  };
  if (root) visit(root, 0);
  return out;
}

type Checked =
  | { readonly ok: true; readonly target: RefreshTarget }
  | { readonly ok: false; readonly code: RefreshDiagnosticCode };

/** Check `location` against `pkg`. */
function check(pkg: PptxPackage, location: RefreshLocation, rendered: Rendered): Checked {
  const { slideId, slideIndex, shapeId, paragraphIndex, start, end, text } = location;
  const offsets = start !== undefined || end !== undefined || text !== undefined;
  if (
    (slideId === undefined) === (slideIndex === undefined) ||
    (slideIndex !== undefined && !isIndex(slideIndex)) ||
    (shapeId !== undefined && !isIndex(shapeId)) ||
    (paragraphIndex !== undefined && (shapeId === undefined || !isIndex(paragraphIndex))) ||
    (offsets && paragraphIndex === undefined)
  )
    return { ok: false, code: 'invalid-selector' };
  if (
    slideId !== undefined &&
    (!Number.isSafeInteger(slideId) || slideId < MIN_SLIDE_ID || slideId > MAX_SLIDE_ID)
  )
    return { ok: false, code: 'invalid-slide-id' };
  const slides = presentationInfo(pkg).slides;
  const index =
    slideIndex !== undefined ? slideIndex : slides.findIndex((s) => Number(s.sldId) === slideId);
  const slide = slides[index];
  if (!slide) return { ok: false, code: 'slide-not-found' };
  const part = location.notes ? slide.notesPart : slide.partName;
  if (!part) return { ok: false, code: 'slide-not-found' };
  const notes = location.notes ? { notes: true as const } : {};
  if (shapeId === undefined) return { ok: true, target: { slideIndex: index, part, ...notes } };
  const found = shapesById(xmlPartOf(pkg, part)?.root, shapeId);
  if (found.length === 0) return { ok: false, code: 'shape-not-found' };
  if (found.length > 1) return { ok: false, code: 'ambiguous-shape-id' };
  const shape = found[0]!;
  if (!location.notes && !rendered(index, shape.id)) return { ok: false, code: 'not-rendered' };
  const base = { slideIndex: index, part, nodeId: shape.id, ...notes };
  if (paragraphIndex === undefined) return { ok: true, target: base };
  const texts = shapeParagraphTexts(shape);
  const paragraph = texts[paragraphIndex];
  if (paragraph === undefined) return { ok: false, code: 'paragraph-not-found' };
  if (!offsets) {
    const whole = { paragraph: paragraphIndex, start: 0, end: paragraph.length };
    return { ok: true, target: { ...base, text: whole } };
  }
  if (!isIndex(start) || !isIndex(end) || start >= end || end > paragraph.length)
    return { ok: false, code: 'invalid-offsets' };
  if (paragraph.slice(start, end) !== text) return { ok: false, code: 'text-mismatch' };
  return { ok: true, target: { ...base, text: { paragraph: paragraphIndex, start, end } } };
}

/** Whether shape `nodeId` of slide `index` has a layout position. */
export type Rendered = (index: number, nodeId: string) => boolean;

/**
 * What makes a change new, as in DOCX Editor: its text and its description, here with what a
 * computed change changed. Positions are not part of it: a slide inserted before a change
 * moves its slide index, and the change is not new.
 */
const fingerprintOf = (input: RefreshChangeInput): string => {
  const computed = (input as Partial<ComputedRefreshChange>).computed;
  return JSON.stringify([
    input.description ?? null,
    input.location?.text ?? null,
    input.unavailableReason ?? null,
    computed ? [computed.scope, computed.action, computed.aspects ?? null] : null,
  ]);
};

/**
 * The located changes of one result. `seen` holds the fingerprints of the previous results of
 * the submission: a change is new when its id is new, or its fingerprint changed.
 *
 * @internal
 */
export function resolveRefreshChanges(
  pkg: PptxPackage,
  resultId: string,
  inputs: readonly RefreshChangeInput[],
  seen: ReadonlyMap<string, string>,
  rendered: Rendered
): LocatedChange[] {
  const slideCount = presentationInfo(pkg).slides.length;
  return inputs.map((input) => {
    const fingerprint = fingerprintOf(input);
    const computed = (input as Partial<ComputedRefreshChange>).computed;
    const base = {
      id: input.id,
      ...(input.description !== undefined ? { description: input.description } : {}),
      resultId,
      isNew: seen.get(input.id) !== fingerprint,
      ...(computed ? { computed } : {}),
    };
    const done = (
      change: RefreshChange,
      targets: readonly RefreshTarget[] = []
    ): LocatedChange => ({
      change: Object.freeze(change),
      fingerprint,
      targets,
    });
    if (input.unavailableReason) {
      const status = input.unavailableReason === 'deleted' ? 'deleted' : 'unavailable';
      return done({
        ...base,
        status,
        diagnostic: {
          code: input.unavailableReason === 'deleted' ? 'deleted' : 'missing-location',
          ...(input.location ? { location: input.location } : {}),
        },
      });
    }
    if (!input.location) {
      // A layout, master, or theme change shows on the slides that use it.
      const affected = (computed?.affectedSlides ?? []).filter((i) => isIndex(i) && i < slideCount);
      if (affected.length > 0) {
        const slides = presentationInfo(pkg).slides;
        const targets = affected.map((i) => ({ slideIndex: i, part: slides[i]!.partName }));
        return done({ ...base, status: 'available' }, targets);
      }
      return done({ ...base, status: 'unavailable', diagnostic: { code: 'missing-location' } });
    }
    const location = input.location;
    const checked = check(pkg, location, rendered);
    if (!checked.ok)
      return done({
        ...base,
        status: checked.code === 'not-rendered' ? 'unavailable' : 'invalid',
        diagnostic: { code: checked.code, location },
      });
    return done({ ...base, status: 'available', location }, [checked.target]);
  });
}

/** Whether `value` is a well-typed location, before any check against the presentation. */
export function isLocationShape(value: unknown): value is RefreshLocation {
  if (!value || typeof value !== 'object') return false;
  const v = value as Record<string, unknown>;
  const num = (k: string) => v[k] === undefined || typeof v[k] === 'number';
  return (
    num('slideId') &&
    num('slideIndex') &&
    num('shapeId') &&
    num('paragraphIndex') &&
    num('start') &&
    num('end') &&
    (v.text === undefined || typeof v.text === 'string') &&
    (v.notes === undefined || typeof v.notes === 'boolean')
  );
}

const SCOPES: ReadonlySet<string> = new Set<RefreshComputedChange['scope']>([
  'slide',
  'shape',
  'text',
  'notes',
  'layout',
  'master',
  'theme',
  'package',
]);
const ACTIONS: ReadonlySet<string> = new Set<RefreshComputedChange['action']>([
  'added',
  'removed',
  'moved',
  'changed',
]);
/**
 * The aspects a computed change may name: `RefreshChangeAspect`, and the aspects a newer review
 * module may add. A string outside the list makes the update malformed.
 */
const ASPECTS: ReadonlySet<string> = new Set([
  'position',
  'size',
  'rotation',
  'geometry',
  'fill',
  'line',
  'effects',
  'text',
  'textFormat',
  'picture',
  'table',
  'chart',
  'diagram',
  'transition',
  'animation',
  'background',
  'layout',
  'visibility',
  'other',
]);
/** The most aspects and affected slides one computed change may name. */
const MAX_COMPUTED_ITEMS = 10000;

/** Whether `value` is a well-typed `computed` field of a change. */
export function isComputedShape(value: unknown): value is RefreshComputedChange {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return false;
  const v = value as Record<string, unknown>;
  const list = (k: string, ok: (x: unknown) => boolean) =>
    v[k] === undefined ||
    (Array.isArray(v[k]) &&
      (v[k] as unknown[]).length <= MAX_COMPUTED_ITEMS &&
      (v[k] as unknown[]).every(ok));
  return (
    typeof v.scope === 'string' &&
    SCOPES.has(v.scope) &&
    typeof v.action === 'string' &&
    ACTIONS.has(v.action) &&
    list('aspects', (x) => typeof x === 'string' && ASPECTS.has(x)) &&
    list('affectedSlides', isIndex) &&
    (v.previousSlideIndex === undefined || isIndex(v.previousSlideIndex))
  );
}

/** A frozen copy of a `computed` field with only its documented fields. */
export function ownComputed(c: RefreshComputedChange): RefreshComputedChange {
  return Object.freeze({
    scope: c.scope,
    action: c.action,
    ...(c.aspects ? { aspects: Object.freeze(c.aspects.slice() as RefreshChangeAspect[]) } : {}),
    ...(c.affectedSlides ? { affectedSlides: Object.freeze(c.affectedSlides.slice()) } : {}),
    ...(c.previousSlideIndex !== undefined ? { previousSlideIndex: c.previousSlideIndex } : {}),
  });
}

/** A copy of a location with only its documented fields. */
export function ownLocation(l: RefreshLocation): RefreshLocation {
  const out: Record<string, unknown> = {};
  for (const k of [
    'slideId',
    'slideIndex',
    'shapeId',
    'paragraphIndex',
    'start',
    'end',
    'text',
    'notes',
  ] as const)
    if (l[k] !== undefined) out[k] = l[k];
  return out as RefreshLocation;
}
