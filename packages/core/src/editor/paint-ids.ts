// Prefixes for the SVG ids of one paint target. Two editors on one page (also from two copies
// of this module), or an editor and an export, never share an id, so one's gradients and clip
// paths never apply to the other's slides.

let painters = 0;

/** A new id prefix, unique on the page. */
export function uniquePrefix(): string {
  painters += 1;
  return `pe${Math.random().toString(36).slice(2, 8)}${painters}`;
}
