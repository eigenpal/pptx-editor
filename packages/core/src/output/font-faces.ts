// Register measured faces with the page's font set under their internal ids, so paint draws
// the same bytes HarfBuzz measured: one font per measurement and paint. Uses the FontFace API: no CSS text is
// built from file data.

import type { RegisteredFace } from '../layout/index.ts';

/** Registered faces, removable when the document closes. */
export interface FontFaceHandle {
  readonly ready: Promise<void>;
  dispose(): void;
}

/** Options of `registerFontFaces`. */
export interface RegisterFontFacesOptions {
  /**
   * The family name of each face in the font set: the face's id when not given. A host that
   * paints under other names (and rewrites the painted family names) gives them here.
   */
  readonly family?: (face: RegisteredFace) => string;
}

/** Add faces to `doc.fonts`. Bytes are copied into the FontFace; nothing touches disk. */
export function registerFontFaces(
  doc: Document,
  faces: readonly RegisteredFace[],
  options: RegisterFontFacesOptions = {}
): FontFaceHandle {
  const added: FontFace[] = [];
  const loads: Promise<unknown>[] = [];
  // A DOM without the CSS Font Loading API (a test DOM, an old embedded view) draws with the
  // host's own fonts, as DOCX Editor's `embedded-font-faces.ts` does. Layout is unchanged.
  // The constructor comes from the document's own window: another window's font set refuses
  // a FontFace made by this one.
  const FontFaceCtor = ((doc.defaultView ?? globalThis) as { FontFace?: typeof FontFace }).FontFace;
  if (typeof FontFaceCtor !== 'function' || !doc.fonts)
    return { ready: Promise.resolve(), dispose: () => {} };
  for (const f of faces) {
    // Every face is the only member of its family and is declared regular, so the page draws
    // it as-is; paint adds font-weight/font-style only to synthesize a missing style.
    const family = options.family?.(f) ?? f.id;
    const face = new FontFaceCtor(family, f.bytes.slice().buffer, {
      weight: '400',
      style: 'normal',
    });
    loads.push(face.load().catch(() => undefined));
    doc.fonts.add(face);
    added.push(face);
  }
  return {
    ready: Promise.all(loads).then(() => undefined),
    dispose: () => {
      for (const face of added) doc.fonts.delete(face);
    },
  };
}
