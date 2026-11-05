// Office.js `Presentation.insertSlidesFromBase64`: slides of another `.pptx` into this one.
//
// The other file is untrusted bytes. Its base64 text is capped before it is decoded, and the
// bytes go through `readPptx` with its zip, size, ratio, part name, and XML limits, as any file
// this host opens. The copy is `commands/slide-import.ts`, in the same batch as every other
// operation of the request.

import { elements, NS, presentationInfo, readPptx, xmlPartOf, attr } from '../store/index.ts';
import { fromBase64 } from './commands/base64.ts';
import { importSlides, type ImportFormatting } from './commands/slide-import.ts';
import { normalizePresentation } from './commands/slide-sections.ts';
import type { AutomationInsertSlidesOptions } from './operations.ts';
import type { AutomationValue } from './protocol.ts';
import { AutomationFailure, edit, type PlanContext } from './targets.ts';

/** The largest presentation an import reads, in bytes. */
export const MAX_IMPORT_BYTES = 100 * 1024 * 1024;
/** Its base64 text: 4 characters for 3 bytes, with room for line breaks. */
const MAX_IMPORT_BASE64 = Math.ceil(MAX_IMPORT_BYTES / 3) * 4 * 1.1;
/** The most slide ids one import names. */
const MAX_IDS = 10_000;

const FORMATTING: Readonly<Record<string, ImportFormatting>> = Object.assign(Object.create(null), {
  keepSource: 'keepSource',
  useDestinationTheme: 'useDestinationTheme',
});

/** The `p14:creationId` of a slide part (the `p14` extension), or undefined. */
function creationIdOf(pkg: Parameters<typeof xmlPartOf>[0], part: string): string | undefined {
  const ext = elements(xmlPartOf(pkg, part)?.root).find((e) => e.local === 'extLst');
  for (const e of elements(ext))
    for (const c of elements(e))
      if (c.ns === NS.p14 && c.local === 'creationId') return attr(c, 'val');
  return undefined;
}

/**
 * The index of the slide a slide id names: `nnn` or `nnn#` (`p:sldId/@id`), `#ccc`
 * (`p14:creationId`), or `nnn#ccc` (both must match).
 */
function slideIndex(
  pkg: Parameters<typeof xmlPartOf>[0],
  key: unknown,
  name: string,
  detail: string
): number {
  if (typeof key !== 'string' || key.length === 0 || key.length > 64)
    throw new AutomationFailure('invalid-argument', `invalid ${name}`, name);
  const m = /^(\d*)(?:#(\d*))?$/.exec(key.trim());
  if (!m || (!m[1] && !m[2]))
    throw new AutomationFailure('invalid-argument', `invalid ${name}`, name);
  const index = presentationInfo(pkg).slides.findIndex(
    (s) =>
      (!m[1] || Number(s.sldId) === Number(m[1])) &&
      (!m[2] || Number(creationIdOf(pkg, s.partName)) === Number(m[2]))
  );
  if (index < 0) throw new AutomationFailure('item-not-found', 'no such slide', detail);
  return index;
}

export function insertSlides(
  ctx: PlanContext,
  base64: string,
  options: AutomationInsertSlidesOptions = {}
): AutomationValue {
  if (typeof options !== 'object' || options === null)
    throw new AutomationFailure('invalid-argument', 'invalid options', 'options');
  if (typeof base64 !== 'string' || base64.length === 0 || base64.length > MAX_IMPORT_BASE64)
    throw new AutomationFailure('invalid-argument', 'invalid base64', 'import-base64');
  const formatting = FORMATTING[options.formatting ?? 'keepSource'];
  if (!formatting)
    throw new AutomationFailure('invalid-argument', 'invalid formatting', 'formatting');
  const bytes = fromBase64(base64.replace(/\s+/g, ''));
  if (!bytes || bytes.length === 0 || bytes.length > MAX_IMPORT_BYTES)
    throw new AutomationFailure('invalid-argument', 'invalid base64', 'import-base64');
  const read = readPptx(bytes);
  if (!read.ok)
    throw new AutomationFailure(
      'invalid-argument',
      'the file is not a presentation this host opens',
      `import-package:${read.error.reason}`
    );
  const source = read.pkg;
  const ids = options.sourceSlideIds;
  if (ids !== undefined && (!Array.isArray(ids) || ids.length > MAX_IDS))
    throw new AutomationFailure('invalid-argument', 'invalid sourceSlideIds', 'sourceSlideIds');
  const slides = ids
    ? ids.map((id) => slideIndex(source, id, 'sourceSlideIds', 'import-slide'))
    : presentationInfo(source).slides.map((_, i) => i);
  if (new Set(slides).size !== slides.length)
    throw new AutomationFailure('invalid-argument', 'a slide is named twice', 'sourceSlideIds');
  const after =
    options.targetSlideId === undefined
      ? -1
      : slideIndex(ctx.batch.pkg, options.targetSlideId, 'targetSlideId', 'import-target');
  if (slides.length === 0) return { kind: 'applied' };
  edit(() => {
    normalizePresentation(ctx.batch);
    importSlides(ctx.batch, { source, slides, after, formatting });
  });
  return { kind: 'applied' };
}
