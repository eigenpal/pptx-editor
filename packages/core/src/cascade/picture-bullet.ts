// Picture bullets: `a:buBlip` (ECMA-376 Part 1, 21.1.2.4.2). The `a:blip/@r:embed` of a level
// names a relationship of the part that holds that level: the presentation part for
// `p:defaultTextStyle`, the master for `p:txStyles` and master placeholders, the layout for
// layout placeholders, and the shape's own part for its list style and paragraphs. So the
// cascade records, for each chain entry, the part it came from.

import {
  binaryPartOf,
  child,
  elements,
  imageSize,
  imageSizeEmu,
  NS,
  relTarget,
  attr,
  type Element,
  type PptxPackage,
} from '../store/index.ts';

/** Where the picture of a picture bullet comes from. */
export interface PictureBulletSource {
  readonly pkg: PptxPackage;
  /** The part that holds a chain entry (`a:lvlNpPr`, `a:defPPr`, `a:pPr`), or undefined. */
  readonly partOf: (entry: Element) => string | undefined;
}

/** The picture of a resolved picture bullet. */
export interface PictureBulletImage {
  /** The part whose relationships resolve `rId` and `svgRId`. */
  readonly part: string;
  readonly rId: string;
  /** An `asvg:svgBlip` alternative, when present. */
  readonly svgRId?: string;
  /** The picture's width over its height. */
  readonly aspect: number;
}

/** The most pixels a bullet picture may have; the PDF writer's picture limit. */
export const MAX_BULLET_PICTURE_PIXELS = 25_000_000;

/**
 * The picture of an `a:buBlip` in `entry`, or undefined when its part is unknown, or its
 * relationship names no internal PNG, JPEG, GIF, or BMP part of a sane size. Undefined draws no
 * bullet: an external target is never fetched.
 */
export function pictureBulletImage(
  buBlip: Element,
  entry: Element,
  source: PictureBulletSource | undefined
): PictureBulletImage | undefined {
  const part = source?.partOf(entry);
  if (!source || !part) return undefined;
  const blip = child(buBlip, 'a:blip');
  const rId = attr(blip, 'r:embed');
  if (!rId) return undefined;
  const bytes = binaryPartOf(source.pkg, relTarget(source.pkg, part, rId));
  const size = bytes ? imageSize(bytes) : null;
  if (
    !size ||
    size.width <= 0 ||
    size.height <= 0 ||
    size.width * size.height > MAX_BULLET_PICTURE_PIXELS
  )
    return undefined;
  const emu = imageSizeEmu(size);
  const aspect = emu.cx / emu.cy;
  if (!(aspect > 0 && Number.isFinite(aspect))) return undefined;
  const svgRId = svgBlipOf(blip);
  return { part, rId, ...(svgRId ? { svgRId } : {}), aspect };
}

/** The `asvg:svgBlip/@r:embed` inside a blip's extension list. */
function svgBlipOf(blip: Element | undefined): string | undefined {
  for (const ext of elements(child(blip, 'a:extLst'))) {
    const svg = elements(ext).find((e) => e.ns === NS.asvg && e.local === 'svgBlip');
    const id = attr(svg, 'r:embed');
    if (id) return id;
  }
  return undefined;
}

/**
 * A `partOf` over lists of chain entry containers and their parts: an `a:lstStyle`,
 * `p:defaultTextStyle`, or `p:txStyles` child, and the part that holds it. `own` is the part of
 * the paragraphs themselves, for an `a:pPr`.
 */
export function entryParts(
  containers: readonly (readonly [Element | undefined, string | undefined])[],
  own: string | undefined
): (entry: Element) => string | undefined {
  return (entry) => {
    if (entry.local === 'pPr') return own;
    for (const [container, part] of containers)
      if (container && container.children.includes(entry)) return part;
    return undefined;
  };
}
