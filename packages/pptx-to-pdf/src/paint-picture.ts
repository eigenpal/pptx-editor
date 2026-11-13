/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pptx-to-pdf/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Pictures as PDF paint, from the same resolved picture fills the editor's SVG paint reads:
// each picture part embedded once, cropped by `srcRect`, placed by `fillRect`, tiled by
// `a:tile`, and recolored by its color effects.

import {
  closePath,
  concatTransformationMatrix,
  drawObject,
  lineTo,
  moveTo,
  popGraphicsState,
  pushGraphicsState,
  PDFContentStream,
  PDFName,
  PDFOperator,
  PDFOperatorNames,
  PDFBool,
  type PDFDocument,
  type PDFImage,
  PDFRawStream,
  type PDFRef,
} from 'pdf-lib';
import type { ResolvedFill } from '@pptx-editor.dev/core/cascade';
import {
  blipColorSteps,
  LUMA_WEIGHTS,
  metafilePng,
  tileGeometry,
  type BlipColorStep,
  type TileGeometry,
} from '@pptx-editor.dev/core/output';
import { binaryPartOf, imageSize, metafileFormat, relTarget } from '@pptx-editor.dev/core/store';
import { around, clip, fillEvenOdd, type PageOps } from './page-ops.ts';
import { fontEmbeddingDecision } from './pdf-font-embedding.ts';
import { drawSvgPicture, svgDrawingFor } from './paint-svg-picture.ts';
import type { SvgPictureDrawing } from '@pptx-editor.dev/core/editor';

/**
 * The largest picture this writer decodes. `pdf-lib` decodes every PNG pixel, so a small file
 * that declares a huge size could exhaust memory; the header is read first and checked here.
 */
export const MAX_IMAGE_PIXELS = 25_000_000;
const MAX_IMAGE_SIDE = 20_000;

type PictureFill = Extract<ResolvedFill, { type: 'picture' }>;

const PNG_SIGNATURE = [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a];
const isPng = (b: Uint8Array): boolean => PNG_SIGNATURE.every((v, i) => b[i] === v);
const isJpeg = (b: Uint8Array): boolean => b[0] === 0xff && b[1] === 0xd8 && b[2] === 0xff;

/** Embed a picture part once per document: the format from its signature, its size checked first. */
async function imageFor(target: string, p: PageOps): Promise<PDFImage | null> {
  const cached = p.res.images.get(target);
  if (cached !== undefined) return cached;
  let image: PDFImage | null = null;
  let bytes = binaryPartOf(p.res.pkg, target);
  // An EMF or WMF picture draws as its PNG, converted as in the editor.
  const metafile = bytes !== undefined && metafileFormat(bytes) !== undefined;
  if (bytes && metafile) {
    // Metafile text draws only with fonts that their embedding permission lets the PDF carry.
    const fonts = (p.res.metafileFonts ??= [...p.res.faces.values()]
      .filter((f) => {
        const face = {
          identity: f.id,
          bytes: f.bytes,
          faceIndex: 0,
          request: { family: f.family },
        };
        return fontEmbeddingDecision(face).kind !== 'refuse';
      })
      .map((f) => f.bytes));
    let png = await metafilePng(bytes, fonts, p.res.work.signal);
    // When the export's signal or deadline stopped the conversion, the export fails.
    p.res.work.check();
    // A picture that the converter's own time limit stopped tries once more (a busy machine),
    // then reports `picture.metafile`.
    if (png === undefined) png = await metafilePng(bytes, fonts, p.res.work.signal);
    p.res.work.check();
    bytes = png ?? undefined;
  }
  const png = bytes ? isPng(bytes) : false;
  if (!bytes && metafile)
    p.diag.add('picture.metafile', 'unsupported', 'An EMF or WMF picture does not convert.');
  else if (!bytes) p.diag.add('picture.missing', 'unsupported', 'A picture part is missing.');
  else if (!png && !isJpeg(bytes))
    p.diag.add('picture.format', 'unsupported', 'A picture is not PNG, JPEG, EMF, or WMF.');
  else {
    const size = imageSize(bytes);
    if (!size || size.width <= 0 || size.height <= 0)
      p.diag.add('picture.unreadable', 'unsupported', 'A picture header cannot be read.');
    else if (
      size.width > MAX_IMAGE_SIDE ||
      size.height > MAX_IMAGE_SIDE ||
      size.width * size.height > MAX_IMAGE_PIXELS
    )
      p.diag.add(
        'picture.too-large',
        'unsupported',
        `A ${size.width} × ${size.height} picture exceeds the ${MAX_IMAGE_PIXELS} pixel limit.`
      );
    else {
      // The picture's file bytes become PDF content once per document. pdf-lib keeps a PNG
      // decoded until the save, and only the pixel limit above bounds each one. DOCX Editor
      // passes a PNG through without decoding it.
      p.res.work.reserveContent(bytes.length);
      try {
        image = png ? await p.res.doc.embedPng(bytes) : await p.res.doc.embedJpg(bytes);
        await smoothImage(image, p);
      } catch {
        // An image that fails to embed is neither drawn nor saved: the save embeds every image
        // again, and the failed one would fail the whole export.
        if (image) forgetImage(p.res.doc, image);
        image = null;
        p.diag.add('picture.unreadable', 'unsupported', 'A picture cannot be decoded.');
      }
    }
  }
  p.res.images.set(target, image);
  return image;
}

/**
 * Remove an image from the list that `PDFDocument.save` embeds. `pdf-lib` keeps the list private
 * (`images`) and keeps a failed embed's rejected task, so the save would throw the same error.
 */
function forgetImage(doc: PDFDocument, image: PDFImage): void {
  const list = (doc as unknown as { images?: unknown }).images;
  if (!Array.isArray(list)) return;
  const at = list.indexOf(image);
  if (at >= 0) list.splice(at, 1);
}

/**
 * Mark an embedded picture for smooth sampling (`/Interpolate`, ISO 32000-1, 8.9.5.3), as the
 * editor's paint samples it. Without the flag, a PDF viewer can draw an enlarged picture with
 * hard pixel edges. The picture is embedded now, so its image dictionary exists.
 */
async function smoothImage(image: PDFImage, p: PageOps): Promise<void> {
  await image.embed();
  const ctx = p.res.doc.context;
  const dict = ctx.lookup(image.ref);
  if (!(dict instanceof PDFRawStream)) return;
  dict.dict.set(PDFName.of('Interpolate'), PDFBool.True);
  // A PNG with alpha keeps it in a soft mask image: smooth its edges too.
  const mask = ctx.lookup(dict.dict.get(PDFName.of('SMask')));
  if (mask instanceof PDFRawStream) mask.dict.set(PDFName.of('Interpolate'), PDFBool.True);
}

/** A rectangle path. */
function rectOps(x: number, y: number, w: number, h: number): PDFOperator[] {
  return [moveTo(x, y), lineTo(x + w, y), lineTo(x + w, y + h), lineTo(x, y + h), closePath()];
}

/**
 * At most this many color steps are drawn. Each step is about 150 bytes of function code that a
 * viewer runs for every pixel, and a file decides how many effects a picture has. A picture with
 * each effect once takes 7 steps.
 */
export const MAX_COLOR_STEPS = 64;

/** A number for a PostScript calculator function: plain decimal notation, bounded. */
function psNumber(v: number): string {
  const x = Number.isFinite(v) ? Math.max(-1e6, Math.min(1e6, v)) : 0;
  const s = x.toFixed(6).replace(/\.?0+$/, '');
  return s === '-0' || s === '' ? '0' : s;
}

/**
 * The PostScript calculator code (ISO 32000-1, 7.10.5) of color steps, from `inputs` (1 or 3)
 * components in [0, 1] to red, green, and blue: each step as the SVG filter primitive computes
 * it, clamped to [0, 1].
 */
export function colorStepsCode(steps: readonly BlipColorStep[], inputs: 1 | 3): string {
  const clamp = 'dup 0 lt { pop 0 } if dup 1 gt { pop 1 } if';
  const code: string[] = inputs === 1 ? ['dup dup'] : [];
  const [wr, wg, wb] = LUMA_WEIGHTS.map(psNumber);
  for (const s of steps) {
    if (s.op === 'gray') {
      // r g b → r·wr + g·wg + b·wb, three times.
      code.push(`${wb} mul exch ${wg} mul add exch ${wr} mul add ${clamp} dup dup`);
      continue;
    }
    // Blue, then green, then red: each channel on the top, mapped, and rolled under the others.
    for (const i of [2, 1, 0]) {
      const slope = s.op === 'table' ? s.to[i]! - s.from[i]! : s.slope;
      const intercept = s.op === 'table' ? s.from[i]! : s.intercept;
      code.push(`${psNumber(slope)} mul ${psNumber(intercept)} add ${clamp} 3 1 roll`);
    }
  }
  return `{ ${code.join(' ')} }`;
}

/**
 * The image with color steps applied: a copy of its image dictionary and its encoded data, in a
 * `DeviceN` color space whose tint transform is the steps' PostScript function. A viewer maps each
 * decoded pixel through the function, so the export never decodes a picture for its effects, and
 * a JPEG stays a JPEG. The alpha (`SMask`) is the original's. Written once per image and steps.
 * An image in a color space other than RGB and gray (a CMYK JPEG) is drawn without the effects.
 */
function recolored(image: PDFImage, allSteps: readonly BlipColorStep[], p: PageOps): PDFRef {
  if (allSteps.length === 0) return image.ref;
  if (allSteps.length > MAX_COLOR_STEPS)
    p.diag.add(
      'picture.color-effects',
      'approximation',
      `Pictures with more than ${MAX_COLOR_STEPS} color effect steps are drawn with the first ${MAX_COLOR_STEPS}.`
    );
  const steps = allSteps.slice(0, MAX_COLOR_STEPS);
  const key = `${image.ref.toString()}|${JSON.stringify(steps)}`;
  const cache = (p.res.recolored ??= new Map());
  const known = cache.get(key);
  const ctx = p.res.doc.context;
  const source = ctx.lookup(image.ref);
  const space = source instanceof PDFRawStream ? source.dict.get(PDFName.of('ColorSpace')) : null;
  const inputs = space === PDFName.of('DeviceRGB') ? 3 : space === PDFName.of('DeviceGray') ? 1 : 0;
  if (known === null || !(source instanceof PDFRawStream) || inputs === 0) {
    // Reported on each slide that draws the picture.
    p.diag.add(
      'picture.color-effects',
      'approximation',
      'Picture color effects on a CMYK picture are not drawn.'
    );
    cache.set(key, null);
    return image.ref;
  }
  if (known) return known;
  const code = colorStepsCode(steps, inputs);
  const fn = ctx.register(
    ctx.stream(code, {
      FunctionType: 4,
      Domain: Array.from({ length: inputs }, () => [0, 1]).flat(),
      Range: [0, 1, 0, 1, 0, 1],
    })
  );
  const names = inputs === 3 ? ['PptxRed', 'PptxGreen', 'PptxBlue'] : ['PptxGray'];
  const dict = source.dict.clone(ctx);
  dict.set(
    PDFName.of('ColorSpace'),
    ctx.obj(['DeviceN', names.map((n) => PDFName.of(n)), 'DeviceRGB', fn])
  );
  // The encoded picture is written a second time.
  p.res.work.reserveContent(source.contents.length + code.length + dict.sizeInBytes());
  const ref = ctx.register(PDFRawStream.of(dict, source.contents));
  cache.set(key, ref);
  return ref;
}

/**
 * The tiling pattern of a tiled picture (ISO 32000-1, 8.7.3.3): one cell of 1 tile, or 2 or 4
 * tiles mirrored by `flip`, with the tile geometry of the SVG paint (`core/output`,
 * `tileGeometry`). The pattern space is the box's, with the grid's origin at `g.x`, `g.y`. Written
 * once per picture and geometry.
 */
function tilePattern(image: PDFRef, g: TileGeometry, p: PageOps): PDFRef {
  const { width: tw, height: th, flipX: fx, flipY: fy } = g;
  const cw = fx ? 2 * tw : tw;
  const ch = fy ? 2 * th : th;
  const key = `tile|${image.toString()}|${[g.x, g.y, tw, th, fx, fy].join(',')}`;
  let ref = p.res.patterns.get(key);
  if (ref) return ref;
  const ctx = p.res.doc.context;
  // The image space is the unit square with y up; a tile is tw × th with y down.
  const placements: number[][] = [[tw, 0, 0, -th, 0, th]];
  if (fx) placements.push([-tw, 0, 0, -th, 2 * tw, th]);
  if (fy) placements.push([tw, 0, 0, th, 0, th]);
  if (fx && fy) placements.push([-tw, 0, 0, th, 2 * tw, th]);
  const ops: PDFOperator[] = [];
  for (const m of placements)
    ops.push(
      pushGraphicsState(),
      concatTransformationMatrix(m[0]!, m[1]!, m[2]!, m[3]!, m[4]!, m[5]!),
      drawObject('Im'),
      popGraphicsState()
    );
  const dict = ctx.obj({
    Type: 'Pattern',
    PatternType: 1,
    PaintType: 1,
    TilingType: 1,
    BBox: [0, 0, cw, ch],
    XStep: cw,
    YStep: ch,
    Matrix: [1, 0, 0, 1, g.x, g.y],
    Resources: ctx.obj({ XObject: ctx.obj({ Im: image }) }),
  });
  let bytes = dict.sizeInBytes();
  for (const op of ops) bytes += op.sizeInBytes() + 1;
  p.res.work.count(ops.length);
  p.res.work.reserveContent(bytes);
  ref = ctx.register(PDFContentStream.of(dict, ops, false));
  p.res.patterns.set(key, ref);
  return ref;
}

/** Whether a picture fill has an image this writer draws (the image is embedded once). */
export async function canDrawPicture(
  sourcePart: string,
  fill: PictureFill,
  p: PageOps
): Promise<boolean> {
  if (fill.svgRId && svgDrawing(sourcePart, fill, p)) return true;
  const target = relTarget(p.res.pkg, sourcePart, fill.rId);
  return target ? (await imageFor(target, p)) !== null : false;
}

/** The vector drawing of a picture fill's SVG picture, or null when the editor draws its PNG. */
function svgDrawing(sourcePart: string, fill: PictureFill, p: PageOps): SvgPictureDrawing | null {
  const target = fill.svgRId ? relTarget(p.res.pkg, sourcePart, fill.svgRId) : undefined;
  return target ? svgDrawingFor(target, p) : null;
}

/**
 * Where the whole picture sits in a `w` × `h` box: `fillRect` places the picture's visible part,
 * and `srcRect` crops it, so the whole picture extends past that place by the crop.
 */
function pictureBox(
  fill: PictureFill,
  w: number,
  h: number
): { x: number; y: number; w: number; h: number } {
  const fr = fill.fillRect;
  const bx = fr.l * w;
  const by = fr.t * h;
  const bw = w * (1 - fr.l - fr.r);
  const bh = h * (1 - fr.t - fr.b);
  // A crop that leaves nothing visible on an axis (two concurrent crops merged edge by edge)
  // draws uncropped on that axis, as the SVG paint and the crop commands read it.
  const raw = fill.srcRect;
  const okW = 1 - raw.l - raw.r >= 0.001;
  const okH = 1 - raw.t - raw.b >= 0.001;
  const sr = {
    l: okW ? raw.l : 0,
    r: okW ? raw.r : 0,
    t: okH ? raw.t : 0,
    b: okH ? raw.b : 0,
  };
  const iw = bw / (1 - sr.l - sr.r);
  const ih = bh / (1 - sr.t - sr.b);
  return { x: bx - sr.l * iw, y: by - sr.t * ih, w: iw, h: ih };
}

/**
 * Draw a picture fill over a `w` × `h` box, as the SVG `pictureImage`: `fillRect` places the
 * picture in the box, and `srcRect` crops it, so the picture's visible part fills that place.
 * The caller clips to the shape.
 */
export async function drawPicture(
  sourcePart: string,
  fill: PictureFill,
  w: number,
  h: number,
  p: PageOps
): Promise<void> {
  // The SVG paint draws no other effect either, so the export reports them as the editor shows.
  const unsupported = fill.effects.find((e) => e.kind === 'unsupported');
  if (unsupported)
    p.diag.add(
      'picture.color-effects',
      'approximation',
      `The picture color effect ${clip(unsupported.name)} is not drawn.`
    );
  const alpha = fill.effects.reduce((a, e) => (e.kind === 'alphaModFix' ? a * e.amount : a), 1);
  // An SVG picture that the editor draws as vector art draws as PDF paths, without color
  // effects or tiles. Any other one draws the deck's own raster picture of the SVG, so it is
  // reported, not refused.
  if (fill.svgRId) {
    const drawing = svgDrawing(sourcePart, fill, p);
    if (drawing && fill.mode !== 'tile' && blipColorSteps(fill.effects).length === 0) {
      const r = pictureBox(fill, w, h);
      drawSvgPicture(drawing, r.x, r.y, r.w, r.h, alpha, p);
      return;
    }
    p.diag.add(
      'picture.svg-fallback',
      'information',
      drawing
        ? 'SVG pictures with color effects or tiles are drawn from their PNG fallback.'
        : 'SVG pictures outside the vector subset are drawn from their PNG fallback.'
    );
  }
  const target = relTarget(p.res.pkg, sourcePart, fill.rId);
  if (!target) {
    p.diag.add('picture.missing', 'unsupported', 'A picture part is missing or external.');
    return;
  }
  const image = await imageFor(target, p);
  if (!image) return;
  const ref = recolored(image, blipColorSteps(fill.effects), p);
  if (fill.mode === 'tile' && fill.tile) {
    // The tiles sit on the box, so a pattern in the space of a form drawn at the box's matrix.
    const g = tileGeometry(fill.tile, p.res.imageSize?.(sourcePart, fill.rId), w, h);
    const pattern = p.name('Pattern', tilePattern(ref, g, p));
    const bbox = around(w, h);
    const form = p.form(
      [
        PDFOperator.of(PDFOperatorNames.NonStrokingColorspace, [PDFName.of('Pattern')]),
        PDFOperator.of(PDFOperatorNames.NonStrokingColorN, [pattern]),
        ...rectOps(bbox[0], bbox[1], bbox[2] - bbox[0], bbox[3] - bbox[1]),
        fillEvenOdd(),
      ],
      bbox
    );
    p.push(pushGraphicsState());
    p.alpha(alpha, 1);
    p.push(drawObject(p.name('XObject', form)), popGraphicsState());
    return;
  }
  const r = pictureBox(fill, w, h);
  const name = p.name('XObject', ref);
  p.push(pushGraphicsState());
  p.alpha(alpha, 1);
  // The image space is the unit square with y up; the picture is iw × ih EMU with y down.
  p.push(concatTransformationMatrix(r.w, 0, 0, -r.h, r.x, r.y + r.h));
  p.push(drawObject(name), popGraphicsState());
}
