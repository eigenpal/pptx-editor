// Slide pictures: PNG and standalone SVG from the same `paintSlide` output the editor shows.
//
// A picture must stand alone: a PNG is drawn from the SVG as an image, which loads nothing
// from the page, and an SVG file leaves the page. So pictures and fonts go inline. Pictures
// come from package parts only (never an external target). Fonts:
// - PNG: every face the slide uses is inlined in the SVG that is drawn, in memory only. The
//   PNG holds pixels, as the screen does; no font file leaves the page.
// - SVG: a face is inlined only when its license lets it travel with a document: the
//   metric-compatible substitutes, and embedded deck fonts whose `fsType` is installable or
//   editable. Other faces (preview-and-print deck fonts, host and local fonts) are referenced
//   by family name, so a viewer with the font installed shows it.

import { fontEmbeddingPermission, relTarget, writeZip } from '../store/index.ts';
import { elementsWithAttribute, paintSlide } from '../output/index.ts';
import type { RegisteredFace, SlideLayout } from '../layout/index.ts';
import type { DocumentSession } from './document-session.ts';
import { toBase64 } from '../automation/commands/base64.ts';
import { CommandRefusal } from '../automation/commands/op-batch.ts';
import { uniquePrefix } from './paint-ids.ts';
import { FACE_LOAD_WAIT_MS } from '../output/font-faces.ts';
import { within } from './within.ts';
import { FONT_NAME } from './selection-state.ts';

/** Options for a slide picture. @public */
export interface SlideImageOptions {
  /** Width in pixels, from 16 to 8192; the height follows the slide's aspect ratio. Default 1920. */
  readonly width?: number;
  /** The label of an element the editor draws as a placeholder (chart, diagram), as on the canvas. */
  readonly frameLabel?: (kind: string) => string;
  /**
   * PNG and ZIP export only: how long to wait for each slide's EMF and WMF pictures, in
   * milliseconds. Default 30000; `Infinity` waits until each one converted or failed.
   */
  readonly picturesTimeoutMs?: number;
}

const SVG_NS = 'http://www.w3.org/2000/svg';
const MAX_SIDE = 8192;
const FACE_ID = /^[A-Za-z0-9_-]{1,64}$/;

/** The pixel size of a picture of the session's slides. */
function pixelSize(session: DocumentSession, options: SlideImageOptions): [number, number] {
  const asked = Math.round(options.width ?? 1920);
  if (!Number.isFinite(asked) || asked < 16 || asked > MAX_SIDE)
    throw new CommandRefusal('errors.exportSize');
  const { width, height } = session.info;
  const h = Math.max(1, Math.round((asked * height) / width));
  if (h > MAX_SIDE) throw new CommandRefusal('errors.exportSize');
  return [asked, h];
}

/** The OS/2 `fsType` of an sfnt, or undefined when the table is missing or cut short. */
export function sfntFsType(bytes: Uint8Array): number | undefined {
  if (bytes.length < 12) return undefined;
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const tables = view.getUint16(4);
  for (let i = 0; i < tables; i += 1) {
    const record = 12 + i * 16;
    if (record + 16 > bytes.length) return undefined;
    if (view.getUint32(record) !== 0x4f532f32) continue; // 'OS/2'
    const offset = view.getUint32(record + 8);
    return offset + 10 <= bytes.length ? view.getUint16(offset + 8) : undefined;
  }
  return undefined;
}

/**
 * @internal Whether a face may be inlined in an SVG file that leaves the page: a substitute,
 * or an embedded face whose `fsType` allows installable or editable embedding.
 */
export function mayTravel(face: RegisteredFace): boolean {
  if (face.origin === 'substitute') return true;
  if (face.origin !== 'embedded') return false;
  const fsType = sfntFsType(face.bytes);
  if (fsType === undefined) return false;
  const permission = fontEmbeddingPermission(fsType);
  return permission === 'installable' || permission === 'editable';
}

/** `@font-face` for one face, with its bytes as a data URL. */
function fontFaceRule(face: RegisteredFace): string | null {
  // Face ids are the engine's own (`pptx-f1-0`), never a name from the file; check it anyway,
  // because it goes into CSS.
  if (!FACE_ID.test(face.id)) return null;
  const otf = face.bytes[0] === 0x4f && face.bytes[1] === 0x54 && face.bytes[2] === 0x54;
  const src = `url(data:${otf ? 'font/otf' : 'font/ttf'};base64,${toBase64(face.bytes)})`;
  return [
    '@font-face{font-family:"',
    face.id,
    '";font-weight:400;font-style:normal;src:',
    src,
    '}',
  ].join('');
}

/**
 * Paint slide `index` as an SVG that needs nothing from the page: pictures as data URLs, the
 * faces `inline` accepts as `@font-face` rules, and the family name after every other face.
 */
function paintAlone(
  session: DocumentSession,
  index: number,
  options: SlideImageOptions,
  inline: (face: RegisteredFace) => boolean
): { svg: SVGSVGElement; width: number; height: number } {
  if (typeof document === 'undefined') throw new CommandRefusal('errors.exportFailed', 'no DOM');
  const lay = session.layout(index);
  if (!lay) throw new CommandRefusal('errors.exportFailed', `slide ${index + 1}`);
  const [width, height] = pixelSize(session, options);
  const svg = paintLayout(session, lay, width, height, inline, options.frameLabel);
  return { svg, width, height };
}

/**
 * @internal Paint a laid-out slide at `width` by `height` pixels, standing alone (see
 * `paintAlone`), into `doc`: the page's document, or a DOM-free one on a server
 * (`createSvgDocument`).
 */
export function paintLayout(
  session: DocumentSession,
  lay: SlideLayout,
  width: number,
  height: number,
  inline: (face: RegisteredFace) => boolean,
  frameLabel?: (kind: string) => string,
  doc: Document = document
): SVGSVGElement {
  const pictures = new Map<string, string | undefined>();
  const svg = paintSlide(doc, lay, {
    images: (source, rId) => {
      const pkg = session.store.package;
      const target = relTarget(pkg, source, rId);
      if (!target) return undefined;
      if (!pictures.has(target)) {
        // The same pictures as the editor's paint: a damaged or oversized picture, and a
        // metafile that has not converted, export as the placeholder.
        const pic = session.picture(target);
        pictures.set(target, pic ? `data:${pic.mime};base64,${toBase64(pic.bytes)}` : undefined);
      }
      return pictures.get(target);
    },
    imageSize: session.imageSize,
    idPrefix: uniquePrefix(),
    hitTargets: false,
    ...(frameLabel ? { frameLabel } : {}),
  });
  svg.setAttribute('width', String(width));
  svg.setAttribute('height', String(height));
  const faces = new Map(session.faces.map((f) => [f.id, f]));
  const rules = new Map<string, string>();
  for (const el of elementsWithAttribute(svg, 'font-family')) {
    const list = el.getAttribute('font-family') ?? '';
    const face = faces.get(list.split(',')[0].trim());
    if (!face) continue;
    if (inline(face)) {
      const rule = rules.has(face.id) ? null : fontFaceRule(face);
      if (rule) rules.set(face.id, rule);
    } else if (FONT_NAME.test(face.family))
      el.setAttribute('font-family', `${face.id}, "${face.family}", sans-serif`);
  }
  if (rules.size) {
    const style = doc.createElementNS(SVG_NS, 'style');
    style.textContent = [...rules.values()].join('\n');
    svg.insertBefore(style, svg.firstChild);
  }
  return svg;
}

/**
 * Slide `index` as a standalone SVG document. See the file comment for the fonts it holds. An
 * EMF or WMF picture exports as the placeholder until it has converted: await
 * `session.picturesReady(index)` first. @public
 */
export function exportSlideSvg(
  session: DocumentSession,
  index: number,
  options: SlideImageOptions = {}
): string {
  const { svg } = paintAlone(session, index, options, mayTravel);
  return `<?xml version="1.0" encoding="UTF-8"?>\n${new XMLSerializer().serializeToString(svg)}`;
}

/** Load an SVG as an image and wait until it is decoded with its fonts and pictures. */
async function decodeSvg(text: string): Promise<HTMLImageElement> {
  const url = URL.createObjectURL(new Blob([text], { type: 'image/svg+xml' }));
  try {
    const img = new Image();
    const loaded = new Promise<void>((resolve, reject) => {
      img.onload = () => resolve();
      img.onerror = () => reject(new CommandRefusal('errors.exportFailed', 'decode'));
    });
    img.src = url;
    await loaded;
    await img.decode();
    return img;
  } finally {
    URL.revokeObjectURL(url);
  }
}

/** The bounded wait of each font set `ready` promise: one wait for each loading period. */
const settledWaits = new WeakMap<Promise<unknown>, Promise<unknown>>();

/**
 * The page's fonts, for a picture of a slide: settles when the browser's font set is ready, or
 * after the face load wait. A face whose load never settles keeps the set loading, and must not
 * hold an export. The set gives the same `ready` promise while it stays loading, so a ZIP of
 * every slide waits once, not once for each slide. @internal
 */
export function fontsSettled(): Promise<unknown> {
  const ready = typeof document === 'undefined' ? undefined : document.fonts?.ready;
  if (!ready) return Promise.resolve();
  let wait = settledWaits.get(ready);
  if (!wait) {
    wait = within(ready, FACE_LOAD_WAIT_MS, undefined);
    settledWaits.set(ready, wait);
  }
  return wait;
}

/**
 * Slide `index` as a PNG of `options.width` pixels (default 1920), drawn from the editor's own
 * paint. It waits for the slide's EMF and WMF pictures (`session.picturesReady`); one that is
 * not ready after `options.picturesTimeoutMs` (default 30 seconds) draws as the placeholder.
 * To know whether every picture draws, await `picturesReady` first.
 * Rejects with a `CommandRefusal` (`errors.exportSize`, `errors.exportFailed`). @public
 */
export async function exportSlidePng(
  session: DocumentSession,
  index: number,
  options: SlideImageOptions = {}
): Promise<Blob> {
  const timeoutMs = options.picturesTimeoutMs;
  await Promise.all([fontsSettled(), session.picturesReady(index, { timeoutMs })]);
  const { svg, width, height } = paintAlone(session, index, options, () => true);
  return drawPng(svg, width, height);
}

/** @internal Draw a standalone SVG on a canvas of `width` by `height` pixels as a PNG. */
export async function drawPng(svg: SVGSVGElement, width: number, height: number): Promise<Blob> {
  const img = await decodeSvg(new XMLSerializer().serializeToString(svg));
  const canvas = document.createElement('canvas');
  canvas.width = width;
  canvas.height = height;
  const ctx = canvas.getContext('2d');
  if (!ctx) throw new CommandRefusal('errors.exportFailed', 'canvas');
  ctx.drawImage(img, 0, 0, width, height);
  return new Promise<Blob>((resolve, reject) =>
    canvas.toBlob(
      (blob) => (blob ? resolve(blob) : reject(new CommandRefusal('errors.exportFailed', 'png'))),
      'image/png'
    )
  );
}

/**
 * Every slide as a PNG in one ZIP archive (`slide-01.png`, `slide-02.png`, …), written by the
 * package's bounded ZIP writer. @public
 */
export async function exportSlidesPngZip(
  session: DocumentSession,
  options: SlideImageOptions = {}
): Promise<Uint8Array> {
  const count = session.info.slides.length;
  const digits = Math.max(2, String(count).length);
  const entries = new Map<string, Uint8Array>();
  for (let i = 0; i < count; i += 1) {
    const png = await exportSlidePng(session, i, options);
    entries.set(
      `/slide-${String(i + 1).padStart(digits, '0')}.png`,
      new Uint8Array(await png.arrayBuffer())
    );
  }
  return writeZip(entries);
}
