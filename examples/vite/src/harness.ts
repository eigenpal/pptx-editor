// The fidelity harness page. Playwright loads it, hands it deck bytes, and prints the slides to
// PDF. It paints through the SAME session, layout, and paint the editor uses: there is no second
// renderer for tests.

import { openPresentation, type DocumentSession } from '@pptx-editor.dev/core/editor';
import { paintSlide, registerFontFaces } from '@pptx-editor.dev/core/output';
import { writePptx } from '@pptx-editor.dev/core/store';
import type { FontSource, LayoutItem } from '@pptx-editor.dev/core/layout';

type RegisteredFace = DocumentSession['faces'][number];
import { installDrawing } from '@pptx-editor.dev/pro/drawing';

// The harness renders as the demos do: with the Pro drawing renderers (charts, glow, inner shadow,
// reflection, and soft edge), installed before any deck opens.
installDrawing();

/**
 * A font file on the harness server, with the SHA-256 of its bytes. A page that holds bytes with
 * that digest from an earlier deck uses them, and does not fetch the file again.
 */
interface FontFile {
  url: string;
  sha256?: string;
}

/** A deck and its local fonts, as paths on the harness server (`openDeck`). */
interface DeckSource {
  deck: string;
  fonts: ({ family: string; bold: boolean; italic: boolean } & FontFile)[];
  /** The fallback faces, in order (`FALLBACK_FONTS` in scripts/fidelity/local-fonts.ts). */
  fallback?: FontFile[];
}

/** The size and SHA-256 of a file that the page got. */
interface ReceivedFile {
  bytes: number;
  sha256: string;
}

/** The pages of a PDF to draw (`rasterPdf`), each with the paths that take its pixels and PNG. */
interface RasterRequest {
  pdf: string;
  width: number;
  height: number;
  scale: number;
  /**
   * pdf.js options from the harness server: the image limit, and the CMap, WebAssembly, and ICC
   * profile folders.
   */
  maxImageSize: number;
  cMapUrl: string;
  wasmUrl: string;
  iccUrl: string;
  pages: { rgba: string; png: string }[];
}

interface HarnessApi {
  /** Fetch the deck and its fonts from the harness server, and open the deck. */
  open(source: DeckSource): Promise<{
    slides: number;
    error?: string;
    received: { deck: ReceivedFile; fonts: (ReceivedFile & { family: string })[] };
  }>;
  /** Paint one slide, and return its diagnostics, requested fonts, and markup (`markupOf`). */
  render(index: number): Promise<{
    diagnostics: readonly string[];
    fonts: string[];
    markup: string;
  }>;
  /** Paint the slides at `indexes` as print pages, one page each, at the slide size. */
  print(indexes: number[]): Promise<void>;
  /** Remove the print pages. */
  endPrint(): void;
  /** Draw each PDF page with pdf.js and send its pixels (RGBA, gzip) and PNG to the server. */
  rasterPdf(request: RasterRequest): Promise<void>;
  reopenSaved(): Promise<{ slides: number; error?: string }>;
}

let session: DocumentSession | null = null;
let fonts: FontSource[] = [];
/** The registered fallback faces, and the family list that replaces `sans-serif` (`paintAt`). */
let fallbackFaces: FontFace[] = [];
let fallbackFamilies = '';
/** The digests of the fallback faces, in order: the same list keeps the same faces. */
let fallbackKey = '';
/**
 * The page's font faces, loaded once for each distinct font file: the bytes of each fetched
 * font by its SHA-256, the SHA-256 of each face's bytes, and the digests registered: one
 * `FontFace` for each, under the family `pptx-h-<sha256>`. Every deck of a lane that uses a font reuses them, so a
 * large face (Noto Sans CJK JP, 16 MB) is fetched, hashed, and decoded once for each page, not
 * for each deck and each reopen. A face of one digest draws the same in every deck.
 */
const fontBytes = new Map<string, Uint8Array<ArrayBuffer>>();
const digestOf = new WeakMap<Uint8Array, string>();
const pageFaces = new Set<string>();
/** The page family of each face id of the open session (`pptx-f<measurer>-<face>`). */
let faceFamilies = new Map<string, string>();
/** The type and SHA-256 of each image blob, by its URL (`markupOf`). */
const blobDigests = new Map<string, Promise<string>>();

/** Fetch a file that the harness server gives to this page. */
async function fetchBytes(url: string, what: string): Promise<Uint8Array<ArrayBuffer>> {
  let r: Response;
  try {
    r = await fetch(url, { cache: 'no-store' });
    if (r.ok) return new Uint8Array(await r.arrayBuffer());
  } catch (e) {
    throw new Error(`${what} did not arrive from the harness server: ${(e as Error).message}`);
  }
  throw new Error(`${what} did not arrive from the harness server (HTTP ${r.status})`);
}

async function sha256Of(bytes: BufferSource): Promise<string> {
  const digest = new Uint8Array(await crypto.subtle.digest('SHA-256', bytes));
  return [...digest].map((b) => b.toString(16).padStart(2, '0')).join('');
}

/** The SHA-256 of face or font bytes, computed once for each array. */
async function digestOfBytes(bytes: Uint8Array): Promise<string> {
  let digest = digestOf.get(bytes);
  if (!digest) {
    digest = await sha256Of(bytes.slice());
    digestOf.set(bytes, digest);
  }
  return digest;
}

async function received(bytes: Uint8Array<ArrayBuffer>): Promise<ReceivedFile> {
  return { bytes: bytes.length, sha256: await digestOfBytes(bytes) };
}

/** The bytes of a font file: the ones this page already holds with its digest, or a fetch. */
async function fontFile(file: FontFile, what: string): Promise<Uint8Array<ArrayBuffer>> {
  const held = file.sha256 ? fontBytes.get(file.sha256) : undefined;
  if (held) return held;
  const bytes = await fetchBytes(file.url, what);
  fontBytes.set(await digestOfBytes(bytes), bytes);
  return bytes;
}

/**
 * Give each face of `list` its page family, and register (with the engine's
 * `registerFontFaces`, under that family) each face whose bytes the page has not registered yet.
 */
async function registerFaces(list: readonly RegisteredFace[]): Promise<Map<string, string>> {
  const families = new Map<string, string>();
  const fresh: RegisteredFace[] = [];
  for (const f of list) {
    const digest = await digestOfBytes(f.bytes);
    families.set(f.id, `pptx-h-${digest}`);
    if (pageFaces.has(digest)) continue;
    pageFaces.add(digest);
    fresh.push(f);
  }
  await registerFontFaces(document, fresh, { family: (f) => families.get(f.id)! }).ready;
  return families;
}

/**
 * The painted SVG with every session-bound name replaced by the digest of the bytes behind it:
 * each image blob URL (with its type), and each face id (`pptx-f<measurer>-<face>`, new for each session). Two
 * paints with the same markup draw the same bytes with the same instructions, so the save check
 * compares the markup alone.
 */
async function markupOf(svg: SVGSVGElement): Promise<string> {
  const urls = new Set<string>();
  for (const img of svg.querySelectorAll('image')) {
    const href = img.getAttribute('href');
    if (href?.startsWith('blob:')) urls.add(href);
  }
  const digests = new Map<string, string>();
  for (const url of urls) {
    let digest = blobDigests.get(url);
    if (!digest) {
      // The type with the bytes: the page decodes a blob by its type (an SVG image needs it).
      digest = fetch(url).then(
        async (r) => `${r.headers.get('content-type')}:${await sha256Of(await r.arrayBuffer())}`
      );
      blobDigests.set(url, digest);
    }
    digests.set(url, await digest);
  }
  let markup = new XMLSerializer()
    .serializeToString(svg)
    .replace(/pptx-h-([0-9a-f]{64})/g, (_, digest: string) => `face:${digest}`)
    .replace(/pptx-f\d+-\d+/g, (id) => `face:${id}`);
  for (const [url, digest] of digests) markup = markup.split(url).join(`image:${digest}`);
  return markup;
}
const EMU_PER_PT = 12_700;

type Layout = NonNullable<ReturnType<DocumentSession['layout']>>;

/**
 * Paint a slide layout as an SVG at the slide size in points. Where the paint names the generic
 * `sans-serif` family (a frame label, and the glyphs that a run's face lacks), the harness names
 * its open fallback faces instead, so no system font takes part. A glyph that they also lack
 * still takes a system font, and the fidelity check reports that font.
 */
function paintAt(layout: Layout, idPrefix: string): SVGSVGElement {
  const svg = paintSlide(document, layout, {
    images: session!.imageUrl,
    imageSize: session!.imageSize,
    idPrefix,
  });
  // Each face id of the session names the page's face of the same bytes (`registerFaces`).
  for (const el of svg.querySelectorAll('[font-family]')) {
    let family = (el.getAttribute('font-family') ?? '').replace(
      /pptx-f\d+-\d+/g,
      (id) => faceFamilies.get(id) ?? id
    );
    if (fallbackFamilies) family = family.replace(/(^|,\s*)sans-serif$/, `$1${fallbackFamilies}`);
    el.setAttribute('font-family', family);
  }
  svg.setAttribute('width', `${layout.slide.width / EMU_PER_PT}pt`);
  svg.setAttribute('height', `${layout.slide.height / EMU_PER_PT}pt`);
  return svg;
}

/** Wait for the fonts and for every image of the SVGs. */
async function loaded(svgs: readonly SVGSVGElement[]): Promise<void> {
  await document.fonts.ready;
  await Promise.all(
    svgs.flatMap((svg) =>
      [...svg.querySelectorAll('image')].map(
        (img) =>
          new Promise<void>((done) => {
            const probe = new Image();
            probe.onload = () => done();
            probe.onerror = () => done();
            probe.src = img.getAttribute('href') ?? '';
          })
      )
    )
  );
}

/** The families that the slide's text asks for (the requested Latin face of each text run). */
function collectFonts(items: readonly LayoutItem[], out: Set<string>): void {
  for (const it of items) {
    if (it.kind === 'group') collectFonts(it.children, out);
    if (it.kind !== 'shape' || !it.text) continue;
    for (const line of it.text.lines)
      for (const run of line.runs)
        if (run.kind === 'text' && run.text.trim() && run.props.latin) out.add(run.props.latin);
  }
}

async function adopt(next: DocumentSession): Promise<void> {
  session?.dispose();
  session = next;
  faceFamilies = await registerFaces(session.faces);
  // EMF and WMF pictures paint as their PNG, as in the editor once they have converted.
  await session.picturesReady();
}

const api: HarnessApi = {
  async open(source) {
    // No slide of the last deck may render, or reopen, with this deck's fonts, also when a
    // fetch below fails.
    session?.dispose();
    session = null;
    fonts = [];
    faceFamilies = new Map();
    blobDigests.clear();
    // The same fallback files keep their faces; another list replaces them.
    const fallback = source.fallback ?? [];
    const key = fallback.map((f) => f.sha256 ?? f.url).join('\u0000');
    if (key !== fallbackKey || fallback.some((f) => !f.sha256)) {
      for (const face of fallbackFaces) document.fonts.delete(face);
      fallbackFaces = [];
      fallbackFamilies = '';
      fallbackKey = '';
      for (const [i, file] of fallback.entries()) {
        const face = new FontFace(
          `pptx-harness-fallback-${i}`,
          await fontFile(file, 'a fallback font')
        );
        document.fonts.add(await face.load());
        fallbackFaces.push(face);
      }
      fallbackFamilies = fallbackFaces.map((f) => f.family).join(', ');
      fallbackKey = key;
    }
    const deck = await fetchBytes(source.deck, 'the deck');
    const got = { deck: await received(deck), fonts: [] as (ReceivedFile & { family: string })[] };
    for (const f of source.fonts) {
      const bytes = await fontFile(f, `the font ${f.family}`);
      got.fonts.push({ family: f.family, ...(await received(bytes)) });
      fonts.push({ family: f.family, bold: f.bold, italic: f.italic, bytes, origin: 'local' });
    }
    const r = await openPresentation(deck, { fonts });
    if (!r.ok) return { slides: 0, error: r.error.reason, received: got };
    await adopt(r.session);
    return { slides: r.session.info.slides.length, received: got };
  },
  async render(index) {
    if (!session) throw new Error('no document');
    const layout = session.layout(index);
    if (!layout) throw new Error(`no slide ${index}`);
    // A picture that did not convert paints as the placeholder: the slide's report says so.
    const pictures = await session.picturesReady(index);
    const svg = paintAt(layout, `s${index}`);
    document.getElementById('slide')!.replaceChildren(svg);
    await loaded([svg]);
    const fonts = new Set<string>();
    collectFonts(layout.items, fonts);
    return {
      diagnostics: pictures
        ? layout.diagnostics
        : [...layout.diagnostics, 'picture not converted: emf-wmf'],
      fonts: [...fonts].sort(),
      markup: await markupOf(svg),
    };
  },
  async print(indexes) {
    if (!session) throw new Error('no document');
    api.endPrint();
    const pages = document.createElement('div');
    pages.id = 'print';
    const svgs: SVGSVGElement[] = [];
    let size = { width: 0, height: 0 };
    for (const index of indexes) {
      const layout = session.layout(index);
      if (!layout) throw new Error(`no slide ${index}`);
      size = { width: layout.slide.width / EMU_PER_PT, height: layout.slide.height / EMU_PER_PT };
      const svg = paintAt(layout, `p${index}`);
      const page = document.createElement('div');
      page.style.cssText = `width:${size.width}pt;height:${size.height}pt;overflow:hidden;break-after:page`;
      page.append(svg);
      pages.append(page);
      svgs.push(svg);
    }
    const style = document.createElement('style');
    style.id = 'print-style';
    // Every slide of a deck has one size, so one page size serves all pages.
    style.textContent = `@page { size: ${size.width}pt ${size.height}pt; margin: 0 } @media print { #slide { display: none } }`;
    document.head.append(style);
    document.body.append(pages);
    await loaded(svgs);
  },
  endPrint() {
    document.getElementById('print')?.remove();
    document.getElementById('print-style')?.remove();
  },
  async rasterPdf(r) {
    // Loaded on first use, so the editor demo never loads pdf.js.
    const pdfjs = await import('pdfjs-dist');
    pdfjs.GlobalWorkerOptions.workerSrc = (
      await import('pdfjs-dist/build/pdf.worker.mjs?url')
    ).default;
    const doc = await pdfjs.getDocument({
      data: await fetchBytes(r.pdf, 'the PDF'),
      // Glyphs as paths: no system font rasterizer takes part. A PDF with a font that is not
      // embedded never comes here (`score-deck.ts`), so no standard font data is needed.
      disableFontFace: true,
      maxImageSize: r.maxImageSize,
      cMapUrl: r.cMapUrl,
      cMapPacked: true,
      // pdf.js converts ICC-based colors with its WebAssembly color engine; without it, an
      // ICC-based color draws as the device color with the same numbers.
      wasmUrl: r.wasmUrl,
      iccUrl: r.iccUrl,
      // pdf.js turns worker fetches, and with them the color engine, off unless every data URL
      // is set; this page sets no standard font URL.
      useWorkerFetch: true,
      verbosity: 0,
    }).promise;
    const send = async (url: string, body: BodyInit): Promise<void> => {
      const res = await fetch(url, { method: 'PUT', body });
      if (!res.ok) throw new Error(`the harness server refused the upload (HTTP ${res.status})`);
    };
    const canvas = document.createElement('canvas');
    canvas.width = r.width;
    canvas.height = r.height;
    const context = canvas.getContext('2d', { willReadFrequently: true })!;
    try {
      for (const [i, target] of r.pages.entries()) {
        const page = await doc.getPage(i + 1);
        context.setTransform(1, 0, 0, 1, 0, 0);
        context.fillStyle = '#ffffff';
        context.fillRect(0, 0, r.width, r.height);
        await page.render({
          canvas,
          canvasContext: context,
          viewport: page.getViewport({ scale: r.scale }),
          background: '#ffffff',
        }).promise;
        const png = await new Promise<Blob>((done, fail) =>
          canvas.toBlob((b) => (b ? done(b) : fail(new Error('no PNG'))), 'image/png')
        );
        // The pixels gzip to a small share of their 8 MB, and Node unpacks them faster than it
        // decodes the PNG.
        const pixels = new Blob([context.getImageData(0, 0, r.width, r.height).data])
          .stream()
          .pipeThrough(new CompressionStream('gzip'));
        const rgba = await new Response(pixels).blob();
        await Promise.all([send(target.png, png), send(target.rgba, rgba)]);
        page.cleanup();
      }
    } finally {
      await doc.loadingTask.destroy();
    }
  },
  async reopenSaved() {
    if (!session) throw new Error('no document');
    const saved = writePptx(session.store.package);
    const r = await openPresentation(saved, { fonts });
    if (!r.ok) return { slides: 0, error: r.error.reason };
    await adopt(r.session);
    return { slides: r.session.info.slides.length };
  },
};

(window as unknown as { __pptxHarness: HarnessApi }).__pptxHarness = api;
document.body.dataset.ready = '1';
