// The fidelity harness page. Playwright loads it, hands it deck bytes, and screenshots each
// slide. It renders through the SAME session, layout, and paint the editor uses: there is no
// second renderer for tests.

import { openPresentation, type DocumentSession } from '@pptx-editor.dev/core/editor';
import { paintSlide, registerFontFaces, type FontFaceHandle } from '@pptx-editor.dev/core/output';
import { writePptx } from '@pptx-editor.dev/core/store';
import type { FontSource, LayoutItem } from '@pptx-editor.dev/core/layout';
import { installDrawing } from '@pptx-editor.dev/pro/drawing';

// The harness renders as the demos do: with the Pro drawing renderers (glow, inner shadow,
// reflection, and soft edge), installed before any deck opens.
installDrawing();

/** A deck and its local fonts, as paths on the harness server (`openDeck`). */
interface DeckSource {
  deck: string;
  fonts: { family: string; bold: boolean; italic: boolean; url: string }[];
}

/** The size and SHA-256 of a file that the page got. */
interface ReceivedFile {
  bytes: number;
  sha256: string;
}

/** One laid-out text line, in points from the slide's top-left. */
interface HarnessLine {
  node: string;
  text: string;
  baselinePt: number;
  leftPt: number;
  /** The line is turned 180° (a net vertical flip), so its baseline is near the top of its box. */
  upsideDown: boolean;
  /** The line's bullet or number, which the PDF text holds before the line's text. */
  bullet?: string;
  /** The bullet is a number (`a:buAutoNum`), so its text must match the PDF exactly. */
  numbered?: boolean;
  /** The line starts outside the slide, where `pdftotext` keeps no text. */
  offSlide?: boolean;
}

interface HarnessApi {
  /** Fetch the deck and its fonts from the harness server, and open the deck. */
  open(source: DeckSource): Promise<{
    slides: number;
    error?: string;
    received: { deck: ReceivedFile; fonts: (ReceivedFile & { family: string })[] };
  }>;
  render(
    index: number,
    widthPx: number
  ): Promise<{ lines: HarnessLine[]; diagnostics: readonly string[]; fonts: string[] }>;
  reopenSaved(): Promise<{ slides: number; error?: string }>;
}

let session: DocumentSession | null = null;
let faces: FontFaceHandle | null = null;
let fonts: FontSource[] = [];

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

async function received(bytes: Uint8Array<ArrayBuffer>): Promise<ReceivedFile> {
  const digest = new Uint8Array(await crypto.subtle.digest('SHA-256', bytes));
  const sha256 = [...digest].map((b) => b.toString(16).padStart(2, '0')).join('');
  return { bytes: bytes.length, sha256 };
}
const EMU_PER_PT = 12_700;

/**
 * Collect line texts and absolute baselines (unrotated text only). A line that starts outside the
 * slide is marked: `pdftotext` keeps no text outside the page.
 */
function collectLines(
  items: readonly LayoutItem[],
  slide: { readonly width: number; readonly height: number },
  out: HarnessLine[]
): void {
  for (const it of items) {
    if (it.kind === 'group') collectLines(it.children, slide, out);
    if (it.kind !== 'shape' || !it.text || !it.textMatrix) continue;
    const m = it.textMatrix;
    if (Math.abs(m[1]) > 1e-6 || Math.abs(m[2]) > 1e-6 || it.text.rotation % 360 !== 0) continue;
    for (const line of it.text.lines) {
      const runs = line.runs.filter((r) => r.kind === 'text');
      const text = runs.map((r) => r.text).join('');
      if (!text.trim()) continue;
      const bullet = line.runs.find((r) => r.kind === 'bullet');
      const x = it.text.box.x + (bullet ?? runs[0])?.x[0];
      const y = it.text.box.y + line.baseline;
      const baseline = m[1] * x + m[3] * y + m[5];
      const left = m[0] * x + m[2] * y + m[4];
      const offSlide = baseline < 0 || baseline > slide.height || left < 0 || left > slide.width;
      const numbered = it.element.text?.paragraphs[line.paragraph]?.bullet?.type === 'autonum';
      out.push({
        node: it.element.nodeId,
        text,
        baselinePt: baseline / EMU_PER_PT,
        leftPt: left / EMU_PER_PT,
        upsideDown: m[3] < 0,
        ...(bullet?.text.trim() ? { bullet: bullet.text.trim(), numbered } : {}),
        ...(offSlide ? { offSlide } : {}),
      });
    }
  }
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
  faces?.dispose();
  session = next;
  faces = registerFontFaces(document, session.faces);
  await faces.ready;
}

const api: HarnessApi = {
  async open(source) {
    // No slide of the last deck may render, or reopen, with this deck's fonts, also when a
    // fetch below fails.
    session?.dispose();
    faces?.dispose();
    session = null;
    faces = null;
    fonts = [];
    const deck = await fetchBytes(source.deck, 'the deck');
    const got = { deck: await received(deck), fonts: [] as (ReceivedFile & { family: string })[] };
    for (const f of source.fonts) {
      const bytes = await fetchBytes(f.url, `the font ${f.family}`);
      got.fonts.push({ family: f.family, ...(await received(bytes)) });
      fonts.push({ family: f.family, bold: f.bold, italic: f.italic, bytes, origin: 'local' });
    }
    const r = await openPresentation(deck, { fonts });
    if (!r.ok) return { slides: 0, error: r.error.reason, received: got };
    await adopt(r.session);
    return { slides: r.session.info.slides.length, received: got };
  },
  async render(index, widthPx) {
    if (!session) throw new Error('no document');
    const layout = session.layout(index);
    if (!layout) throw new Error(`no slide ${index}`);
    const svg = paintSlide(document, layout, {
      images: session.imageUrl,
      imageSize: session.imageSize,
      idPrefix: `s${index}`,
    });
    svg.setAttribute('width', String(widthPx));
    // `renderSizePx` in scripts/fidelity/references.ts computes this size the same way, and
    // `rasterOptions` crops the PDF raster to it.
    svg.setAttribute(
      'height',
      String(Math.round((widthPx * layout.slide.height) / layout.slide.width))
    );
    document.getElementById('slide')!.replaceChildren(svg);
    await document.fonts.ready;
    await Promise.all(
      [...svg.querySelectorAll('image')].map(
        (img) =>
          new Promise<void>((done) => {
            const probe = new Image();
            probe.onload = () => done();
            probe.onerror = () => done();
            probe.src = img.getAttribute('href') ?? '';
          })
      )
    );
    const lines: HarnessLine[] = [];
    collectLines(layout.items, layout.slide, lines);
    const fonts = new Set<string>();
    collectFonts(layout.items, fonts);
    return { lines, diagnostics: layout.diagnostics, fonts: [...fonts].sort() };
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
