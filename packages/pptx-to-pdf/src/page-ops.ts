/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pptx-to-pdf/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// The operator buffer of one PDF page, its resources, and the slide's diagnostics. The paint
// modules (`paint.ts`, `paint-fill.ts`, `paint-table.ts`, `paint-effects.ts`) share it.

import {
  appendBezierCurve,
  closePath,
  concatTransformationMatrix,
  lineTo,
  moveTo,
  setGraphicsState,
  PDFContentStream,
  PDFDict,
  PDFName,
  PDFOperator,
  PDFOperatorNames,
  type PDFDocument,
  type PDFImage,
  type PDFPage,
  type PDFRef,
} from 'pdf-lib';
import type { Rgba } from '@pptx-editor.dev/core/cascade';
import type { Matrix, PathCommand, RegisteredFace } from '@pptx-editor.dev/core/layout';
import type { PptxPackage } from '@pptx-editor.dev/core/store';
import type { EmbeddedFace } from './fonts.ts';
import type { Work } from './options.ts';
import type { PdfDiagnostic } from './types.ts';

export const EMU_PER_PT = 12_700;

/** A file-derived string, bounded and without control characters, for a diagnostic message. */
export function clip(text: string, max = 64): string {
  let clean = '';
  for (const ch of text) {
    const code = ch.codePointAt(0) ?? 0;
    clean += code < 0x20 || code === 0x7f ? ' ' : ch;
  }
  return clean.length > max ? `${clean.slice(0, max - 1)}…` : clean;
}

/** Collects diagnostics for one slide. Codes are fixed strings; file strings go in messages. */
export class SlideDiagnostics {
  readonly #seen = new Set<string>();
  readonly list: PdfDiagnostic[] = [];
  /** `pageIndex` is the output page, undefined when the slide produced no page. */
  constructor(
    readonly slideIndex: number,
    readonly pageIndex: number | undefined
  ) {}

  add(code: string, severity: PdfDiagnostic['severity'], message: string): void {
    const key = `${code}\u0000${message}`;
    if (this.#seen.has(key)) return;
    this.#seen.add(key);
    this.list.push({
      code,
      message,
      severity,
      slideIndex: this.slideIndex,
      slideNumber: this.slideIndex + 1,
      ...(this.pageIndex !== undefined
        ? { pageIndex: this.pageIndex, pageNumber: this.pageIndex + 1 }
        : {}),
    });
  }
}

/** The stop functions of one gradient, written once per document and shared by its shadings. */
export interface StopFunctions {
  readonly color: PDFRef;
  /** The alpha function for a soft mask, or `null` when every stop has the same alpha. */
  readonly alpha: PDFRef | null;
  /** The alpha of every stop, when `alpha` is `null`. */
  readonly flatAlpha: number;
}

/** Fonts, images, and gradient functions shared by every page of one document. */
export interface PaintResources {
  readonly doc: PDFDocument;
  readonly pkg: PptxPackage;
  readonly work: Work;
  readonly faces: ReadonlyMap<string, RegisteredFace>;
  /** Embedded faces by face id; `null` when the face may not be embedded. */
  readonly fonts: Map<string, EmbeddedFace | null>;
  /** The font files that EMF and WMF text draws with, read once per document. */
  metafileFonts?: readonly Uint8Array[];
  /** Embedded pictures by part name; `null` when the part cannot be drawn. */
  readonly images: Map<string, PDFImage | null>;
  /**
   * Stop functions by their drawn stops. A gradient of a group, layout, or theme is written
   * once, however many shapes use it.
   */
  readonly gradients: Map<string, StopFunctions>;
  /** Tiling patterns by preset, colors, and page height, written once per document. */
  readonly patterns: Map<string, PDFRef>;
}

/** Fill, clip, and path operators that `pdf-lib` has no helper for. */
export const fillEvenOdd = (): PDFOperator => PDFOperator.of(PDFOperatorNames.FillEvenOdd);
export const fillEvenOddAndStroke = (): PDFOperator =>
  PDFOperator.of(PDFOperatorNames.FillEvenOddAndStroke);
export const clipEvenOdd = (): PDFOperator => PDFOperator.of(PDFOperatorNames.ClipEvenOdd);
export const clipNonZero = (): PDFOperator => PDFOperator.of(PDFOperatorNames.ClipNonZero);
export const endPath = (): PDFOperator => PDFOperator.of(PDFOperatorNames.EndPath);
export const shadingFill = (name: PDFName): PDFOperator =>
  PDFOperator.of(PDFOperatorNames.ShadingFill, [name]);

/**
 * Form content up to this size is written plain. Deflate saves little on it, and each deflate
 * call costs a fixed setup: a shape with effects writes about 40 small forms.
 */
const PLAIN_FORM_BYTES = 1024;

/** The resource kinds a page names, and the prefix of each name. */
const PREFIX = { Font: 'F', XObject: 'X', ExtGState: 'GS', Shading: 'Sh', Pattern: 'P' } as const;

export class PageOps {
  #ops: PDFOperator[] = [];
  readonly #states = new Map<string, PDFName>();
  readonly #fontKeys = new Map<string, PDFName>();
  /** This page's names of the document's tiling patterns (`PaintResources.patterns`). */
  readonly patterns = new Map<string, PDFName>();
  #resourcesRef: PDFRef | undefined;
  #names = 0;
  #items = 0;
  constructor(
    readonly page: PDFPage,
    readonly res: PaintResources,
    readonly diag: SlideDiagnostics
  ) {}

  /**
   * Count one drawn item, and every 64 items let the timer and the caller's signal run, as DOCX
   * Editor's paint yields. Without it an abort is seen only after the paint and the save.
   */
  async step(): Promise<void> {
    this.#items += 1;
    if (this.#items % 64 === 0) await this.res.work.yield();
  }

  push(...more: PDFOperator[]): void {
    this.res.work.count(more.length);
    let bytes = 0;
    for (const op of more) {
      bytes += op.sizeInBytes() + 1;
      this.#ops.push(op);
    }
    this.res.work.reserveContent(bytes);
  }

  /** Run `draw` with its operators collected apart from the page, for a form XObject. */
  async capture(draw: () => Promise<void>): Promise<PDFOperator[]> {
    const outer = this.#ops;
    this.#ops = [];
    try {
      await draw();
      return this.#ops;
    } finally {
      this.#ops = outer;
    }
  }

  /** The page's resource name for a face, added once per page. */
  fontKey(faceId: string, font: EmbeddedFace): PDFName {
    let key = this.#fontKeys.get(faceId);
    if (!key) {
      key = this.name('Font', font.ref);
      this.#fontKeys.set(faceId, key);
    }
    return key;
  }

  /**
   * Add `ref` to the page resources under a new name from a counter of this page, as DOCX Editor's
   * `packages/docx-to-pdf/src/vector-gradient.ts` does. `pdf-lib`'s own helpers (`uniqueKey`) copy
   * every name on each call, so a page with many effects would cost the square of its resource
   * count. DOCX Editor commit adf0e178.
   */
  name(kind: keyof typeof PREFIX, ref: PDFRef): PDFName {
    const { Resources } = this.page.node.normalizedEntries();
    let dict = Resources.lookupMaybe(PDFName.of(kind), PDFDict);
    if (!dict) {
      dict = this.res.doc.context.obj({});
      Resources.set(PDFName.of(kind), dict);
    }
    this.#names += 1;
    const key = PDFName.of(`${PREFIX[kind]}${this.#names}`);
    dict.set(key, ref);
    return key;
  }

  /**
   * The page resources as one indirect object, shared by reference by every form this page
   * draws. A direct copy in each form would hold every other form too, and the PDF would grow
   * with the square of the number of forms.
   */
  resources(): PDFRef {
    if (this.#resourcesRef) return this.#resourcesRef;
    const dict = this.page.node.normalizedEntries().Resources;
    this.#resourcesRef = this.res.doc.context.register(dict);
    this.page.node.set(PDFName.of('Resources'), this.#resourcesRef);
    return this.#resourcesRef;
  }

  /**
   * Register a form XObject of ready operators that reads this page's resources. `group` makes it
   * a transparency group.
   */
  form(
    ops: PDFOperator[],
    bbox: number[],
    group?: { readonly S: 'Transparency'; readonly CS?: 'DeviceGray' }
  ): PDFRef {
    const ctx = this.res.doc.context;
    let bytes = 0;
    for (const op of ops) bytes += op.sizeInBytes() + 1;
    const dict = ctx.obj({
      Type: 'XObject',
      Subtype: 'Form',
      BBox: bbox,
      ...(group ? { Group: group } : {}),
      Resources: this.resources(),
    });
    // The operators are charged as they are pushed; the dictionary is charged here.
    this.res.work.reserveContent(dict.sizeInBytes());
    return ctx.register(PDFContentStream.of(dict, ops, bytes > PLAIN_FORM_BYTES));
  }

  /** Set an ExtGState; `key` caches it on this page. */
  state(key: string, make: () => PDFDict): void {
    let name = this.#states.get(key);
    if (!name) {
      const dict = make();
      this.res.work.reserveContent(dict.sizeInBytes());
      name = this.name('ExtGState', this.res.doc.context.register(dict));
      this.#states.set(key, name);
    }
    this.push(setGraphicsState(name));
  }

  /** Fill and stroke opacity through an ExtGState. */
  alpha(fillAlphaIn: number, strokeAlphaIn: number): void {
    // File values (for example `alphaModFix`) can fall outside 0–1; PDF allows only 0–1.
    const fillAlpha = Math.max(0, Math.min(1, fillAlphaIn));
    const strokeAlpha = Math.max(0, Math.min(1, strokeAlphaIn));
    if (fillAlpha >= 1 && strokeAlpha >= 1) return;
    this.state(`${fillAlpha}/${strokeAlpha}`, () =>
      this.res.doc.context.obj({ Type: 'ExtGState', ca: fillAlpha, CA: strokeAlpha })
    );
  }

  matrix(m: Matrix): void {
    this.push(concatTransformationMatrix(m[0], m[1], m[2], m[3], m[4], m[5]));
  }

  flush(): void {
    for (let i = 0; i < this.#ops.length; i += 1000)
      this.page.pushOperators(...this.#ops.slice(i, i + 1000));
    this.#ops.length = 0;
  }
}

export const rgb = (c: Rgba): [number, number, number] => [c.r / 255, c.g / 255, c.b / 255];

/** The uniform scale a matrix applies, as the SVG paint reads it for line widths in groups. */
export function matrixScale(m: Matrix): number {
  return Math.sqrt(Math.abs(m[0] * m[3] - m[1] * m[2])) || 1;
}

export function tracePath(commands: readonly PathCommand[], p: PageOps): void {
  let cx = 0;
  let cy = 0;
  for (const c of commands) {
    if (c.op === 'M') p.push(moveTo(c.x, c.y));
    else if (c.op === 'L') p.push(lineTo(c.x, c.y));
    else if (c.op === 'C') p.push(appendBezierCurve(c.x1, c.y1, c.x2, c.y2, c.x, c.y));
    else if (c.op === 'Q') {
      // A quadratic Bézier is the cubic with control points 2/3 of the way to its one control.
      p.push(
        appendBezierCurve(
          cx + ((c.x1 - cx) * 2) / 3,
          cy + ((c.y1 - cy) * 2) / 3,
          c.x + ((c.x1 - c.x) * 2) / 3,
          c.y + ((c.y1 - c.y) * 2) / 3,
          c.x,
          c.y
        )
      );
    } else p.push(closePath());
    if (c.op !== 'Z') {
      cx = c.x;
      cy = c.y;
    }
  }
}

/** A `w` × `h` rectangle at `x`, `y`, as a path. */
export function traceRect(x: number, y: number, w: number, h: number, p: PageOps): void {
  p.push(moveTo(x, y), lineTo(x + w, y), lineTo(x + w, y + h), lineTo(x, y + h), closePath());
}
