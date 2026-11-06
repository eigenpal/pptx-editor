/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pptx-to-pdf/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// The evaluation protocol's fixed settings and its font plan. `export.ts` and
// `layout-summary.ts` share them, as DOCX Editor's `export.ts`, `trace.ts`, and
// `layout-summary.ts` share one font profile (`packages/docx-to-pdf/scripts/evaluation/`,
// commit 398963c7f67207cebf3c38dcc41d5f0bd1339b5f).

import { createHash } from 'node:crypto';
import { readFile, stat } from 'node:fs/promises';
import { dirname, isAbsolute, resolve } from 'node:path';
import type { FontSource } from '@pptx-editor.dev/core/layout';
import {
  PdfDocumentOpenError,
  PdfWorkLimitError,
  type OpenPdfDocumentForExportResult,
} from '../../src/index.ts';

/** Protocol version. `contract.json` names the same number. */
export const PROTOCOL = 1;

/**
 * Settings of every evaluation export. The reference PDF leaves hidden slides out, so the
 * candidate does too. Approximate output is reported in `diagnostics`, never as strict success.
 */
export const EXPORT_SETTINGS = {
  fidelityPolicy: 'best-effort',
  includeHiddenSlides: false,
  timeoutMs: 60_000,
} as const;

/**
 * Session settings. The packaged metric-compatible substitutes from `@docx-editor.dev/fonts`
 * are always on, and no installed font is read unless a font plan names it, as DOCX Editor's
 * evaluation export sets `useSystemFonts: false`.
 */
export const SESSION_SETTINGS = { substitutes: true } as const;

/** Font plan limits. One plan holds at most the reference font set of one machine. */
const MAX_PLAN_FACES = 2_000;
const MAX_FACE_BYTES = 64 * 1024 * 1024;
const MAX_PLAN_BYTES = 1024 * 1024 * 1024;
const MAX_PLAN_FILE_BYTES = 1024 * 1024;

/** The identity of a loaded font plan: what a cache key must include. */
export interface FontPlanIdentity {
  readonly faces: number;
  /** SHA-256 over each face's family, style, and byte digest, in plan order. */
  readonly sha256: string;
}

const SFNT = new Set([0x00010000, 0x74727565 /* true */, 0x4f54544f /* OTTO */]);

/**
 * Read a font plan: `{ "protocol": 1, "faces": [{ "path", "family", "bold", "italic" }] }`.
 * A relative path resolves against the plan's folder. Each file is one TrueType or OpenType
 * face; collections are refused. The faces are read at run time and never copied.
 */
export async function loadFontPlan(
  path: string
): Promise<{ fonts: FontSource[]; identity: FontPlanIdentity }> {
  try {
    return await readFontPlan(path);
  } catch (error) {
    // A missing or unreadable plan or face is the plan's error, never the deck's (`ENOENT`).
    if (error instanceof ProtocolError) throw error;
    const message = error instanceof Error ? error.message : String(error);
    throw new ProtocolError('fontPlan', `Font plan: ${message}`, { cause: error });
  }
}

async function readFontPlan(
  path: string
): Promise<{ fonts: FontSource[]; identity: FontPlanIdentity }> {
  const planFile = await stat(path);
  if (!planFile.isFile()) throw new ProtocolError('fontPlan', 'Font plan must be a regular file');
  if (planFile.size > MAX_PLAN_FILE_BYTES)
    throw new ProtocolError('fontPlan', 'Font plan is too large');
  const plan = JSON.parse(await readFile(path, 'utf8')) as unknown;
  if (!plan || typeof plan !== 'object' || (plan as { protocol?: unknown }).protocol !== PROTOCOL)
    throw new ProtocolError('fontPlan', `Font plan must name protocol ${PROTOCOL}`);
  const faces = (plan as { faces?: unknown }).faces;
  if (!Array.isArray(faces) || faces.length === 0 || faces.length > MAX_PLAN_FACES)
    throw new ProtocolError('fontPlan', `Font plan must list from 1 to ${MAX_PLAN_FACES} faces`);
  const fonts: FontSource[] = [];
  const digest = createHash('sha256');
  let total = 0;
  for (const [index, face] of faces.entries()) {
    const { path: file, family, bold, italic } = (face ?? {}) as Record<string, unknown>;
    if (typeof file !== 'string' || typeof family !== 'string' || !family || family.length > 64)
      throw new ProtocolError('fontPlan', `Font plan face ${index}: path and family are required`);
    if (typeof bold !== 'boolean' || typeof italic !== 'boolean')
      throw new ProtocolError(
        'fontPlan',
        `Font plan face ${index}: bold and italic must be boolean`
      );
    const full = isAbsolute(file) ? file : resolve(dirname(path), file);
    const faceFile = await stat(full);
    if (!faceFile.isFile())
      throw new ProtocolError('fontPlan', `Font plan face ${index}: not a regular file`);
    if (faceFile.size > MAX_FACE_BYTES)
      throw new ProtocolError(
        'fontPlan',
        `Font plan face ${index}: the file is larger than ${MAX_FACE_BYTES} bytes`
      );
    const bytes = new Uint8Array(await readFile(full));
    total += bytes.byteLength;
    if (total > MAX_PLAN_BYTES) throw new ProtocolError('fontPlan', 'Font plan faces exceed 1 GiB');
    const tag = bytes.byteLength >= 4 ? new DataView(bytes.buffer).getUint32(0) : 0;
    if (!SFNT.has(tag))
      throw new ProtocolError(
        'fontPlan',
        `Font plan face ${index}: not a single TrueType or OpenType face`
      );
    fonts.push({ family, bold, italic, bytes, origin: 'plan' });
    const hash = createHash('sha256').update(bytes).digest('hex');
    digest.update(`${family}\u0000${bold ? 1 : 0}${italic ? 1 : 0}\u0000${hash}\n`);
  }
  return { fonts, identity: { faces: fonts.length, sha256: digest.digest('hex') } };
}

/** Command-line arguments: two paths, then `--fonts plan.json` at most once. */
export function parseArguments(
  argv: readonly string[],
  usage: string
): { input: string; output: string; fonts?: string } {
  const positional: string[] = [];
  let fonts: string | undefined;
  for (let i = 0; i < argv.length; i += 1) {
    const arg = argv[i]!;
    if (arg === '--fonts') {
      if (fonts !== undefined)
        throw new ProtocolError('usage', `--fonts is given twice. Expected ${usage}`);
      const value = argv[i + 1];
      if (!value || value.startsWith('--'))
        throw new ProtocolError('usage', `--fonts needs a path. Expected ${usage}`);
      fonts = value;
      i += 1;
    } else if (arg.startsWith('--'))
      throw new ProtocolError('usage', `Unknown option ${arg}. Expected ${usage}`);
    else positional.push(arg);
  }
  const [input, output] = positional;
  if (!input || !output || positional.length !== 2)
    throw new ProtocolError('usage', `Expected ${usage}`);
  return { input, output, ...(fonts ? { fonts } : {}) };
}

/**
 * One deadline for a whole protocol command, open and export together. Past it, `check` throws
 * `PdfWorkLimitError('timeout')`, so a slow deck reports `timeout` in every phase.
 */
export function protocolDeadline(timeoutMs: number = EXPORT_SETTINGS.timeoutMs) {
  const end = Date.now() + timeoutMs;
  return {
    /** The time left, at least 1 ms, for the export's own `timeoutMs`. */
    remaining: (): number => Math.max(1, end - Date.now()),
    check(): void {
      if (Date.now() >= end) throw new PdfWorkLimitError('timeout');
    },
  };
}

/**
 * The typed error for a deck that the session did not open. The protocol's only signal is its
 * deadline, so an aborted open is a `timeout`.
 */
export function openFailure(
  refusal: Extract<OpenPdfDocumentForExportResult, { ok: false }>
): PdfDocumentOpenError | PdfWorkLimitError {
  return refusal.reason === 'aborted'
    ? new PdfWorkLimitError('timeout')
    : new PdfDocumentOpenError(refusal.reason, refusal.detail);
}

/**
 * An error of the protocol itself, with a stable code: `usage` (arguments), `fontPlan` (the
 * plan or a face file), or `evidenceLimit` (the layout text is too large to record).
 */
export class ProtocolError extends Error {
  constructor(
    readonly code: 'usage' | 'fontPlan' | 'evidenceLimit',
    message: string,
    options?: ErrorOptions
  ) {
    super(message, options);
    this.name = 'ProtocolError';
  }
}

/**
 * The `code` of an error, for the evaluator's failure classes: the typed code, a Node.js file
 * code such as `ENOENT` for the deck, or `internal` for an engine failure with no code.
 */
export function errorCode(error: unknown): string {
  const code = (error as { code?: unknown } | null)?.code;
  return typeof code === 'string' ? code : 'internal';
}
