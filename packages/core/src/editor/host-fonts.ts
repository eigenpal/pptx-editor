// The host's fonts for one editor, DOCX Editor's way (PPTX only; the composition, the resolver
// mark, and the resolver lifetime are copies of DOCX Editor's modules). It composes the `fonts`
// origins for each deck, admits the faces (byte limits, content hash, sfnt check), and installs
// them in the session, after the embedded faces. `setFonts` replaces them in place; each `load`
// and each `setFonts` aborts the resolution before it.

import type {
  EditorFontErrorCode,
  FontConfiguration,
  FontFaceRequest,
} from '../contracts/index.ts';
import {
  boundedStructuralFontValidator,
  createFontResourceSnapshot,
  FontResolutionError,
  HARD_MAX_FONT_BYTES,
  trustedFontBytes,
} from '../layout/font-resource.ts';
import {
  composeFontConfiguration,
  type FontConfigurationFragment,
  type FontResolver,
} from '../layout/font-composition.ts';
import {
  composeFontOrigins,
  defineFontResolver,
  type FontOrigin,
} from '../layout/font-resolver.ts';
import type { FontSource as MeasuredFace } from '../layout/index.ts';
import { deckDefaultFamily, deckFontRequest } from './deck-font-request.ts';
import {
  HOST_SUBSTITUTE_ORIGIN,
  type DocumentSession,
  type HostFaces,
  type OpenOptions,
} from './document-session.ts';
import { EditorFontError } from './editor-font-error.ts';
import { FONT_NAME } from '../automation/commands/font-format.ts';
import { createFontResolutionLifetime } from './font-resolution-lifetime.ts';

/** What the `fonts` option and prop take, as DOCX Editor's `DocxEditorConfig.fonts`. @public */
export type EditorFontsOption = FontConfiguration | FontConfigurationFragment | FontResolver;

/** The font options `createPptxEditor` reads. */
export interface HostFontOptions {
  readonly fonts?: EditorFontsOption;
  readonly onFontError?: (error: EditorFontError) => void;
}

/** The options `openPresentation` takes, without the editor's font options. */
export function openOptionsOf(options: HostFontOptions & Omit<OpenOptions, 'fonts'>): OpenOptions {
  const { fonts, onFontError: _onError, ...rest } = options;
  // An array is the old form (`OpenOptions.fonts`, faces tried first), kept for callers that
  // passed the harness faces through the editor.
  const legacy: unknown = fonts;
  return Array.isArray(legacy) ? { ...rest, fonts: legacy as OpenOptions['fonts'] } : rest;
}

const LOAD_REASONS: Readonly<Record<string, EditorFontErrorCode>> = {
  hashMismatch: 'hashMismatch',
  overLimit: 'overLimit',
  malformed: 'malformed',
  forbidden: 'forbidden',
  missing: 'missing',
};

/** A typed error for anything a font origin reported. */
export function toEditorFontError(cause: unknown): EditorFontError {
  if (cause instanceof EditorFontError) return cause;
  if (cause instanceof FontResolutionError)
    return new EditorFontError(cause.code, cause.message, {
      request: cause.request,
      diagnostic: cause.diagnostic,
      cause,
    });
  if (cause && typeof cause === 'object' && 'reason' in cause && 'request' in cause) {
    // A `FontLoadFailure` from `loadFonts` or `customFonts`.
    const failure = cause as { reason: string; request: FontFaceRequest; diagnostic?: string };
    return new EditorFontError(
      LOAD_REASONS[failure.reason] ?? 'loadFailed',
      `Font load failed (${failure.reason})`,
      { request: failure.request, diagnostic: failure.diagnostic ?? failure.reason, cause }
    );
  }
  if (cause && typeof cause === 'object' && 'file' in cause && 'diagnostic' in cause)
    // A packaged-font failure from `@docx-editor.dev/fonts`.
    return new EditorFontError('loadFailed', 'Font load failed', {
      diagnostic: String((cause as { diagnostic: unknown }).diagnostic),
      cause,
    });
  const message = cause instanceof Error ? cause.message : String(cause);
  return new EditorFontError('loadFailed', message || 'Font load failed', {
    diagnostic: message,
    cause,
  });
}

/**
 * Hand a font error to the host's `onFontError`, or print one warning when the host has none,
 * so a failure is never silent. Shared by the React and Vue roots. @internal
 */
export function reportFontError(
  error: EditorFontError,
  handler: ((error: EditorFontError) => void) | undefined
): void {
  if (handler) handler(error);
  else
    console.warn(
      `[pptx-editor] ${error.message}${error.diagnostic ? `: ${error.diagnostic}` : ''}`
    );
}

/** The faces a set of fragments brings, admitted, and every face refused. */
function admit(layers: readonly FontConfigurationFragment[]): {
  faces: HostFaces;
  refused: EditorFontError[];
} {
  const refused: EditorFontError[] = [];
  const [base, ...rest] = layers;
  const config = composeFontConfiguration(base ?? {}, ...rest);
  let snapshot: ReturnType<typeof createFontResourceSnapshot>;
  try {
    snapshot = createFontResourceSnapshot({
      epoch: 0,
      maxFontBytes: config.maxFontBytes ?? HARD_MAX_FONT_BYTES,
      resources: config.sources,
      substitutions: config.substitutions ?? [],
      validateFont: boundedStructuralFontValidator,
    });
  } catch (cause) {
    return { faces: { faces: [], substitutions: [] }, refused: [toEditorFontError(cause)] };
  }
  const measured = (
    request: FontFaceRequest,
    origin: string
  ): MeasuredFace | EditorFontError | null => {
    const font = snapshot.resolve(request);
    if (font instanceof FontResolutionError)
      return font.code === 'missing' ? null : toEditorFontError(font);
    // The measurer and paint open face 0 of a file only.
    if (font.faceIndex !== 0)
      return new EditorFontError('unsupportedFaceIndex', 'Only face 0 of a font file is used', {
        request,
      });
    return {
      family: request.family,
      bold: request.weight >= 600,
      italic: request.style === 'italic',
      bytes: trustedFontBytes(font),
      origin,
    };
  };
  const faces: MeasuredFace[] = [];
  for (const source of config.sources) {
    if (source.availability === 'forbidden') continue;
    const face = measured(source.request, 'host');
    if (face instanceof EditorFontError) refused.push(face);
    else if (face) faces.push(face);
  }
  const substitutions: { face: MeasuredFace; target: string }[] = [];
  for (const substitution of config.substitutions ?? []) {
    const face = measured(substitution.from, HOST_SUBSTITUTE_ORIGIN);
    if (face instanceof EditorFontError) refused.push(face);
    else if (face) substitutions.push({ face, target: substitution.to.family });
  }
  return { faces: { faces, substitutions }, refused };
}

/**
 * Add a font origin after the host's current fonts, as the next origin in first-wins order.
 * The host's faces keep precedence; the added origin answers only for what they leave open.
 * Pass the result to `setFonts`, for example `editor.setFonts(withFontOrigin(editor.fonts,
 * { sources }))`. @public
 */
export function withFontOrigin(
  current: EditorFontsOption | undefined,
  added: FontOrigin
): EditorFontsOption {
  if (current === undefined) return added as EditorFontsOption;
  const origins: readonly FontOrigin[] = [current as FontOrigin, added];
  return defineFontResolver((request) => composeFontOrigins(origins, request));
}

const isAbort = (signal: AbortSignal | undefined): boolean => signal?.aborted === true;

/** The host's fonts for one editor. */
export interface HostFonts {
  /** Resolve the origins for a deck that is opening, before its first paint. */
  open(session: DocumentSession, stale: () => boolean): Promise<void>;
  /** Replace the `fonts` origin, as DOCX Editor's `setFonts`. */
  setFonts(fonts: EditorFontsOption | undefined): Promise<void>;
  /** Ask the origins for picked families that no face covers. */
  request(families: readonly string[]): Promise<void>;
  /** The `fonts` value in force: the option, or the value last passed to `setFonts`. */
  current(): EditorFontsOption | undefined;
  /** Families the origins can serve, for the font picker. */
  supportedFamilies(): readonly string[];
  dispose(): void;
}

/** The host font state of one editor. */
export function createHostFonts(
  options: HostFontOptions,
  host: {
    readonly session: () => DocumentSession | null;
    /** The session's faces changed: register them for paint and publish. */
    readonly installed: (session: DocumentSession) => Promise<void>;
  }
): HostFonts {
  const lifetime = createFontResolutionLifetime();
  let fonts = Array.isArray(options.fonts) ? undefined : options.fonts;
  let destroyed = false;
  // The layers of the current session, first-wins: the origins' answer, then picks.
  let base: FontConfigurationFragment | undefined;
  let picked: FontConfigurationFragment[] = [];
  const requested = new Set<string>();

  const report = (cause: unknown): void =>
    reportFontError(toEditorFontError(cause), options.onFontError);
  const origins = (): FontOrigin[] => (fonts === undefined ? [] : [fonts as FontOrigin]);
  const apply = (session: DocumentSession): void => {
    const layers = [base, ...picked].filter((l): l is FontConfigurationFragment => l !== undefined);
    if (layers.length === 0) return session.clearHostFaces();
    const { faces, refused } = admit(layers);
    for (const error of refused) report(error);
    session.setHostFaces(faces);
  };
  const withoutLate = (fragment: FontConfigurationFragment): FontConfigurationFragment => {
    const { deferred: _late, ...now } = fragment;
    return now;
  };

  /** Late faces of an answer: each arrival replaces the base while its lifetime is current. */
  const follow = (
    session: DocumentSession,
    late: Promise<FontConfigurationFragment | undefined> | undefined,
    current: () => boolean
  ): void => {
    if (!late) return;
    void late.then(async (answer) => {
      if (!answer || !current() || host.session() !== session) return;
      base = withoutLate(answer);
      apply(session);
      await host.installed(session);
      follow(session, answer.deferred, current);
    });
  };

  /** Compose the origins for the session; undefined when the resolution was cancelled. */
  const compose = async (
    session: DocumentSession,
    signal: AbortSignal
  ): Promise<{ answer: FontConfigurationFragment | undefined } | null> => {
    const list = origins();
    if (list.length === 0) return { answer: undefined };
    try {
      const answer = await composeFontOrigins(
        list,
        deckFontRequest(session.store.package, signal),
        {
          onOriginFailure: (failure) => {
            if (!isAbort(signal)) report(failure.cause);
          },
        }
      );
      return isAbort(signal) ? null : { answer };
    } catch (cause) {
      if (isAbort(signal)) return null;
      throw cause;
    }
  };

  return {
    async open(session, stale) {
      base = undefined;
      picked = [];
      requested.clear();
      // A `setFonts` while the deck opens cancels this resolution; resolve again for it.
      for (let attempt = 0; attempt < 4; attempt += 1) {
        const lease = lifetime.advance();
        lease.cancelPrevious();
        let result: Awaited<ReturnType<typeof compose>>;
        try {
          result = await compose(session, lease.signal);
        } catch (cause) {
          report(cause);
          return;
        }
        if (stale() || destroyed) return;
        if (!result) continue;
        base = result.answer ? withoutLate(result.answer) : undefined;
        apply(session);
        follow(session, result.answer?.deferred, lease.current);
        return;
      }
    },

    async setFonts(next) {
      if (destroyed) throw new Error('A destroyed editor cannot replace fonts');
      const lease = lifetime.advance();
      try {
        fonts = next;
        picked = [];
        requested.clear();
      } finally {
        lease.cancelPrevious();
      }
      const session = host.session();
      if (!session) return;
      // Removal drops the host faces now. A replacement keeps the current faces painted
      // until the new ones install, so the slides lay out once, not twice.
      if (next === undefined) {
        base = undefined;
        apply(session);
        await host.installed(session);
      }
      let result: Awaited<ReturnType<typeof compose>>;
      try {
        result = await compose(session, lease.signal);
      } catch (cause) {
        result = { answer: undefined };
        report(cause);
      }
      if (!result || !lease.current() || host.session() !== session) return;
      if (next === undefined && result.answer === undefined) return;
      // Nothing usable, or a failure: the replaced faces go anyway.
      base = result.answer ? withoutLate(result.answer) : undefined;
      apply(session);
      await host.installed(session);
      follow(session, result.answer?.deferred, lease.current);
    },

    async request(families) {
      const session = host.session();
      const list = origins();
      if (!session || list.length === 0) return;
      const have = new Set(session.faces.map((f) => f.family.toLowerCase()));
      const wanted = families.filter(
        (f) => FONT_NAME.test(f) && !have.has(f.toLowerCase()) && !requested.has(f.toLowerCase())
      );
      if (wanted.length === 0) return;
      for (const f of wanted) requested.add(f.toLowerCase());
      const lease = lifetime.capture();
      let answer: FontConfigurationFragment | undefined;
      try {
        answer = await composeFontOrigins(
          list,
          {
            families: wanted,
            defaultFamily: deckDefaultFamily(session.store.package),
            signal: lease.signal,
          },
          { onOriginFailure: (failure) => !isAbort(lease.signal) && report(failure.cause) }
        );
      } catch (cause) {
        if (!isAbort(lease.signal)) report(cause);
        return;
      }
      if (!answer || !lease.current() || host.session() !== session) return;
      // A pick adds faces; it never replaces the deck's.
      picked = [...picked, withoutLate(answer)];
      apply(session);
      await host.installed(session);
    },

    current() {
      return fonts;
    },

    supportedFamilies() {
      const out: string[] = [];
      for (const layer of [base, ...picked])
        for (const family of layer?.supportedFamilies ?? []) out.push(family);
      return out;
    },

    dispose() {
      destroyed = true;
      lifetime.dispose();
    },
  };
}
