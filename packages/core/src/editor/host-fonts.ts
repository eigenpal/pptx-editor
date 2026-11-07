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
import { within } from './within.ts';

/** The four faces of each packaged substitute family. */
const SUBSTITUTE_STYLES = [
  { weight: 400, style: 'normal' },
  { weight: 700, style: 'normal' },
  { weight: 400, style: 'italic' },
  { weight: 700, style: 'italic' },
] as const;

/**
 * The open families that each packaged substitute family covers: the face the engine loads,
 * and the metric-compatible face that `googleFonts()` serves for the same family (Arimo,
 * Tinos, Cousine), which draws the same metrics.
 */
const PACKAGED_SUBSTITUTE: Readonly<
  Record<NonNullable<DocumentSession['substituteFamilies']>[number], readonly string[]>
> = {
  Calibri: ['Carlito'],
  Cambria: ['Caladea'],
  Arial: ['Liberation Sans', 'Arimo'],
  'Times New Roman': ['Liberation Serif', 'Tinos'],
  'Courier New': ['Liberation Mono', 'Cousine'],
  'Century Gothic': ['TeX Gyre Adventor'],
  'Segoe UI': ['Selawik'],
};

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

/**
 * The longest a deck that opens waits for the host's fonts, as DOCX Editor's
 * `PREPARED_WORK_WAIT_MS`. After it, the deck paints with the embedded and built-in faces, and
 * the host's faces install when they arrive. A stalled download never holds the loading screen.
 */
export const FONT_WAIT_MS = 2000;

/** The answer `within` gives when the host's fonts did not arrive in time. */
const LATE = Symbol('late');

/**
 * How long late faces may hold `resolving` and hide the substitution notice, as DOCX Editor's
 * `LATE_FONT_WAIT_MS`. A stalled download must not hold them for the whole session. Faces
 * that arrive after this still install.
 */
export const LATE_FONT_WAIT_MS = 10_000;

/** The host's fonts for one editor. */
export interface HostFonts {
  /**
   * Resolve the origins for a deck that is opening, before its first paint, for at most
   * `waitMs` (default {@link FONT_WAIT_MS}); faces that arrive later install then.
   */
  open(session: DocumentSession, stale: () => boolean, waitMs?: number): Promise<void>;
  /**
   * The deck that `open` resolved for does not mount (a later load, a throw): the deck on
   * screen gets back the font state it had, and its late faces resolve again.
   */
  abandon(session: DocumentSession): void;
  /** Replace the `fonts` origin, as DOCX Editor's `setFonts`. */
  setFonts(fonts: EditorFontsOption | undefined): Promise<void>;
  /** Ask the origins for picked families that no face covers. */
  request(families: readonly string[]): Promise<void>;
  /** The `fonts` value in force: the option, or the value last passed to `setFonts`. */
  current(): EditorFontsOption | undefined;
  /** Families the origins can serve, for the font picker. */
  supportedFamilies(): readonly string[];
  /** Whether faces the deck asked for are still on the way, for at most {@link LATE_FONT_WAIT_MS}. */
  resolving(): boolean;
  /**
   * Shaped measurement that arrives after `session` opened: count it in `resolving`. Then,
   * while the editor shows `session`, register the new faces, or report why it did not load.
   */
  awaitShaping(session: DocumentSession): void;
  dispose(): void;
}

/** The host font state of one editor. */
export function createHostFonts(
  options: HostFontOptions,
  host: {
    /** Overrides {@link FONT_WAIT_MS}; for tests. */
    readonly waitMs?: number;
    /** Overrides {@link LATE_FONT_WAIT_MS}; for tests. */
    readonly lateWaitMs?: number;
    /** `resolving()` changed with nothing installed: publish. */
    readonly changed: () => void;
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
  // Late answers still on the way, each with the timer that stops waiting for it.
  const waiting = new Map<object, ReturnType<typeof setTimeout>>();
  // The open that has not mounted, and the state of the deck on screen before it.
  let opening: {
    readonly session: DocumentSession;
    readonly base: FontConfigurationFragment | undefined;
    readonly picked: readonly FontConfigurationFragment[];
    readonly requested: readonly string[];
    /** Faces of the deck on screen were still on the way. */
    readonly pending: boolean;
  } | null = null;
  const abandoned = new WeakSet<DocumentSession>();
  /** Stop waiting for `token`; `quiet` when an installation publishes the change. */
  const settle = (token: object, quiet = false): void => {
    const timer = waiting.get(token);
    if (timer === undefined) return;
    clearTimeout(timer);
    waiting.delete(token);
    if (!quiet && !destroyed) host.changed();
  };
  /** Waits for the session's own shaped measurement: new host fonts do not cancel them. */
  const shapingWaits = new WeakSet<object>();
  /** Stop waiting; `keepShaping` keeps the waits for the session's shaped measurement. */
  const stopWaiting = (keepShaping = false): void => {
    for (const [token, timer] of waiting) {
      if (keepShaping && shapingWaits.has(token)) continue;
      clearTimeout(timer);
      waiting.delete(token);
    }
  };
  /** Wait for `late`; `arrived` installs it and answers whether it installed anything. */
  const track = <T>(late: Promise<T>, arrived: (value: T, token: object) => boolean): object => {
    const token = {};
    waiting.set(
      token,
      setTimeout(() => settle(token), host.lateWaitMs ?? LATE_FONT_WAIT_MS)
    );
    late.then(
      (value) => {
        let installed = false;
        try {
          installed = arrived(value, token);
        } catch (cause) {
          // A late answer is untrusted shape: a failure to install is reported, never thrown.
          report(cause);
        }
        if (!installed) settle(token);
      },
      (cause: unknown) => {
        if (waiting.has(token)) report(cause);
        settle(token);
      }
    );
    return token;
  };

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
    track(late, (answer, token) => {
      if (!answer || !current() || host.session() !== session) return false;
      base = withoutLate(answer);
      apply(session);
      // Quietly: the installation publishes `resolving` with the new faces.
      settle(token, true);
      follow(session, answer.deferred, current);
      void host.installed(session);
      return true;
    });
  };

  /**
   * The faces the engine loads itself: the packaged substitutes of the session's families.
   * An origin that serves the same files (`packagedFonts()`) skips them, so no file loads twice.
   * The faces carry the substitute's own family (Carlito, not Calibri): an origin that serves
   * the requested family itself, such as `customFonts()` with the real Calibri, still loads it.
   */
  const engineFaces = (session: DocumentSession): { resolvedFaces?: FontFaceRequest[] } => {
    const families = session.substituteFamilies;
    if (!families || families.length === 0) return {};
    return {
      resolvedFaces: families.flatMap((family) =>
        PACKAGED_SUBSTITUTE[family].flatMap((open) =>
          SUBSTITUTE_STYLES.map((s) => ({ family: open, weight: s.weight, style: s.style }))
        )
      ),
    };
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
        { ...deckFontRequest(session.store.package, signal), ...engineFaces(session) },
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
    async open(session, stale, waitMs = host.waitMs ?? FONT_WAIT_MS) {
      // An earlier open that has not mounted keeps the state from before it.
      const earlier = opening && host.session() !== opening.session ? opening : null;
      opening = earlier
        ? { ...earlier, session }
        : { session, base, picked, requested: [...requested], pending: waiting.size > 0 };
      stopWaiting();
      base = undefined;
      picked = [];
      requested.clear();
      // A `setFonts` while the deck opens cancels this resolution; resolve again for it, in
      // the same wait: the retries do not extend it.
      const deadline = performance.now() + waitMs;
      for (let attempt = 0; attempt < 4; attempt += 1) {
        const lease = lifetime.advance();
        lease.cancelPrevious();
        let result: Awaited<ReturnType<typeof compose>> | typeof LATE;
        const composed = compose(session, lease.signal);
        try {
          result = await within(composed, Math.max(0, deadline - performance.now()), LATE);
        } catch (cause) {
          report(cause);
          return;
        }
        if (stale() || destroyed) return;
        if (result === LATE) {
          // The deck opens with the faces it has; the host's faces replace them on arrival.
          track(composed, (late, token) => {
            if (!late || !lease.current() || destroyed || abandoned.has(session)) return false;
            base = late.answer ? withoutLate(late.answer) : undefined;
            apply(session);
            follow(session, late.answer?.deferred, lease.current);
            // Before the editor adopts the session, it registers these faces itself.
            if (host.session() !== session) return false;
            settle(token, true);
            void host.installed(session);
            return true;
          });
          return;
        }
        if (!result) continue;
        base = result.answer ? withoutLate(result.answer) : undefined;
        apply(session);
        follow(session, result.answer?.deferred, lease.current);
        return;
      }
    },

    abandon(session) {
      if (destroyed || opening?.session !== session) return;
      const kept = opening;
      opening = null;
      abandoned.add(session);
      base = kept.base;
      picked = [...kept.picked];
      requested.clear();
      for (const family of kept.requested) requested.add(family);
      // The open stopped the shown deck's waits: wait for its shaping again (a settled one
      // ends at once), and resolve its late faces again.
      const shown = host.session();
      if (shown?.shaping) shapingWaits.add(track(shown.shaping, () => false));
      if (!shown || !kept.pending) return;
      const lease = lifetime.advance();
      lease.cancelPrevious();
      void compose(shown, lease.signal).then(
        (result) => {
          if (!result || !lease.current() || host.session() !== shown) return;
          base = result.answer ? withoutLate(result.answer) : undefined;
          apply(shown);
          void host.installed(shown);
          follow(shown, result.answer?.deferred, lease.current);
        },
        (cause: unknown) => report(cause)
      );
    },

    async setFonts(next) {
      if (destroyed) throw new Error('A destroyed editor cannot replace fonts');
      const lease = lifetime.advance();
      try {
        stopWaiting(true);
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
            ...engineFaces(session),
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

    resolving: () => waiting.size > 0,

    awaitShaping(session) {
      const work = session.shaping;
      if (!work) return;
      shapingWaits.add(track(work, () => false));
      void work.then((ok) => {
        if (destroyed || host.session() !== session) return;
        if (ok) void host.installed(session);
        else report(session.shapingError);
      });
    },

    supportedFamilies() {
      const out: string[] = [];
      for (const layer of [base, ...picked])
        for (const family of layer?.supportedFamilies ?? []) out.push(family);
      return out;
    },

    dispose() {
      destroyed = true;
      stopWaiting();
      lifetime.dispose();
    },
  };
}
