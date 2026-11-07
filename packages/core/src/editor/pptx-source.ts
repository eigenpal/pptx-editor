// Adapted from DOCX Editor `packages/react/src/editor/useDocxSource.ts` at commit 5993bd223.
//
// Open a presentation from a URL or bytes, with its fonts: the framework-free part of
// `usePptxSource`. The React and Vue adapters own only the lifetime (cancel on unmount or on a
// new source); the rules are here, so the two cannot drift.
//
// TWO FONT PATHS, and the path decides whether the deck waits:
//
//   EAGER: a value, a promise, or a zero-argument loader. The answer is complete before the
//   file is parsed, so the bytes are held back until it lands and the deck lays out once.
//
//   ON DEMAND: a resolver marked by `defineFontResolver` (`packagedFonts()`, `googleFonts()`).
//   Its answer depends on the families the file names, which nothing knows until the engine
//   has parsed it. There is nothing to wait for, so the bytes go straight through.
//
// The two are the same as types (`() => X` is assignable to `(request) => X`), so the mark,
// not the arity, picks the path.

import type { FontConfiguration } from '../contracts/index.ts';
import {
  composeFontConfiguration,
  type FontConfigurationBase,
  type FontConfigurationFragment,
  type FontResolutionRequest,
} from '../layout/font-composition.ts';
import {
  composeFontOrigins,
  defineFontResolver,
  isFontResolver,
  type FontOrigin,
  type MarkedFontResolver,
} from '../layout/font-resolver.ts';
import { FONT_WAIT_MS } from './host-fonts.ts';

/** A complete font configuration, or a fragment that composes with the defaults. @public */
export type PptxFontsInput = FontConfiguration | FontConfigurationFragment;

/**
 * One entry of a `fonts` list: any `FontOrigin`, or a zero-argument loader. A function that
 * takes a request must be marked by `defineFontResolver`. @public
 */
export type PptxFontOrigin = FontOrigin | (() => PptxFontsInput | Promise<PptxFontsInput>);

/** How a host supplies fonts: one origin, or a list of them in precedence order. @public */
export type PptxFontsSource = PptxFontOrigin | readonly PptxFontOrigin[];

/**
 * The presentation: a URL to fetch, or bytes already in hand. A string is always a URL. @public
 */
export type PptxSource = string | URL | Uint8Array | ArrayBuffer;

/** Options for `usePptxSource`. @public */
export interface UsePptxSourceOptions {
  /** The fonts. Eager origins hold the deck back until they settle, for at most 2 seconds. */
  readonly fonts?: PptxFontsSource;
  /** Passed to `fetch` for a URL source: credentials and headers. */
  readonly fetchOptions?: RequestInit;
}

/** The bytes of a source that needs no fetch, or null for a URL. @internal */
export function pptxSourceBytes(source: PptxSource): Uint8Array | null {
  if (source instanceof Uint8Array) return source;
  if (source instanceof ArrayBuffer) return new Uint8Array(source);
  return null;
}

/** One origin or several, as a list. */
function fontOrigins(source: PptxFontsSource): readonly PptxFontOrigin[] {
  return Array.isArray(source) ? source : [source as PptxFontOrigin];
}

/** A loader in an on-demand list is still a loader: it is called with no argument. */
function asFontOrigin(origin: PptxFontOrigin): FontOrigin {
  if (typeof origin !== 'function' || isFontResolver(origin)) return origin as FontOrigin;
  const loader = origin as () => PptxFontsInput | Promise<PptxFontsInput>;
  return defineFontResolver(async () => loader());
}

/** Which path a `fonts` option takes: absent, eager, or on demand. @internal */
export type PptxFontsShape = 'absent' | 'eager' | 'on-demand';

/**
 * The path of `source`. One on-demand origin makes the whole list on demand: the others may
 * be complete before the parse, but the composition is not. @internal
 */
export function pptxFontsShape(source: PptxFontsSource | undefined): PptxFontsShape {
  if (source === undefined) return 'absent';
  return fontOrigins(source).some((origin) => isFontResolver(origin)) ? 'on-demand' : 'eager';
}

/**
 * One resolver for the life of a hook, which reads the latest `fonts` at each request. Inline
 * `packagedFonts()` is a new function at each render; this keeps the `fonts` prop stable, so
 * the editor does not resolve its fonts again on each render. @internal
 */
export function pptxSourceFontsResolver(
  latest: () => PptxFontsSource | undefined
): MarkedFontResolver {
  return defineFontResolver((request: FontResolutionRequest) =>
    composeFontOrigins(fontOrigins(latest() ?? []).map(asFontOrigin), request)
  );
}

/** Eager path only: a zero-argument loader is called, everything else is taken as it is. */
async function resolveEagerOrigin(origin: PptxFontOrigin): Promise<PptxFontsInput | undefined> {
  if (typeof origin !== 'function') return (await origin) as PptxFontsInput | undefined;
  return await (origin as () => PptxFontsInput | Promise<PptxFontsInput>)();
}

/**
 * Resolve an eager `fonts` option to one configuration, or undefined when no origin answered.
 * One origin that fails does not take the others with it. A font failure never fails the deck:
 * it prints a warning, and the engine measures with its fallback. @internal
 */
export async function resolveEagerPptxFonts(
  source: PptxFontsSource
): Promise<FontConfiguration | undefined> {
  try {
    const settled = await Promise.allSettled(fontOrigins(source).map(resolveEagerOrigin));
    const resolved: PptxFontsInput[] = [];
    for (const outcome of settled) {
      if (outcome.status === 'rejected') {
        console.warn(
          '[fonts] font loading failed for one origin; the others still compose',
          outcome.reason
        );
        continue;
      }
      if (outcome.value !== undefined) resolved.push(outcome.value);
    }
    if (resolved.length === 0) return undefined;
    return composeFontConfiguration(
      resolved[0] as FontConfigurationBase,
      ...(resolved.slice(1) as readonly FontConfigurationFragment[])
    );
  } catch (cause) {
    console.warn('[fonts] font loading failed; the fixed measurer stays in effect', cause);
    return undefined;
  }
}

/**
 * The longest eager fonts hold the deck back, as the editor's own font wait. Fonts that arrive
 * later still apply: the editor reflows. @internal
 */
export const EAGER_FONTS_WAIT_MS = FONT_WAIT_MS;

/**
 * Resolve an eager `fonts` option for a hook. `release` runs once, when the fonts settle or
 * after `waitMs`, whichever is first, so an origin that never answers cannot hold the deck.
 * `resolved` runs when they settle, also after the release. Returns the cancel. @internal
 */
export function watchEagerPptxFonts(
  source: PptxFontsSource,
  handlers: {
    readonly resolved: (fonts: FontConfiguration) => void;
    readonly release: () => void;
  },
  waitMs = EAGER_FONTS_WAIT_MS
): () => void {
  let live = true;
  let released = false;
  const release = (): void => {
    if (!live || released) return;
    released = true;
    handlers.release();
  };
  const timer = setTimeout(() => {
    if (live && !released)
      console.warn(`[fonts] fonts did not settle in ${waitMs} ms; the deck opens without them`);
    release();
  }, waitMs);
  void resolveEagerPptxFonts(source).then((composed) => {
    clearTimeout(timer);
    if (live && composed) handlers.resolved(composed);
    release();
  });
  return () => {
    live = false;
    clearTimeout(timer);
  };
}

/**
 * A URL source fails when no bytes arrive for this long. The limit is on progress, not on the
 * whole download, so a slow connection that keeps sending bytes still finishes. @internal
 */
export const SOURCE_STALL_MS = 10_000;

/**
 * Fetch a URL source with the browser's `fetch`, as the caller wrote it. Rejects with a
 * `TimeoutError` when no response or no body bytes arrive for `stallMs`, and with the abort
 * when `signal` aborts. A URL from user input needs the caller's own check first: this adds
 * no allowlist. @internal
 */
export async function fetchPptxSource(
  url: string | URL,
  init: RequestInit | undefined,
  signal: AbortSignal,
  stallMs = SOURCE_STALL_MS
): Promise<Uint8Array> {
  signal.throwIfAborted();
  const controller = new AbortController();
  const cancelled = (): void => controller.abort(signal.reason);
  signal.addEventListener('abort', cancelled, { once: true });
  // An abort rejects the wait itself too: a body stream that ignores it cannot hold the open.
  const aborted = new Promise<never>((_, reject) =>
    controller.signal.addEventListener('abort', () => reject(controller.signal.reason))
  );
  aborted.catch(() => undefined);
  let timer: ReturnType<typeof setTimeout> | undefined;
  const arm = (): void => {
    clearTimeout(timer);
    timer = setTimeout(
      () =>
        controller.abort(
          new DOMException(`no bytes of the presentation for ${stallMs} ms`, 'TimeoutError')
        ),
      stallMs
    );
  };
  arm();
  try {
    const response = await Promise.race([
      fetch(String(url), { ...init, signal: controller.signal }),
      aborted,
    ]);
    if (!response.ok)
      throw new Error(`could not open the presentation: ${response.status} ${response.statusText}`);
    if (!response.body)
      return new Uint8Array(await Promise.race([response.arrayBuffer(), aborted]));
    const reader = response.body.getReader();
    const chunks: Uint8Array[] = [];
    let size = 0;
    try {
      for (;;) {
        arm();
        const next = await Promise.race([reader.read(), aborted]);
        if (next.done) break;
        chunks.push(next.value);
        size += next.value.byteLength;
      }
    } catch (cause) {
      void reader.cancel(cause).catch(() => undefined);
      throw cause;
    }
    const bytes = new Uint8Array(size);
    let at = 0;
    for (const chunk of chunks) (bytes.set(chunk, at), (at += chunk.byteLength));
    return bytes;
  } finally {
    clearTimeout(timer);
    signal.removeEventListener('abort', cancelled);
  }
}

/** Whether a fetch failure is the hook cancelling itself. @internal */
export function isAbortError(cause: unknown): boolean {
  return cause instanceof Error && cause.name === 'AbortError';
}
