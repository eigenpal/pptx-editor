// The steps of an open before the deck mounts: let the loading screen paint, parse, wait for
// the fonts, and register the faces. Each wait has a limit, so an open never waits forever.

import { registerFontFaces, type FontFaceHandle } from '../output/index.ts';
import {
  openPresentation,
  type DocumentSession,
  type OpenError,
  type OpenOptions,
} from './document-session.ts';
import { FONT_WAIT_MS, type HostFonts } from './host-fonts.ts';
import { yieldToPaint } from './loading.ts';
import { within } from './within.ts';

/** The result of `prepareDeck`. */
export type PreparedDeck =
  | { readonly kind: 'stale' }
  | { readonly kind: 'refused'; readonly error: OpenError }
  | {
      readonly kind: 'ready';
      readonly session: DocumentSession;
      readonly faces: FontFaceHandle | null;
      /** `session.fontGeneration` when `faces` registered: a later one needs registering. */
      readonly generation: number;
    };

/**
 * Parse `bytes` and wait for its fonts. One wait covers the text engine, the packaged
 * substitutes, and the host's fonts: after `FONT_WAIT_MS`, the deck shows with what arrived,
 * and the rest installs when it lands. `step` names each phase, for the log of a failed open.
 * A session that a later load supersedes, or that a throw interrupts, is disposed here.
 */
export async function prepareDeck(
  bytes: Uint8Array,
  options: OpenOptions,
  hostFonts: HostFonts,
  stale: () => boolean,
  step: (phase: string) => void
): Promise<PreparedDeck> {
  // Let the loading screen paint before the parse blocks the main thread (a large deck).
  await yieldToPaint();
  if (stale()) return { kind: 'stale' };
  step('parse');
  // The parse does not wait for the text engine: it and the host's fonts wait side by side.
  const r = await openPresentation(bytes, { ...options, shapingWaitMs: 0 });
  if (!r.ok) return stale() ? { kind: 'stale' } : { kind: 'refused', error: r.error };
  const session = r.session;
  try {
    step('fonts');
    // One wait from the end of the parse, for the text engine and the host's fonts at once.
    const left = Math.max(0, session.parsedAt + FONT_WAIT_MS - performance.now());
    const shaping = session.shaping;
    const [shaped] = await Promise.all([
      shaping ? within(shaping, left, undefined) : undefined,
      stale() ? undefined : hostFonts.open(session, stale, left),
    ]);
    // Shaped measurement that arrived in time is done. One that failed stays, so the editor
    // reports why.
    if (shaped === true) session.shaping = null;
    if (stale()) return (hostFonts.abandon(session), session.dispose(), { kind: 'stale' });
    step('faces');
    const generation = session.fontGeneration;
    if (typeof document === 'undefined') return { kind: 'ready', session, faces: null, generation };
    const faces = registerFontFaces(document, session.faces);
    await faces.ready;
    if (stale()) {
      hostFonts.abandon(session);
      return (faces.dispose(), session.dispose(), { kind: 'stale' });
    }
    return { kind: 'ready', session, faces, generation };
  } catch (cause) {
    hostFonts.abandon(session);
    session.dispose();
    throw cause;
  }
}
