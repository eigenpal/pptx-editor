// Development warnings for runs that cannot do what the caller meant. After
// `packages/core/src/editor/dev-warnings.ts`: each message prints once, never in a production build.

/** Vite's flag when it is a boolean, else the Node convention; production when neither says. */
function isDevelopment(): boolean {
  let dev: unknown;
  try {
    dev = (import.meta as { readonly env?: { readonly DEV?: unknown } }).env?.DEV;
  } catch {
    dev = undefined;
  }
  if (typeof dev === 'boolean') return dev;
  const hasProcess = typeof process !== 'undefined';
  return hasProcess && process.env.NODE_ENV !== 'production';
}

const printed = new Set<string>();

/** Print `message` once, in development builds only. */
function devWarning(message: string): void {
  if (!isDevelopment() || printed.has(message)) return;
  printed.add(message);
  console.warn(`[pptx-editor] ${message}`);
}

/** Forget printed messages. Tests only. @internal */
export function resetDevWarnings(): void {
  printed.clear();
}

/**
 * A run ended with writes that no `sync()` sent: they were dropped. `refusals` are the argument
 * refusals among them, which that sync would have given; the warning names them so they are not
 * lost.
 */
export function warnUnsynced(writes: number, refusals: readonly Error[] = []): void {
  const lost = refusals.map((error) => {
    const e = error as { code?: string; debugInfo?: { errorLocation?: string } };
    const where = e.debugInfo?.errorLocation ? ` at ${e.debugInfo.errorLocation}` : '';
    return ` ${e.code ?? 'Error'}${where}: ${error.message}`;
  });
  devWarning(
    `run() ended with ${writes} queued ${writes === 1 ? 'write' : 'writes'} that no sync() sent, ` +
      `so ${writes === 1 ? 'it' : 'they'} did not apply. Call \`await context.sync()\` before the ` +
      'callback returns.' +
      (lost.length === 0
        ? ''
        : ` That sync would have refused with ${lost.length === 1 ? 'this error' : 'these errors'}:` +
          lost.join(';'))
  );
}
