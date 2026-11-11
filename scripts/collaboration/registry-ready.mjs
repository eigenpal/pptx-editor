// Copyright (c) 2026 Eigenpal, Inc. Licensed under Apache-2.0.
// Retry policy follows DOCX Editor scripts/collaboration/registry.mjs at
// 2f113bf5bb3b3283969df7d96f2c019ee95b2edf. Exact versions remain required.
import { setTimeout as sleep } from 'node:timers/promises';
import { names, stable } from './common.mjs';

export const PUBLICATION_TIMEOUT_MS = 10 * 60_000;
const TRANSIENT = new Set([404, 408, 429, 500, 502, 503, 504]);

function retryAfter(response, now) {
  const value = response.headers.get('retry-after');
  if (!value) return 0;
  const seconds = Number(value);
  return Math.max(0, Number.isFinite(seconds) ? seconds * 1000 : Date.parse(value) - now) || 0;
}

/** Wait for every exact package version under one deadline before capture changes files. */
export async function waitForPublishedPackages(
  version,
  {
    request = fetch,
    pause = sleep,
    now = Date.now,
    random = Math.random,
    timeoutMs = PUBLICATION_TIMEOUT_MS,
    log = console.log,
  } = {}
) {
  if (!stable(version)) throw new Error('Invalid stable release');
  const deadline = now() + timeoutMs;
  const pending = new Set(names);
  const controller = new AbortController();
  let attempt = 0;
  try {
    while (pending.size && now() < deadline) {
      let retryDelay = 0;
      await Promise.all(
        [...pending].map(async (name) => {
          let response, metadata;
          try {
            response = await request(
              `https://registry.npmjs.org/${encodeURIComponent(name)}/${version}`,
              {
                signal: AbortSignal.any([
                  controller.signal,
                  AbortSignal.timeout(Math.max(1, Math.min(30_000, deadline - now()))),
                ]),
                headers: { 'cache-control': 'no-cache', accept: 'application/json' },
                cache: 'no-store',
              }
            );
            // Reading the body can fail after headers arrive. It shares the request timeout.
            if (response.ok) metadata = await response.json();
          } catch {
            return;
          }
          if (!response.ok) {
            retryDelay = Math.max(retryDelay, retryAfter(response, now()));
            await response.body?.cancel();
            if (TRANSIENT.has(response.status)) return;
            throw new Error(`Registry refused ${name}@${version}: HTTP ${response.status}`);
          }
          // Invalid release data is permanent, not a network failure.
          if (
            metadata?.version !== version ||
            !metadata.dist?.integrity ||
            !metadata.dist?.tarball ||
            !metadata.gitHead
          )
            throw new Error(`Incomplete registry metadata for ${name}@${version}`);
          pending.delete(name);
        })
      );
      if (!pending.size) return;
      const remaining = deadline - now();
      if (remaining <= 0) break;
      const backoff = Math.min(30_000, 5000 * 2 ** Math.min(attempt++, 3));
      const delay = Math.min(
        remaining,
        Math.max(retryDelay, Math.round(backoff * (0.8 + 0.2 * random())))
      );
      log(
        `Waiting for npm propagation: ${[...pending].join(', ')}@${version}; ${Math.ceil(remaining / 1000)}s remaining, retry in ${Math.ceil(delay / 1000)}s`
      );
      await pause(delay);
    }
    throw new Error(
      `Published packages unavailable before deadline: ${[...pending].join(', ')}@${version}`
    );
  } finally {
    // Stop other pending requests if one package fails permanently.
    controller.abort();
  }
}
