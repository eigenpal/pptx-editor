// Copyright (c) 2026 Eigenpal, Inc. Licensed under Apache-2.0.
import { setTimeout as sleep } from 'node:timers/promises';
import { names, stable } from './common.mjs';

/** Wait for registry propagation before capture changes files or creates a branch. */
export async function waitForPublishedPackages(
  version,
  { request = fetch, pause = sleep, attempts = 18, delayMs = 10_000, log = console.log } = {}
) {
  if (!stable(version)) throw new Error('Invalid stable release');
  const pending = new Set(names);
  for (let attempt = 1; attempt <= attempts; attempt++) {
    await Promise.all(
      [...pending].map(async (name) => {
        let response;
        try {
          response = await request(
            `https://registry.npmjs.org/${encodeURIComponent(name)}/${version}`,
            {
              signal: AbortSignal.timeout(10_000),
              headers: { 'Cache-Control': 'no-cache' },
            }
          );
        } catch {
          // A network timeout is transient, but the attempt budget still applies.
          return;
        }
        if (!response.ok) {
          await response.body?.cancel();
          if (response.status === 404 || response.status === 429 || response.status >= 500) return;
          throw new Error(`Registry refused ${name}@${version}: HTTP ${response.status}`);
        }
        const metadata = await response.json();
        if (
          metadata.version !== version ||
          !metadata.dist?.integrity ||
          !metadata.dist?.tarball ||
          !metadata.gitHead
        )
          throw new Error(`Incomplete registry metadata for ${name}@${version}`);
        pending.delete(name);
      })
    );
    if (!pending.size) return;
    if (attempt < attempts) {
      log(
        `Waiting for npm propagation (${attempt}/${attempts}): ${[...pending].join(', ')}@${version}`
      );
      await pause(delayMs);
    }
  }
  throw new Error(
    `Published packages unavailable after ${attempts} attempts: ${[...pending].join(', ')}@${version}`
  );
}
