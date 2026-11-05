/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// License keys stay in this process. Licensing makes no network requests.

/** Accepted by every pro entry point. @public */
export interface ProLicenseOptions {
  /**
   * Your license key. Optional: evaluation use is permitted, production use requires a
   * commercial agreement under the EigenPal Pro License (see LICENSE.md).
   */
  readonly licenseKey?: string;
}

let storedKey: string | undefined;

/** Remember the most recent key an entry point was constructed with. */
export function rememberLicenseKey(key: string | undefined): void {
  if (key !== undefined) storedKey = key;
}

/** The most recent license key supplied to an entry point. */
export function currentLicenseKey(): string | undefined {
  return storedKey;
}
