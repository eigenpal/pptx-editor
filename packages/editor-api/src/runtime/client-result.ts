/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Copied from DOCX Editor's `packages/editor-api/src/runtime/client-result.ts` (commit 5993bd223945).

import { fail } from './errors.ts';

/**
 * A value a queued method promised to produce, readable after the next `sync()`.
 *
 * Not a `Promise`: a method call inside a batch has not been sent yet, so there is nothing that
 * could resolve on its own. Reading one early is `ValueNotLoaded`.
 *
 * @example
 * ```ts
 * const count = context.presentation.slides.getCount();
 * await context.sync();
 * console.log(count.value);
 * ```
 *
 * @public
 */
export class ClientResult<T> {
  #filled = false;
  #value: T | undefined;
  readonly #target: string;

  private constructor(target: string) {
    this.#target = target;
  }

  /** The value, once a `sync()` has filled it in. */
  get value(): T {
    if (!this.#filled) fail({ code: 'ValueNotLoaded', target: this.#target });
    return this.#value as T;
  }

  /** Whether the sync that fills this has happened. */
  get isLoaded(): boolean {
    return this.#filled;
  }

  /** @internal The box and the way to fill it, so only the creator can settle it. */
  static create<T>(target: string): { result: ClientResult<T>; fill: (value: T) => void } {
    const result = new ClientResult<T>(target);
    return {
      result,
      fill(value: T) {
        result.#filled = true;
        result.#value = value;
      },
    };
  }
}
