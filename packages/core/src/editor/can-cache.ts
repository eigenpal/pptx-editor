// The answers of `can` and the formatting state for toggles, kept until the state a command
// compiles against moves. Toolbars ask for every slot on every change, so each answer is
// computed once per state. Moved out of `pptx-editor.ts` (the 1000-line cap).

import type { CanResult } from '../contracts/index.ts';

/** `can` answers and one derived value, both reset when `state()` changes. @internal */
export class CanCache<Active> {
  readonly #answers = new Map<string, CanResult>();
  #state: readonly unknown[] = [];
  #active: Active | null = null;

  constructor(private readonly state: () => readonly unknown[]) {}

  /** Drop every answer when the state moved since the last call. */
  sync(): void {
    const now = this.state();
    if (now.length === this.#state.length && now.every((v, i) => v === this.#state[i])) return;
    this.#state = now;
    this.#answers.clear();
    this.#active = null;
  }

  /** The cached answer for `key`, or `compute()` cached. */
  answer(key: string, compute: () => CanResult): CanResult {
    this.sync();
    const hit = this.#answers.get(key);
    if (hit) return hit;
    const result = compute();
    this.#answers.set(key, result);
    return result;
  }

  /** The derived value for the current state, or `compute()` cached. */
  active(compute: () => Active): Active {
    this.sync();
    return (this.#active ??= compute());
  }
}
