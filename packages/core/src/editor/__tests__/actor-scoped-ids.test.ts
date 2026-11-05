// Copied from DOCX Editor: packages/core/src/store/__tests__/actor-scoped-ids.test.ts at a59ace33
// (the format-neutral cases only: stripes, dense ids, and the striped ceiling). The cases that
// drive DOCX Editor's WordprocessingML minters are left out; PPTX Editor's mint sites are
// tested in `replica-ids.test.ts` and in `packages/pro/src/__tests__/yjs-ids.test.ts`.
//
// Actor-scoped decimal id allocation: stripes, solo density, and hostile input.

import { describe, expect, test } from 'bun:test';
import {
  ACTOR_ID_STRIPE,
  actorStripe,
  MAX_DECIMAL_ID,
  nextDenseDecimalId,
  nextStripedDecimalId,
} from '../../automation/commands/actor-scoped-ids.ts';

describe('actor stripes are disjoint', () => {
  test('alice and bob hash to different residues', () => {
    expect(actorStripe('alice')).not.toBe(actorStripe('bob'));
    expect(actorStripe('alice')).toBeLessThan(ACTOR_ID_STRIPE);
    expect(actorStripe('bob')).toBeLessThan(ACTOR_ID_STRIPE);
  });

  test('two actors minting from the same used set never share an id', () => {
    const used = new Set<string>();
    const alice: string[] = [];
    const bob: string[] = [];
    for (let index = 0; index < 20; index += 1) {
      const left = nextStripedDecimalId(used, 'alice', MAX_DECIMAL_ID);
      used.add(left);
      alice.push(left);
      const right = nextStripedDecimalId(used, 'bob', MAX_DECIMAL_ID);
      used.add(right);
      bob.push(right);
    }
    expect(new Set([...alice, ...bob]).size).toBe(40);
    const aliceStripe = actorStripe('alice');
    for (const id of alice) expect(Number(id) % ACTOR_ID_STRIPE).toBe(aliceStripe);
    const bobStripe = actorStripe('bob');
    for (const id of bob) expect(Number(id) % ACTOR_ID_STRIPE).toBe(bobStripe);
  });

  test('the stripe space is wide enough that a real room rarely shares a class', () => {
    // Same-stripe actors mint the SAME id, so the count is a birthday problem, not a style
    // choice. This was 16, where a five-person room collided half the time. The bound below is
    // what keeps that from being reintroduced as a "keep ids small" tidy-up.
    const room = 8;
    const pairs = (room * (room - 1)) / 2;
    // Rough birthday bound: expected colliding pairs is pairs / classes. At 16 this was 1.75.
    expect(pairs / ACTOR_ID_STRIPE).toBeLessThan(0.005);
  });

  test('actor ids of the shape hosts mint spread across the space', () => {
    // Deterministic sample, so this measures our modulus rather than the platform's RNG.
    const stripes = new Set<number>();
    for (let index = 0; index < 200; index += 1) {
      stripes.add(
        actorStripe(`Reviewer ${index}:0f1e2d3c-4b5a-6978-8796-a5b4c3d2e1f${index % 10}`)
      );
    }
    // At 65,536 classes, 200 draws expect ~0.3 collisions; demanding near-perfect spread here
    // would be testing FNV, so this only catches a modulus that collapses the space.
    expect(stripes.size).toBeGreaterThan(190);
  });
});

describe('solo allocation stays dense', () => {
  test('dense ids start at zero and climb by one', () => {
    expect(nextDenseDecimalId(-1, undefined, MAX_DECIMAL_ID)).toBe('0');
    expect(nextDenseDecimalId(0, undefined, MAX_DECIMAL_ID)).toBe('1');
    expect(nextDenseDecimalId(7, undefined, MAX_DECIMAL_ID)).toBe('8');
  });
});

describe('ceiling and hostile input stay ignored for seeding', () => {
  test('a striped mint past the ceiling refuses rather than colliding', () => {
    const used = new Set<string>();
    const stripe = actorStripe('alice');
    used.add(String(stripe));
    expect(() => nextStripedDecimalId(used, 'alice', stripe)).toThrow('no free decimal id');
  });
});
