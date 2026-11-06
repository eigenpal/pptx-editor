import { describe, expect, test } from 'bun:test';
import { runFeatureEdit } from './feature-edit-kit.ts';
import { FEATURE_EDITS } from './feature-edits.ts';

// Each recipe edits a fixture of its feature row through the public editor API, saves, reopens,
// and finds the edit again. Its check must fail on the unedited deck, a second save must give the
// same package, and the edit must change only the parts it owns.
describe('feature edits survive save and reopen', () => {
  for (const [id, recipe] of Object.entries(FEATURE_EDITS)) {
    test(id, async () => {
      const result = await runFeatureEdit(id, recipe);
      expect(result.failure).toBeUndefined();
      expect(result.ok).toBe(true);
    });
  }
});
