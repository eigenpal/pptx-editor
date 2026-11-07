import { describe, expect, test } from 'bun:test';
import { SAVE_VERIFIED_DIAGRAM_EDITS } from '../../docs/site/data/save-verified-fixtures.ts';
import { DIAGRAM_EDITS } from './diagram-edits.ts';
import { runFeatureEdit } from './feature-edit-kit.ts';

// Each recipe types in a diagram shape, or adds or deletes one, saves, and reopens. The data
// model and the saved drawing must both hold the edit and name only points that exist, and the
// copy must open with no repair prompt: its fingerprint is in `SAVE_VERIFIED_DIAGRAM_EDITS`. A new fingerprint needs
// a new prompt check (CONTRIBUTING.md, "Fidelity harness").
describe('diagram edits survive save and reopen', () => {
  for (const [id, recipe] of Object.entries(DIAGRAM_EDITS)) {
    test(id, async () => {
      const result = await runFeatureEdit(id, recipe);
      expect(result.failure).toBeUndefined();
      expect(result.ok).toBe(true);
      expect(result.fingerprint).toBe(SAVE_VERIFIED_DIAGRAM_EDITS[id]);
    });
  }

  test('every verified copy has a recipe', () => {
    const orphans = Object.keys(SAVE_VERIFIED_DIAGRAM_EDITS).filter((id) => !DIAGRAM_EDITS[id]);
    expect(orphans).toEqual([]);
  });
});
