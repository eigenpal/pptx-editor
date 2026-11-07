import { describe, expect, test } from 'bun:test';
import { SAVE_VERIFIED_COMMENT_EDITS } from '../../docs/site/data/save-verified-fixtures.ts';
import { COMMENT_EDITS, commentFingerprint } from './comment-edits.ts';
import { runFeatureEdit } from './feature-edit-kit.ts';

// Each recipe edits comments, saves, and reopens. The threads must come back as edited, and the
// copy must open with no repair prompt: its fingerprint is in `SAVE_VERIFIED_COMMENT_EDITS`. A
// new fingerprint needs a new prompt check (CONTRIBUTING.md, "Fidelity harness").
describe('comment edits survive save and reopen', () => {
  for (const [id, recipe] of Object.entries(COMMENT_EDITS)) {
    test(id, async () => {
      const result = await runFeatureEdit(id, recipe);
      expect(result.failure).toBeUndefined();
      expect(result.ok).toBe(true);
      expect(commentFingerprint(id, result)).toBe(SAVE_VERIFIED_COMMENT_EDITS[id]);
    });
  }

  test('every verified copy has a recipe', () => {
    const orphans = Object.keys(SAVE_VERIFIED_COMMENT_EDITS).filter((id) => !COMMENT_EDITS[id]);
    expect(orphans).toEqual([]);
  });
});
