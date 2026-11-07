import { describe, expect, test } from 'bun:test';
import { pptxFeatures } from '../../docs/site/data/pptx-features.ts';
import { SAVE_VERIFIED_KEPT_CONTENT_EDITS } from '../../docs/site/data/save-verified-fixtures.ts';
import { runFeatureEdit } from './feature-edit-kit.ts';
import { KEPT_CONTENT_EDITS, KEPT_CONTENT_ROWS } from './kept-content-edits.ts';

// Each recipe duplicates, pastes, or groups the kept content of a `preserved` row, saves, and
// reopens. The kept objects must come back with the same XML and the same parts, and the copy
// must open with no repair prompt: its fingerprint is in `SAVE_VERIFIED_KEPT_CONTENT_EDITS`. A new
// fingerprint needs a new prompt check (CONTRIBUTING.md, "Fidelity harness").
describe('kept content survives structural edits', () => {
  for (const [id, recipe] of Object.entries(KEPT_CONTENT_EDITS)) {
    test(id, async () => {
      const result = await runFeatureEdit(id, recipe);
      expect(result.failure).toBeUndefined();
      expect(result.ok).toBe(true);
      expect(result.fingerprint).toBe(SAVE_VERIFIED_KEPT_CONTENT_EDITS[id]);
    });
  }
});

describe('structural recipes', () => {
  test('cover every row that keeps content it does not edit', () => {
    // A row that now edits its content (charts edit their data) keeps its structural recipes:
    // the content it does not edit (the chart's formatting) must still survive them.
    const missing = pptxFeatures
      .filter((f) => f.roundTrip === 'preserved' && !KEPT_CONTENT_ROWS.includes(f.id))
      .map((f) => f.id);
    expect(missing).toEqual([]);
    const ids = new Set(pptxFeatures.map((f) => f.id));
    expect(KEPT_CONTENT_ROWS.filter((id) => !ids.has(id))).toEqual([]);
  });

  test('every verified copy has a recipe', () => {
    const orphans = Object.keys(SAVE_VERIFIED_KEPT_CONTENT_EDITS).filter(
      (id) => !KEPT_CONTENT_EDITS[id]
    );
    expect(orphans).toEqual([]);
  });
});
