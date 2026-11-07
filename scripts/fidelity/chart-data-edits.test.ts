import { describe, expect, test } from 'bun:test';
import { SAVE_VERIFIED_CHART_DATA_EDITS } from '../../docs/site/data/save-verified-fixtures.ts';
import { CHART_DATA_EDITS } from './chart-data-edits.ts';
import { runFeatureEdit } from './feature-edit-kit.ts';

// Each recipe edits the data of a chart, saves, and reopens. The chart part's caches and
// references and the embedded workbook's cells must agree, and the copy must open with no
// repair prompt: its fingerprint is in `SAVE_VERIFIED_CHART_DATA_EDITS`. A new fingerprint needs
// a new prompt check (CONTRIBUTING.md, "Fidelity harness").
describe('chart data edits survive save and reopen', () => {
  for (const [id, recipe] of Object.entries(CHART_DATA_EDITS)) {
    test(id, async () => {
      const result = await runFeatureEdit(id, recipe);
      expect(result.failure).toBeUndefined();
      expect(result.ok).toBe(true);
      expect(result.fingerprint).toBe(SAVE_VERIFIED_CHART_DATA_EDITS[id]);
    });
  }

  test('every verified copy has a recipe', () => {
    const orphans = Object.keys(SAVE_VERIFIED_CHART_DATA_EDITS).filter(
      (id) => !CHART_DATA_EDITS[id]
    );
    expect(orphans).toEqual([]);
  });
});
