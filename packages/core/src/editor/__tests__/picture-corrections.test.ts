import { describe, expect, test } from 'bun:test';
import { chromeControlForSlot } from '../index.ts';
import { correctionGridRows, correctionSwatch } from '../picture-corrections.ts';

describe('the Corrections grid', () => {
  test('has a row per contrast and a column per brightness, from -40% to +40%', () => {
    const rows = correctionGridRows(chromeControlForSlot('image.corrections')!.options!);
    expect(rows.map((r) => r.contrast)).toEqual([-0.4, -0.2, 0, 0.2, 0.4]);
    for (const row of rows) {
      expect(row.cells.map((c) => c.brightness)).toEqual([-0.4, -0.2, 0, 0.2, 0.4]);
      expect(row.cells.every((c) => c.contrast === row.contrast)).toBe(true);
    }
  });

  test('leaves out an option that is not a pair', () => {
    const rows = correctionGridRows([
      { value: 'x', labelKey: 'pictureFormat.corrections' },
      { value: '20,0', labelKey: 'pictureFormat.corrections' },
    ]);
    expect(rows).toHaveLength(1);
    expect(rows[0]!.cells[0]!.brightness).toBe(0.2);
  });

  test('a sample with no correction keeps its colors; brightness and contrast change them', () => {
    expect(correctionSwatch(0, 0).map((p) => p.fill)).toEqual([
      '#8ec5e8',
      '#f6c445',
      '#5e9e4a',
      '#2f6b3a',
    ]);
    // Brightness +40% adds 102 of 255 to each channel, clamped at 255.
    expect(correctionSwatch(0.4, 0)[2]!.fill).toBe('#c4ffb0');
    // Contrast -40% scales each channel by 0.6 around mid-gray: 47 × 0.6 + 51 = 79.
    expect(correctionSwatch(0, -0.4)[3]!.fill).toBe('#4f7356');
  });
});
