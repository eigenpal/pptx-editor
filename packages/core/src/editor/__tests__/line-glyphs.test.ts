import { describe, expect, test } from 'bun:test';
import { lineGlyphParts, lineSlotShown } from '../insert-controls.ts';
import { LINE_DASHES } from '../line-presets.ts';

describe('line glyphs', () => {
  test('every preset dash but solid draws a dashed stroke', () => {
    for (const d of LINE_DASHES) {
      const [part] = lineGlyphParts(d);
      expect(part?.kind).toBe('stroke');
      const dash = part?.kind === 'stroke' ? part.dash : undefined;
      if (d === 'solid') expect(dash).toBeUndefined();
      else expect(dash).toBeString();
    }
  });

  test('a trigger without a value shows its slot default', () => {
    expect(lineSlotShown('insert.line', undefined)).toBe('line');
    expect(lineSlotShown('shape.lineEnd', undefined)).toBe('none');
    expect(lineSlotShown('shape.lineDash', undefined)).toBe('solid');
    expect(lineSlotShown('table.borderStyle', null)).toBe('solid');
    expect(lineSlotShown('shape.lineDash', 'dashDot')).toBe('dashDot');
  });

  test('a dash outside the menu still draws dashed on the trigger', () => {
    expect(lineSlotShown('shape.lineDash', 'sysDashDotDot')).toBe('sysDashDotDot');
    expect(lineGlyphParts('sysDashDotDot')[0]).toMatchObject({ dash: '3 5 0 5 0 5' });
    expect(lineSlotShown('shape.lineDash', 'custom')).toBe('dash');
    expect(lineSlotShown('insert.line', 'elbow')).toBe('elbow');
  });
});
