// The custom color field shows the current color while it is empty: a hex choice, a theme
// choice as its swatch renders, or black (the browser selector's default) with no color.

import { describe, expect, test } from 'bun:test';
import type { ThemeSwatch } from '../../contracts/index.ts';
import { customHexCurrent } from '../color-fields.ts';

const rows: ThemeSwatch[][] = [
  [
    { slot: 'accent1', color: '#4472C4', shift: 0 } as ThemeSwatch,
    { slot: 'accent2', color: '#ED7D31', shift: 0 } as ThemeSwatch,
  ],
  [{ slot: 'accent1', color: '#DAE3F3', shift: 80 } as ThemeSwatch],
];

describe('customHexCurrent', () => {
  test('a hex choice shows itself', () => {
    expect(customHexCurrent('#c00000', rows)).toBe('C00000');
  });
  test('a theme choice shows its swatch as rendered', () => {
    expect(customHexCurrent({ theme: 'accent2' }, rows)).toBe('ED7D31');
    expect(customHexCurrent({ theme: 'accent1', shift: 80 }, rows)).toBe('DAE3F3');
  });
  test('no color, or a theme color the grid does not show, shows black', () => {
    expect(customHexCurrent(null, rows)).toBe('000000');
    expect(customHexCurrent(undefined, rows)).toBe('000000');
    expect(customHexCurrent({ theme: 'accent6' }, rows)).toBe('000000');
  });
});
