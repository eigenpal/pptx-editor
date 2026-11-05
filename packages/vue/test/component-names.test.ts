// Each exported Vue component carries its export name as its runtime `name`, so the component
// tree in the browser's developer tools shows the documented names.

import { expect, test } from 'bun:test';
import * as vue from '../src/index.ts';

test('every exported component is named as its export', () => {
  const wrong: string[] = [];
  for (const [exported, value] of Object.entries(vue)) {
    if (!value || typeof value !== 'object' || !('setup' in value)) continue;
    const name = (value as { name?: string }).name;
    if (name !== exported) wrong.push(`${exported}: ${String(name)}`);
  }
  expect(wrong).toEqual([]);
});
