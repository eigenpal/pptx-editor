// A border on every cell of a large table parses its line fragment once, not once for each
// cell edge. The parsed fragment is shared, so the result stays the same.

import { expect, test } from 'bun:test';
import { bigTable } from '../../../../../scripts/stress/interaction-decks.ts';
import { specFromXml } from '../../automation/commands/shape-commands.ts';
import { createPptxEditor } from '../index.ts';

test('a fragment parses once, and a long fragment is not kept', () => {
  const ln = '<a:lnL w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:lnL>';
  expect(specFromXml(ln)).toBe(specFromXml(ln));
  const long = `<a:t>${'x'.repeat(5_000)}</a:t>`;
  expect(specFromXml(long)).not.toBe(specFromXml(long));
  expect(specFromXml(long)).toEqual(specFromXml(long));
});

test('borders on every cell of a 100 x 30 table apply and save', async () => {
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await editor.load(bigTable(100, 30))).ok).toBe(true);
  const frame = editor.session!.layout(0)!.items[0]!.element.nodeId;
  const at = { paragraph: 0, offset: 0 };
  editor.setSelection({
    kind: 'text',
    nodeId: frame,
    cell: { row: 0, column: 0 },
    anchor: at,
    focus: at,
  });
  expect(editor.exec({ type: 'selectTableRegion', region: 'table' }).ok).toBe(true);
  const command = {
    type: 'setTableBorders',
    scope: 'all',
    spec: { color: '#0000FF', width: 2 },
  } as const;
  expect(editor.can(command).ok).toBe(true);
  expect(editor.exec(command).ok).toBe(true);
  expect(editor.save()?.byteLength).toBeGreaterThan(0);
  editor.destroy();
});
