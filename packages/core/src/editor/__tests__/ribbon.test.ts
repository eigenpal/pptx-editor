import { describe, expect, test } from 'bun:test';
import {
  CHROME_GROUPS,
  EMPTY_EDITOR_SNAPSHOT,
  RIBBON_TABS,
  ribbonTabGroups,
  ribbonTabs,
  type ChromeGroup,
  type EditorSnapshot,
  type RibbonTab,
} from '../index.ts';

// The ribbon arrangement.

const ids = (layout: ReturnType<typeof ribbonTabs>) => layout.map((l) => l.tab.id);

describe('ribbon tabs', () => {
  test('without a selection: the Quick Access Toolbar and the main tabs', () => {
    expect(ids(ribbonTabs(EMPTY_EDITOR_SNAPSHOT))).toEqual([
      'quickAccess',
      'home',
      'insert',
      'design',
      'slideShow',
      'view',
    ]);
  });

  test('Shape Format shows while shapes are selected; Picture and Table tabs wait for theirs', () => {
    const shapes = {
      ...EMPTY_EDITOR_SNAPSHOT,
      selection: { kind: 'shapes', nodeIds: ['n1'] },
    } as EditorSnapshot;
    const layout = ribbonTabs(shapes);
    expect(ids(layout)).toContain('shapeFormat');
    expect(ids(layout)).not.toContain('pictureFormat');
    expect(layout.find((l) => l.tab.id === 'shapeFormat')!.groups.map((g) => g.id)).toEqual([
      'insert',
      'shape',
      'arrange',
    ]);
    const picture = { ...shapes, image: {} } as EditorSnapshot;
    expect(ids(ribbonTabs(picture))).toContain('pictureFormat');
    expect(ids(ribbonTabs(picture))).not.toContain('shapeFormat');
    const table = { ...shapes, table: {} } as EditorSnapshot;
    expect(ids(ribbonTabs(table)).slice(-2)).toEqual(['tableDesign', 'tableLayout']);
    expect(ids(ribbonTabs(table))).not.toContain('shapeFormat');
    // Text editing in a shape shows Shape Format; text editing in a table cell does not.
    const text = { ...EMPTY_EDITOR_SNAPSHOT, selection: { kind: 'text' } } as EditorSnapshot;
    expect(ids(ribbonTabs(text))).toContain('shapeFormat');
    const cell = { ...text, table: {} } as EditorSnapshot;
    expect(ids(ribbonTabs(cell))).toContain('tableLayout');
    expect(ids(ribbonTabs(cell))).not.toContain('shapeFormat');
  });

  test('every registered control is on a tab, and a tab adds no slot of its own', () => {
    const placed = new Set(
      RIBBON_TABS.flatMap((tab) =>
        ribbonTabGroups(tab).flatMap((g) => g.controls.map((c) => `${g.id}.${c.id}`))
      )
    );
    const registered = new Set(
      CHROME_GROUPS.flatMap((g) => g.controls.map((c) => `${g.id}.${c.id}`))
    );
    // The section commands live on the slide rail's menu.
    const offRibbon = [...registered].filter((slot) => !placed.has(slot));
    expect(offRibbon.every((slot) => slot.startsWith('section.'))).toBe(true);
    for (const slot of placed) expect(registered.has(slot)).toBe(true);
  });

  test('a picked group keeps only its picked controls, in the pick order', () => {
    const groups = [
      { id: 'slide', labelKey: 'toolbar.groups.more', controls: [{ id: 'a' }, { id: 'b' }] },
    ] as unknown as ChromeGroup[];
    const tab = {
      id: 'home',
      labelKey: 'ribbon.tabs.home',
      groups: [{ id: 'slide', controls: ['b', 'missing'] }, 'nothing'],
    } as unknown as RibbonTab;
    expect(ribbonTabGroups(tab, groups).map((g) => g.controls.map((c) => c.id))).toEqual([['b']]);
  });

  test('a tab names each group once, as the bar keys its groups by id', () => {
    for (const tab of RIBBON_TABS) {
      const groupIds = ribbonTabGroups(tab).map((g) => g.id);
      expect(new Set(groupIds).size).toBe(groupIds.length);
    }
  });

  test('Arrange follows the reference order on Home, Shape Format, and Picture Format', () => {
    const order = ['order', 'selectionPane', 'align', 'group', 'ungroup', 'rotate'];
    for (const id of ['home', 'shapeFormat', 'pictureFormat']) {
      const tab = RIBBON_TABS.find((t) => t.id === id)!;
      const arrange = ribbonTabGroups(tab).find((g) => g.id === 'arrange')!;
      expect(arrange.controls.map((c) => c.id)).toEqual(order);
    }
  });
});
