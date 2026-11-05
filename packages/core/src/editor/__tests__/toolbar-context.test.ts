import { describe, expect, test } from 'bun:test';
import {
  EMPTY_EDITOR_SNAPSHOT,
  TOOLBAR_BAR_GROUPS,
  TOOLBAR_PICTURE_GROUPS,
  arrangeToolbarGroups,
  formattingBarChromeGroups,
  isToolbarBarSlot,
  toolbarContextOf,
  toolbarGroupsFor,
  type EditorSnapshot,
  type SelectedImageState,
  type TableContext,
  type ToolbarContext,
} from '../index.ts';

// The preset toolbar shows only the groups that act on the selection.

const snap = (over: Partial<EditorSnapshot>): EditorSnapshot => ({
  ...EMPTY_EDITOR_SNAPSHOT,
  ...over,
});
const at = { paragraph: 0, offset: 0 };
const ids = (context: ToolbarContext) => toolbarGroupsFor(context).map((g) => g.id);
const slots = (context: ToolbarContext) =>
  toolbarGroupsFor(context).flatMap((g) => g.controls.map((c) => `${g.id}.${c.id}`));

describe('toolbarContextOf', () => {
  test('names the selection: nothing, text, shapes, pictures, or a table', () => {
    expect(toolbarContextOf(snap({}))).toBe('none');
    expect(toolbarContextOf(snap({ selection: { kind: 'shapes', nodeIds: [] } }))).toBe('none');
    expect(
      toolbarContextOf(snap({ selection: { kind: 'text', nodeId: 'a', anchor: at, focus: at } }))
    ).toBe('text');
    const shapes = { kind: 'shapes', nodeIds: ['a'] } as const;
    expect(toolbarContextOf(snap({ selection: shapes }))).toBe('shape');
    expect(toolbarContextOf(snap({ selection: shapes, image: {} as SelectedImageState }))).toBe(
      'picture'
    );
    // Text in a table cell reads as the table.
    const cell = { row: 0, column: 0 };
    expect(
      toolbarContextOf(
        snap({
          selection: { kind: 'text', nodeId: 'a', anchor: at, focus: at, cell },
          table: {} as TableContext,
        })
      )
    ).toBe('table');
  });
});

describe('toolbarGroupsFor', () => {
  test('nothing selected: Insert, the slide, and Find', () => {
    expect(ids('none')).toEqual(['history', 'insert', 'slide', 'find']);
    expect(slots('none').filter((s) => s.startsWith('slide.'))).toEqual([
      'slide.background',
      'slide.layout',
    ]);
  });

  test('text shows the text groups and none of the shape, arrange, or picture groups', () => {
    expect(ids('text')).toEqual([
      'history',
      'font',
      'text',
      'paragraph',
      'list',
      'insert',
      'format',
    ]);
  });

  test('a shape adds its fill, line, and arrange groups to the basic text groups', () => {
    expect(ids('shape')).toEqual([
      'history',
      'font',
      'text',
      'paragraph',
      'insert',
      'shape',
      'format',
      'arrange',
      'more',
    ]);
  });

  test('a picture shows the picture group and none of the text or shape groups', () => {
    expect(ids('picture')).toEqual(['history', 'insert', 'image', 'arrange', 'more']);
  });

  test('a table selected as an object keeps Delete, Duplicate, and Arrange', () => {
    const frame = { kind: 'shapes', nodeIds: ['t'] } as const;
    expect(toolbarContextOf(snap({ selection: frame, table: {} as TableContext }))).toBe(
      'tableFrame'
    );
    expect(ids('tableFrame')).toEqual([
      'history',
      'font',
      'text',
      'paragraph',
      'insert',
      'format',
      'arrange',
      'more',
      'table',
    ]);
  });

  test('a focused notes field shows Undo, Redo, the four notes formats, and Find', () => {
    expect(slots('notes')).toEqual([
      'history.undo',
      'history.redo',
      'text.bold',
      'text.italic',
      'text.underline',
      'text.strike',
      'find.find',
      'find.replace',
    ]);
  });

  test('a table adds the table group after the text groups', () => {
    expect(ids('table').at(-1)).toBe('table');
    expect(slots('table')).not.toContain('table.insert');
  });

  test('the controls that the Font dialog and the Format menu set stay off every bar', () => {
    const off = ['script.super', 'script.sub', 'text.characterSpacing', 'paragraph.columns'];
    for (const context of ['none', 'text', 'shape', 'picture', 'table'] as const)
      for (const slot of off) expect(slots(context)).not.toContain(slot);
    // A host still places them: they go after the groups, not into a hidden group.
    expect(isToolbarBarSlot('script.super')).toBe(false);
    expect(isToolbarBarSlot('slide.background')).toBe(true);
  });

  test('the same array comes back for the same context', () => {
    expect(toolbarGroupsFor('text')).toBe(toolbarGroupsFor('text'));
  });
});

describe('the bar across selections', () => {
  const order = (context: ToolbarContext, after: string) =>
    arrangeToolbarGroups(new Map([['acme', { id: 'acme', after }]]), toolbarGroupsFor(context)).map(
      (e) => (e.kind === 'host' ? e.override.id : e.group.id)
    );

  test('a host group whose anchor is not shown follows the nearest shown group before it', () => {
    expect(order('text', 'list')).toEqual([
      'history',
      'font',
      'text',
      'paragraph',
      'list',
      'acme',
      'insert',
      'format',
    ]);
    // A shape shows no list group: the host group follows the paragraph group, not the end.
    expect(order('shape', 'list').indexOf('acme')).toBe(
      order('shape', 'list').indexOf('paragraph') + 1
    );
    // A picture shows no text group: the host group follows Undo and Redo.
    expect(order('picture', 'text').slice(0, 2)).toEqual(['history', 'acme']);
  });

  test('contextual false keeps a host group at its anchor', () => {
    const full = toolbarGroupsFor('text', { contextual: false });
    const placed = arrangeToolbarGroups(
      new Map([['acme', { id: 'acme', after: 'list' }]]),
      full
    ).map((e) => (e.kind === 'host' ? e.override.id : e.group.id));
    expect(placed[placed.indexOf('list') + 1]).toBe('acme');
  });

  test('contextual false shows every bar group, with the picture or table group after them', () => {
    expect(toolbarGroupsFor('none', { contextual: false })).toBe(TOOLBAR_BAR_GROUPS);
    expect(toolbarGroupsFor('text', { contextual: false })).toBe(TOOLBAR_BAR_GROUPS);
    expect(toolbarGroupsFor('picture', { contextual: false })).toBe(TOOLBAR_PICTURE_GROUPS);
    expect(toolbarGroupsFor('table', { contextual: false }).at(-1)?.id).toBe('table');
    expect(toolbarGroupsFor('tableFrame', { contextual: false }).at(-1)?.id).toBe('table');
  });

  test('the formatting bar has no empty group', () => {
    for (const g of formattingBarChromeGroups()) expect(g.controls.length).toBeGreaterThan(0);
    expect(formattingBarChromeGroups().map((g) => g.id)).not.toContain('script');
  });
});
