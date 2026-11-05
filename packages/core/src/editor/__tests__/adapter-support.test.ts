import { describe, expect, spyOn, test } from 'bun:test';
import { createTranslator } from '@pptx-editor.dev/i18n';
import {
  EMPTY_EDITOR_SNAPSHOT,
  RIBBON_TABS,
  TOOLBAR_BAR_GROUPS,
  TOOLBAR_COLLAPSE_ORDER,
  TOOLBAR_NAMED_PARTS,
  TOOLBAR_PICTURE_GROUPS,
  arrangeToolbarGroups,
  builtInGroupPriority,
  chromeSlotIds,
  ribbonTabGroups,
  toolbarCollapse,
  toolbarGroupPriority,
  chromeKeepsFocus,
  chromeCommandState,
  isToolbarBarSlot,
  platformShortcuts,
  resetDevWarnings,
  sameChromeCommandState,
  toolbarFitCount,
  toolbarGroupLabel,
  translateKeyOrText,
  type ToolbarGroupOverride,
} from '../index.ts';

describe('adapter support', () => {
  test('the empty snapshot is frozen and editable', () => {
    expect(Object.isFrozen(EMPTY_EDITOR_SNAPSHOT)).toBe(true);
    expect(EMPTY_EDITOR_SNAPSHOT.status).toBe('empty');
    expect(EMPTY_EDITOR_SNAPSHOT.mode).toBe('edit');
  });

  test('command state without an editor is disabled with the engine reason', () => {
    const state = chromeCommandState(null, 'history.undo');
    expect(state.isEnabled).toBe(false);
    expect(state.disabledReason).toBe('disabled.noEditor');
    expect(sameChromeCommandState(state, chromeCommandState(null, 'history.undo'))).toBe(true);
    expect(sameChromeCommandState(state, null)).toBe(false);
  });

  test('shortcut text follows the platform', () => {
    expect(platformShortcuts('Bold (Ctrl+B)', false)).toBe('Bold (Ctrl+B)');
    expect(platformShortcuts('Group (Ctrl+Shift+G)', true)).toBe('Group (⇧⌘G)');
    expect(platformShortcuts('Alt+X', true)).toBe('⌥X');
    expect(platformShortcuts('Ctrl+M', true)).toBe('⌃M');
    expect(platformShortcuts('Ctrl+M', false)).toBe('Ctrl+M');
    expect(platformShortcuts('Replace (Ctrl+H)', true)).toBe('Replace (⇧⌘H)');
    expect(platformShortcuts('Ctrl+Home', true)).toBe('⌘Home');
    // Redo reads ⇧⌘Z on macOS; Ctrl+Y elsewhere.
    expect(platformShortcuts('Redo (Ctrl+Y)', true)).toBe('Redo (⇧⌘Z)');
    expect(platformShortcuts('Redo (Ctrl+Y)', false)).toBe('Redo (Ctrl+Y)');
    expect(platformShortcuts('F5', true)).toBe('⇧⌘↩');
    expect(platformShortcuts('Shift+F5', true)).toBe('⌘↩');
    expect(platformShortcuts('F5', false)).toBe('F5');
  });

  test('translation fallback and the focus rule', () => {
    const t = createTranslator();
    expect(translateKeyOrText(t, 'toolbar.more')).toBe(t('toolbar.more'));
    expect(translateKeyOrText(t, 'Not a key')).toBe('Not a key');
    expect(chromeKeepsFocus({ tagName: 'BUTTON' } as unknown as EventTarget)).toBe(true);
    expect(chromeKeepsFocus({ tagName: 'INPUT' } as unknown as EventTarget)).toBe(false);
  });

  test('the arrangement hides, keeps, and appends groups in order', () => {
    const overrides = new Map<string, ToolbarGroupOverride>([
      ['insert', { id: 'insert', hidden: true }],
      ['text', { id: 'text', label: 'ignored for built-in' }],
      ['acme', { id: 'acme', label: 'Acme' }],
      ['gone', { id: 'gone', hidden: true }],
    ]);
    const arranged = arrangeToolbarGroups(overrides);
    const ids = arranged.map((g) => (g.kind === 'built-in' ? g.group.id : g.override.id));
    expect(ids).toEqual([
      ...TOOLBAR_BAR_GROUPS.map((g) => g.id).filter((id) => id !== 'insert'),
      'acme',
    ]);
    const text = arranged.find((g) => g.kind === 'built-in' && g.group.id === 'text');
    expect(text?.kind === 'built-in' && text.override?.id).toBe('text');
    expect(isToolbarBarSlot('text.bold')).toBe(true);
    expect(isToolbarBarSlot('view.smartGuides')).toBe(false);
  });

  test('host group labels: label, then labelKey, then id', () => {
    const t = (key: string) => `t:${key}`;
    expect(toolbarGroupLabel({ id: 'a', label: 'A', labelKey: 'k' }, t)).toBe('A');
    expect(toolbarGroupLabel({ id: 'a', labelKey: 'k' }, t)).toBe('t:k');
    expect(toolbarGroupLabel({ id: 'a' }, t)).toBe('a');
  });

  test('fit count keeps all groups that fit, else room for More', () => {
    expect(toolbarFitCount([100, 100, 100], 300)).toBe(3);
    expect(toolbarFitCount([100, 100, 100], 299)).toBe(2);
    expect(toolbarFitCount([100, 100, 100], 239)).toBe(1);
    expect(toolbarFitCount([100], 50)).toBe(0);
    expect(toolbarFitCount([], 0)).toBe(0);
  });
});

describe('toolbar layout', () => {
  test('every registered slot has exactly one named part', () => {
    const named = Object.values(TOOLBAR_NAMED_PARTS);
    expect(new Set(named).size).toBe(named.length);
    expect([...named].sort()).toEqual([...chromeSlotIds()].sort());
  });

  test('built-in priorities follow the collapse order: formatting stays, clipboard and Find go', () => {
    const ids = TOOLBAR_BAR_GROUPS.map((g) => g.id);
    // Every bar group has a declared place.
    for (const id of ids) expect(TOOLBAR_COLLAPSE_ORDER.includes(id)).toBe(true);
    expect(builtInGroupPriority('find', TOOLBAR_BAR_GROUPS)).toBe(
      (TOOLBAR_COLLAPSE_ORDER.indexOf('find') + 1) * 10
    );
    expect(builtInGroupPriority('history', TOOLBAR_BAR_GROUPS)).toBe(
      TOOLBAR_COLLAPSE_ORDER.length * 10
    );
    const rank = (id: string) => builtInGroupPriority(id, TOOLBAR_BAR_GROUPS);
    // The font and paragraph controls outrank the clipboard, Find, and Insert.
    for (const low of ['edit', 'find', 'insert', 'more', 'slide'])
      for (const high of ['text', 'paragraph', 'history'])
        expect(rank(low) < rank(high)).toBe(true);
    expect(toolbarGroupPriority('acme', undefined, TOOLBAR_BAR_GROUPS)).toBe(0);
    expect(toolbarGroupPriority('history', 5, TOOLBAR_BAR_GROUPS)).toBe(5);
  });

  test('on the ribbon Home tab, Font and Paragraph stay longest after a narrow fit', () => {
    const home = RIBBON_TABS.find((t) => t.id === 'home')!;
    const groups = ribbonTabGroups(home);
    const items = groups.map((g) => ({ width: 100, priority: builtInGroupPriority(g.id, groups) }));
    // Room for five groups and the More button.
    const hidden = toolbarCollapse(items, 545, 45);
    const shown = groups.filter((_, i) => !hidden.has(i)).map((g) => g.id);
    expect(shown).toEqual(['font', 'text', 'list', 'paragraph', 'insert']);
  });

  test('the contextual image group does not change the other priorities', () => {
    // A host priority set between two built-in groups keeps its place when the group joins.
    for (const g of TOOLBAR_BAR_GROUPS)
      expect(builtInGroupPriority(g.id, TOOLBAR_PICTURE_GROUPS)).toBe(
        builtInGroupPriority(g.id, TOOLBAR_BAR_GROUPS)
      );
    // The image group is the selection's own: it stays longest, above every bar group.
    expect(builtInGroupPriority('image', TOOLBAR_PICTURE_GROUPS)).toBe(
      (TOOLBAR_COLLAPSE_ORDER.length + 1) * 10
    );
  });

  test('groups leave the bar lowest priority first, and pinned groups stay', () => {
    const items = [
      { width: 100, priority: 30 },
      { width: 100, priority: 20, pinned: true },
      { width: 100, priority: 10 },
      { width: 100, priority: 0 },
    ];
    expect([...toolbarCollapse(items, 400)]).toEqual([]);
    expect([...toolbarCollapse(items, 399, 40)]).toEqual([3]);
    expect([...toolbarCollapse(items, 299, 40)]).toEqual([3, 2]);
    expect([...toolbarCollapse(items, 100, 40)].sort()).toEqual([0, 2, 3]);
  });

  test('a host group follows the group its after names', () => {
    const overrides = new Map([
      ['acme', { id: 'acme', after: 'history' }],
      ['beta', { id: 'beta', after: 'acme' }],
      ['tail', { id: 'tail' }],
    ]);
    const ids = arrangeToolbarGroups(overrides).map((e) =>
      e.kind === 'host' ? e.override.id : e.group.id
    );
    expect(ids.slice(0, 4)).toEqual(['history', 'acme', 'beta', 'font']);
    expect(ids[ids.length - 1]).toBe('tail');
  });

  test('placement keeps order at one anchor, waits for a later anchor, and uses hidden groups', () => {
    const order = (overrides: [string, { id: string; after?: string; hidden?: boolean }][]) =>
      arrangeToolbarGroups(new Map(overrides)).map((e) =>
        e.kind === 'host' ? e.override.id : e.group.id
      );
    // Two groups at one anchor keep their order of appearance.
    expect(
      order([
        ['a', { id: 'a', after: 'history' }],
        ['b', { id: 'b', after: 'history' }],
      ]).slice(0, 3)
    ).toEqual(['history', 'a', 'b']);
    // An anchor declared later still places the group after it.
    const forward = order([
      ['b', { id: 'b', after: 'a' }],
      ['a', { id: 'a', after: 'history' }],
    ]);
    expect(forward.slice(0, 3)).toEqual(['history', 'a', 'b']);
    // A hidden built-in group still anchors: the host group takes its place.
    const hidden = order([
      ['history', { id: 'history', hidden: true }],
      ['acme', { id: 'acme', after: 'history' }],
    ]);
    expect(hidden[0]).toBe('acme');
    expect(hidden).not.toContain('history');
  });

  test('the contextual table group is a known anchor while it is not on the bar', () => {
    resetDevWarnings();
    const warn = spyOn(console, 'warn').mockImplementation(() => {});
    try {
      arrangeToolbarGroups(new Map([['acme', { id: 'acme', after: 'table' }]]));
      expect(warn).not.toHaveBeenCalled();
    } finally {
      warn.mockRestore();
    }
  });

  test('an unknown anchor and host-only settings on a built-in group warn', () => {
    resetDevWarnings();
    const warn = spyOn(console, 'warn').mockImplementation(() => {});
    try {
      arrangeToolbarGroups(
        new Map([
          ['acme', { id: 'acme', after: 'nope' }],
          ['find', { id: 'find', after: 'history', label: 'Search' }],
        ])
      );
      const text = warn.mock.calls.map((c: unknown[]) => String(c[0])).join('\n');
      expect(text).toContain('Toolbar.Group "acme" has after="nope", but no group has that id');
      expect(text).toContain('Toolbar.Group "find" is a built-in group, so its after is ignored');
      expect(text).toContain('Toolbar.Group "find" is a built-in group, so its label is ignored');
    } finally {
      warn.mockRestore();
    }
  });

  test('a group that fits in the room a wider group left comes back', () => {
    const items = [
      { width: 100, priority: 90 },
      { width: 300, priority: 20 },
      { width: 100, priority: 10 },
    ];
    // Greedy: the 100 px group leaves, then the 300 px one; the 100 px group then fits again.
    expect([...toolbarCollapse(items, 300, 40)]).toEqual([1]);
  });

  test('equal priorities give up the later group first; Infinity keeps a group longest', () => {
    const items = [
      { width: 100, priority: 0 },
      { width: 100, priority: 0 },
      { width: 100, priority: 0 },
    ];
    expect([...toolbarCollapse(items, 250, 0)]).toEqual([2]);
    expect(toolbarGroupPriority('acme', Infinity, [])).toBe(Infinity);
    expect(toolbarGroupPriority('acme', Number.NaN, [])).toBe(0);
  });
});
