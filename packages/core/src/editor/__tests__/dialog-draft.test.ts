import { describe, expect, test } from 'bun:test';
import {
  createDialogDraft,
  dialogPartKey,
  dialogPartLabel,
  returnAppliesFrom,
} from '../dialog-draft.ts';

// The draft controller every packaged dialog uses, in React and Vue alike.

interface Fields {
  readonly width: number;
  readonly name: string;
}

describe('createDialogDraft', () => {
  test('edits stay in the draft until Apply commits them once', () => {
    const commits: Fields[] = [];
    const draft = createDialogDraft<Fields>({
      initial: { width: 10, name: 'a' },
      commit: (values) => (commits.push(values), { ok: true }),
    });
    draft.setValue('width', 20);
    draft.setValue('name', 'b');
    expect(commits).toEqual([]);
    expect(draft.apply()).toBe(true);
    expect(commits).toEqual([{ width: 20, name: 'b' }]);
  });

  test('validation errors block the commit and clear as the field changes', () => {
    let committed = false;
    const draft = createDialogDraft<Fields>({
      initial: { width: 0, name: 'a' },
      validate: (v) => (v.width > 0 ? {} : { width: 'dialog.errors.positive' }),
      commit: () => ((committed = true), { ok: true }),
    });
    expect(draft.apply()).toBe(false);
    expect(committed).toBe(false);
    expect(draft.errors).toEqual({ width: 'dialog.errors.positive' });
    draft.setValue('width', 5);
    expect(draft.errors).toEqual({});
    expect(draft.apply()).toBe(true);
  });

  test('a refused commit keeps the draft and reports the reason as the form error', () => {
    const draft = createDialogDraft<Fields>({
      initial: { width: 1, name: 'a' },
      commit: () => ({ ok: false, reason: 'disabled.viewing' }),
    });
    draft.setValue('name', 'b');
    expect(draft.apply()).toBe(false);
    expect(draft.errors).toEqual({ form: 'disabled.viewing' });
    expect(draft.values.name).toBe('b');
  });

  test('reset restores the reset values and notifies listeners', () => {
    const draft = createDialogDraft<Fields>({
      initial: { width: 1, name: 'a' },
      resetValues: () => ({ width: 9, name: 'z' }),
      commit: () => ({ ok: true }),
    });
    let calls = 0;
    const off = draft.subscribe(() => (calls += 1));
    draft.setValue('width', 2);
    draft.setValue('width', 2);
    draft.reset();
    off();
    draft.setValue('width', 3);
    expect(calls).toBe(2);
    expect(draft.values.width).toBe(3);
  });

  test('part keys name fields by their field name', () => {
    expect(dialogPartKey('apply')).toBe('apply');
    expect(dialogPartKey('field', 'theme')).toBe('field:theme');
    expect(dialogPartKey(undefined)).toBeNull();
  });
});

describe('dialog frame rules', () => {
  test('Return applies from a field or a radio, not from a button-type input or a button', () => {
    const el = (localName: string, type?: string, role?: string) =>
      ({
        localName,
        type,
        getAttribute: (n: string) => (n === 'role' ? (role ?? null) : null),
      }) as never;
    expect(returnAppliesFrom(el('input', 'text'))).toBe(true);
    expect(returnAppliesFrom(el('input', 'checkbox'))).toBe(true);
    expect(returnAppliesFrom(el('button', undefined, 'radio'))).toBe(true);
    for (const type of ['button', 'submit', 'reset', 'image', 'file'])
      expect(returnAppliesFrom(el('input', type))).toBe(false);
    expect(returnAppliesFrom(el('button'))).toBe(false);
    expect(returnAppliesFrom(el('textarea'))).toBe(false);
  });

  test('part labels name the part as a host writes it, and a field without a name', () => {
    expect(dialogPartLabel('PageSetupDialog', 'apply')).toBe('PageSetupDialog.Apply');
    expect(dialogPartLabel('PageSetupDialog', 'field:theme')).toBe(
      'PageSetupDialog.Field name="theme"'
    );
    expect(dialogPartLabel('PageSetupDialog', 'field:undefined')).toBe(
      'PageSetupDialog.Field (without a name)'
    );
  });
});
