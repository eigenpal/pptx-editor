import { describe, expect, test } from 'bun:test';
import { createTranslator } from '@pptx-editor.dev/i18n';
import { buildPptx } from '../../../../../scripts/fixtures/builder.ts';
import { canonicalPackageFingerprint, readPptx } from '../../store/index.ts';
import { createPptxEditor } from '../pptx-editor.ts';
import {
  backgroundDraftCommand,
  backgroundDraftFrom,
  commitDialogCommand,
  radioGroupStep,
  sameBackground,
  translateDialogErrors,
  validateBackgroundDraft,
  type BackgroundDraftFields,
} from '../dialog-fields.ts';
import { createDialogDraft } from '../dialog-draft.ts';

// The field rules of the packaged dialogs: what the Background dialog applies, and when it
// refuses. The draft opens from the snapshot, which keeps only RGB and hex gradient stops, so
// the rules never rewrite a slide from that lossy value.

const navy = { kind: 'solid', color: '#000080' } as const;
const opening = (background: BackgroundDraftFields['background']): BackgroundDraftFields => ({
  background,
  allSlides: false,
});

describe('Background draft rules', () => {
  test('the draft opens from the slide: inherited, solid, gradient, or a fill it cannot hand back', () => {
    expect(backgroundDraftFrom({ kind: 'solid', color: '#FF0000', inherited: true })).toEqual(
      opening(null)
    );
    expect(backgroundDraftFrom({ kind: 'solid', color: '#FF0000', inherited: false })).toEqual(
      opening({ kind: 'solid', color: '#FF0000' })
    );
    const gradient = { kind: 'linear', angle: 90, stops: [] } as never;
    expect(backgroundDraftFrom({ kind: 'gradient', gradient, inherited: false })).toEqual(
      opening({ kind: 'gradient', gradient })
    );
    expect(backgroundDraftFrom({ kind: 'picture', inherited: false })).toEqual(opening(undefined));
  });

  test('an unchanged draft is no command; a pick of the same fill is one', () => {
    const initial = opening({ kind: 'solid', color: '#FF0000' });
    expect(backgroundDraftCommand(initial, initial)).toBeNull();
    expect(backgroundDraftCommand(opening(undefined), opening(undefined))).toBeNull();
    // A pick is an edit: another editor may have changed the slide since the dialog opened.
    expect(backgroundDraftCommand(opening({ kind: 'solid', color: '#FF0000' }), initial)).toEqual({
      type: 'setSlideBackground',
      background: { kind: 'solid', color: '#FF0000' },
    });
  });

  test('a picked fill is one command, for every slide with Apply to all slides', () => {
    expect(backgroundDraftCommand(opening(navy), opening(null))).toEqual({
      type: 'setSlideBackground',
      background: navy,
    });
    expect(backgroundDraftCommand({ background: navy, allSlides: true }, opening(null))).toEqual({
      type: 'setSlideBackground',
      background: navy,
      allSlides: true,
    });
  });

  test('Apply to all slides needs a picked fill or the layout background', () => {
    const solid = opening({ kind: 'solid', color: '#FF0000' });
    expect(validateBackgroundDraft({ ...solid, allSlides: true }, solid)).toEqual({
      allSlides: 'background.pickFirst',
    });
    expect(
      validateBackgroundDraft({ background: undefined, allSlides: true }, opening(undefined))
    ).toEqual({ allSlides: 'background.pickFirst' });
    // Picking the slide's own fill again is a pick: it can go to every slide.
    expect(
      validateBackgroundDraft(
        { background: { kind: 'solid', color: '#FF0000' }, allSlides: true },
        solid
      )
    ).toEqual({});
    expect(validateBackgroundDraft({ background: navy, allSlides: true }, solid)).toEqual({});
    // The layout's background removes each slide's own fill: nothing is rebuilt from a snapshot.
    expect(validateBackgroundDraft({ background: null, allSlides: true }, opening(null))).toEqual(
      {}
    );
    expect(backgroundDraftCommand({ background: null, allSlides: true }, opening(null))).toEqual({
      type: 'setSlideBackground',
      background: null,
      allSlides: true,
    });
  });

  test('Apply to all slides with the opening theme fill leaves every slide as it was', async () => {
    const themed =
      '<p:bg><p:bgPr><a:solidFill><a:schemeClr val="accent1"><a:lumMod val="75000"/><a:alpha val="60000"/></a:schemeClr></a:solidFill><a:effectLst/></p:bgPr></p:bg>';
    const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    const bytes = buildPptx({
      slides: [
        { layout: 'blank', background: themed, shapes: '' },
        { layout: 'blank', shapes: '' },
      ],
    });
    expect((await editor.load(bytes)).ok).toBe(true);
    const before = editor.save()!;
    const initial = backgroundDraftFrom(editor.snapshot().slideBackground);
    const draft = createDialogDraft({
      initial,
      validate: (values) => validateBackgroundDraft(values, initial),
      commit: (values) => commitDialogCommand(editor, backgroundDraftCommand(values, initial)),
    });
    draft.setValue('allSlides', true);
    expect(draft.apply()).toBe(false);
    expect(draft.errors).toEqual({ allSlides: 'background.pickFirst' });
    expect(editor.snapshot().canUndo).toBe(false);
    const fingerprint = (bytes: Uint8Array) => {
      const pkg = readPptx(bytes);
      if (!pkg.ok) throw new Error('unreadable');
      return canonicalPackageFingerprint(pkg.pkg);
    };
    expect(fingerprint(editor.save()!)).toEqual(fingerprint(before));
  });

  test('an explicit pick applies even when it equals the opening value (Reset on an inherited slide)', () => {
    expect(backgroundDraftCommand(opening(null), opening(null))).toBeNull();
    expect(backgroundDraftCommand(opening(null), opening(null), true)).toEqual({
      type: 'setSlideBackground',
      background: null,
    });
    const solid = opening({ kind: 'solid', color: '#FF0000' });
    expect(validateBackgroundDraft({ ...solid, allSlides: true }, solid, true)).toEqual({});
  });

  test('an error the new values no longer have clears on edit', () => {
    const initial = opening({ kind: 'solid', color: '#FF0000' });
    const draft = createDialogDraft<BackgroundDraftFields>({
      initial,
      validate: (values) => validateBackgroundDraft(values, initial),
      commit: () => ({ ok: true }),
    });
    draft.setValue('allSlides', true);
    expect(draft.apply()).toBe(false);
    expect(draft.errors.allSlides).toBe('background.pickFirst');
    draft.setValue('background', navy);
    expect(draft.errors).toEqual({});
  });

  test('same fill compares by value; pictures compare by identity', () => {
    expect(sameBackground(navy, { kind: 'solid', color: '#000080' })).toBe(true);
    expect(sameBackground(navy, { kind: 'solid', color: '#000081' })).toBe(false);
    expect(sameBackground(null, undefined)).toBe(false);
    const picture = { kind: 'picture', bytes: new Uint8Array([1]), mime: 'image/png' } as const;
    expect(sameBackground(picture, { ...picture })).toBe(false);
    expect(sameBackground(picture, picture)).toBe(true);
  });
});

describe('dialog commits and messages', () => {
  test('a commit runs the command; no editor, a refusal, and a no-op are reported', () => {
    expect(commitDialogCommand(null, null)).toEqual({ ok: true });
    expect(commitDialogCommand(null, { type: 'undo' })).toEqual({
      ok: false,
      reason: 'disabled.noEditor',
    });
    const refuse = { exec: () => ({ ok: false, reason: 'disabled.viewing' }) } as never;
    expect(commitDialogCommand(refuse, { type: 'undo' })).toEqual({
      ok: false,
      reason: 'disabled.viewing',
    });
    const noop = { exec: () => ({ ok: false, noop: true, reason: 'x' }) } as never;
    expect(commitDialogCommand(noop, { type: 'undo' })).toEqual({ ok: true });
  });

  test('errors reach parts and hooks translated; engine text stays as it is', () => {
    const t = createTranslator();
    expect(
      translateDialogErrors<BackgroundDraftFields>(t, {
        background: 'background.pickFirst',
        form: 'The file is locked',
      })
    ).toEqual({
      background:
        "To apply a fill to all slides, pick a color, gradient, pattern, or picture. The slide's current fill cannot be copied.",
      form: 'The file is locked',
    });
  });

  test('arrow keys move through a radio group and wrap', () => {
    const themes = ['light', 'dark'] as const;
    expect(radioGroupStep(themes, 'light', 'ArrowRight')).toBe('dark');
    expect(radioGroupStep(themes, 'dark', 'ArrowDown')).toBe('light');
    expect(radioGroupStep(themes, 'light', 'ArrowLeft')).toBe('dark');
    expect(radioGroupStep(themes, 'dark', 'ArrowUp')).toBe('light');
    expect(radioGroupStep(themes, 'dark', 'Enter')).toBeNull();
  });

  test('a validator key with no message does not block Apply', () => {
    let committed = 0;
    const draft = createDialogDraft({
      initial: { w: 1 },
      validate: () => ({ w: undefined }),
      commit: () => {
        committed += 1;
        return { ok: true };
      },
    });
    expect(draft.apply()).toBe(true);
    expect(committed).toBe(1);
  });
});
