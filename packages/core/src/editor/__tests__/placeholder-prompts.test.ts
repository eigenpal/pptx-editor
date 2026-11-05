// Placeholder prompt text (`parity-baseline` 4.8): the editor surface paints the prompt of an
// empty placeholder in the style typed text takes; thumbnails, the harness pipeline, and the
// saved file never hold it.

import { describe, expect, spyOn, test } from 'bun:test';
import { buildPptx } from '../../../../../scripts/fixtures/builder.ts';
import { buildSlideLayoutsDeck } from '../../../../../scripts/fixtures/slide-layouts.ts';
import type { LayoutItem, SlideLayout } from '../../layout/index.ts';
import { readPptx, xmlPartOf } from '../../store/index.ts';
import { createPptxEditor, type PptxEditorInstance } from '../index.ts';
import { PlaceholderPrompts } from '../placeholder-prompts.ts';
import { resetDevWarnings } from '../dev-warnings.ts';

async function open(): Promise<PptxEditorInstance> {
  const e = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await e.load(buildSlideLayoutsDeck())).ok).toBe(true);
  return e;
}

/** Each shape's painted text, by shape name. */
function painted(layout: SlideLayout | undefined): Record<string, string> {
  const out: Record<string, string> = {};
  for (const item of layout?.items ?? [])
    if (item.kind === 'shape' && item.element.layer === 'slide')
      out[item.element.name] = (item.text?.lines ?? [])
        .flatMap((l) => l.runs.filter((r) => r.kind === 'text').map((r) => r.text))
        .join('|');
  return out;
}

const shapeNamed = (layout: SlideLayout | undefined, name: string): LayoutItem =>
  layout!.items.find((i) => i.element.name === name)!;

describe('placeholder prompts', () => {
  test('an empty placeholder shows the prompt for its type', async () => {
    const e = await open();
    const prompts = new PlaceholderPrompts();
    const layout = e.session!.layout(4);
    expect(painted(prompts.paint(e.session, layout, { kind: 'none' }, false))).toEqual({
      'Title 1': 'Click to add title',
      'Content Placeholder 2': 'Click to add text',
    });
    // The slide's own layout (rail, slide show, harness) holds no prompt.
    expect(painted(layout)).toEqual({ 'Title 1': '', 'Content Placeholder 2': '' });
  });

  test('a layout preview shows the prompts of a new slide; a non-layout id shows nothing', async () => {
    const e = await open();
    const prompts = new PlaceholderPrompts();
    const layouts = e.snapshot().layouts!;
    const twoContent = layouts.find((l) => l.name === 'Two Content')!;
    expect(painted(prompts.preview(e.session!, twoContent.id))).toEqual({
      'Title 1': 'Click to add title',
      'Content Placeholder 2': 'Click to add text',
      'Content Placeholder 3': 'Click to add text',
    });
    expect(layouts.every((l) => l.masterName === 'Fixture')).toBe(true);
    // A slide part is not a layout: no preview, and a development warning.
    const warn = spyOn(console, 'warn').mockImplementation(() => {});
    resetDevWarnings();
    expect(prompts.preview(e.session!, e.session!.info.slides[0].partName)).toBeUndefined();
    expect(warn.mock.calls[0]?.[0]).toContain('names no slide layout');
    warn.mockRestore();
    // The package is untouched.
    expect(e.snapshot().slides.length).toBe(6);
  });

  test('a placeholder type that is an object key of the prompt table reads as no prompt', async () => {
    const sp = (type: string, id: number) =>
      `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="P ${id}"/><p:cNvSpPr/><p:nvPr><p:ph type="${type}" idx="${id}"/></p:nvPr></p:nvSpPr><p:spPr/>` +
      '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>';
    const e = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    const shapes = ['constructor', 'toString', '__proto__', 'hasOwnProperty'].map((t, i) =>
      sp(t, i + 2)
    );
    expect((await e.load(buildPptx({ slides: [{ shapes: shapes.join('') }] }))).ok).toBe(true);
    const layout = e.session!.layout(0);
    expect(() =>
      new PlaceholderPrompts().paint(e.session, layout, { kind: 'none' }, false)
    ).not.toThrow();
  });

  test('a custom prompt comes from the layout placeholder', async () => {
    const e = await open();
    const out = new PlaceholderPrompts().paint(
      e.session,
      e.session!.layout(5),
      { kind: 'none' },
      false
    );
    expect(painted(out)).toEqual({
      'Title 1': 'In their words',
      'Quote Placeholder 2': 'Type a quote here',
    });
  });

  test('the host translator names the prompts; an unknown key falls back to English', async () => {
    const e = await open();
    const french = (key: string): string => (key === 'canvas.prompts.title' ? 'Titre' : key);
    const out = new PlaceholderPrompts(french).paint(
      e.session,
      e.session!.layout(4),
      { kind: 'none' },
      false
    );
    expect(painted(out)['Title 1']).toBe('Titre');
    expect(painted(out)['Content Placeholder 2']).toBe('Click to add text');
  });

  test('the placeholder being edited, and viewing mode, show no prompt', async () => {
    const e = await open();
    const layout = e.session!.layout(4);
    const title = shapeNamed(layout, 'Title 1').element.nodeId;
    const caret = { paragraph: 0, offset: 0 };
    const prompts = new PlaceholderPrompts();
    const editing = prompts.paint(
      e.session,
      layout,
      { kind: 'text', nodeId: title, anchor: caret, focus: caret },
      false
    );
    expect(painted(editing)['Title 1']).toBe('');
    expect(painted(editing)['Content Placeholder 2']).toBe('Click to add text');
    expect(prompts.paint(e.session, layout, { kind: 'none' }, true)).toBe(layout);
  });

  test('typed text takes the style the prompt shows, and nothing is saved', async () => {
    const e = await open();
    e.goToSlide(4);
    const layout = e.session!.layout(4);
    const prompt = new PlaceholderPrompts().paint(e.session, layout, { kind: 'none' }, false);
    const promptRun = shapeNamed(prompt, 'Title 1');
    const title = promptRun.element.nodeId;
    const caret = { paragraph: 0, offset: 0 };
    e.setSelection({ kind: 'text', nodeId: title, anchor: caret, focus: caret });
    expect(e.input('text', 'Results').ok).toBe(true);
    const typed = shapeNamed(e.session!.layout(4), 'Title 1');
    const runOf = (item: LayoutItem) =>
      item.kind === 'shape' ? item.text!.lines[0].runs.find((r) => r.kind === 'text')! : null;
    expect(runOf(typed)!.sizePt).toBe(runOf(promptRun)!.sizePt);
    expect(runOf(typed)!.face).toEqual(runOf(promptRun)!.face);
    expect(runOf(typed)!.props.fill).toEqual(runOf(promptRun)!.props.fill);
    // The saved file holds the typed text and no prompt.
    const r = readPptx(e.save()!);
    if (!r.ok) throw new Error('reopen');
    const xml = (part: string) => JSON.stringify(xmlPartOf(r.pkg, part)?.root ?? null);
    const slide = e.session!.info.slides[4].partName;
    expect(xml(slide)).toContain('Results');
    for (const name of r.pkg.parts.keys()) expect(xml(name)).not.toContain('Click to add');
  });
});

describe('placeholder prompts and the locale', () => {
  test('a locale change repaints the prompts, with the same translator', async () => {
    const e = await open();
    let locale = 'en';
    const prompts = new PlaceholderPrompts((key) =>
      locale === 'de' && key === 'canvas.prompts.title' ? 'Titel hinzufügen' : key
    );
    const layout = e.session!.layout(4);
    expect(painted(prompts.paint(e.session, layout, { kind: 'none' }, false))['Title 1']).toBe(
      'Click to add title'
    );
    locale = 'de';
    expect(painted(prompts.paint(e.session, layout, { kind: 'none' }, false))['Title 1']).toBe(
      'Titel hinzufügen'
    );
  });
});
