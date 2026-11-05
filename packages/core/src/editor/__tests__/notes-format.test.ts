// Speaker notes keep and edit run formatting (`parity-baseline` 4.7): the snapshot reports the
// formatted runs, `formatNotes` sets bold, italic, underline, and strike on a range, and a text
// edit keeps the formatting around it. Each edit is checked after save and reopen.

import { describe, expect, test } from 'bun:test';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { buildPptx } from '../../../../../scripts/fixtures/builder.ts';
import { buildSlideLayoutsDeck } from '../../../../../scripts/fixtures/slide-layouts.ts';
import { canonicalPackageFingerprint, readPptx } from '../../store/index.ts';
import { createPptxEditor, type PptxEditorInstance } from '../index.ts';
import type { NotesParagraph } from '../../contracts/index.ts';
import { notesParagraphs } from '../../automation/commands/notes-commands.ts';
import { diffHunks } from '../../automation/commands/text-diff.ts';

async function open(): Promise<PptxEditorInstance> {
  const e = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await e.load(buildSlideLayoutsDeck())).ok).toBe(true);
  e.goToSlide(1);
  return e;
}

/** The notes as `text[flags]` runs: b, i, u, s for bold, italic, underline, strike. */
function shown(paragraphs: readonly NotesParagraph[]): string[][] {
  return paragraphs.map((p) =>
    p.runs.map(
      (r) =>
        `${r.text}[${r.bold ? 'b' : ''}${r.italic ? 'i' : ''}${r.underline ? 'u' : ''}${r.strike ? 's' : ''}]`
    )
  );
}

function reopened(e: PptxEditorInstance) {
  const r = readPptx(e.save()!);
  if (!r.ok) throw new Error('reopen');
  return r.pkg;
}

const FIRST = 'Open with the headline number and pause.';

describe('notes formatting in the snapshot', () => {
  test('reports each run with its formatting; the text equals the plain notes', async () => {
    const e = await open();
    const snap = e.snapshot();
    expect(shown(snap.notesFormatted!)).toEqual([
      ['Open with the []', 'headline number[b]', ' and []', 'pause[i]', '.[]'],
      ['Then show the chart.[u]'],
    ]);
    expect(snap.notesFormatted!.map((p) => p.runs.map((r) => r.text).join('')).join('\n')).toBe(
      snap.notes
    );
    // Reference-stable while the notes do not change.
    e.exec({ type: 'setSlidesHidden', indices: [3], hidden: false });
    expect(e.snapshot().notesFormatted).toBe(snap.notesFormatted);
  });

  test('reads the notes master body placeholder list style too', async () => {
    const files = unzipSync(buildSlideLayoutsDeck());
    const name = 'ppt/notesMasters/notesMaster1.xml';
    files[name] = strToU8(
      strFromU8(files[name]!).replace(
        '<a:lstStyle/>',
        '<a:lstStyle><a:lvl1pPr><a:defRPr b="1"/></a:lvl1pPr></a:lstStyle>'
      )
    );
    const e = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    expect((await e.load(zipSync(files))).ok).toBe(true);
    e.goToSlide(1);
    // Every run without its own `b` is bold now; the toggle then turns bold off.
    expect(shown(e.snapshot().notesFormatted!)[1]).toEqual(['Then show the chart.[bu]']);
    e.exec({
      type: 'formatNotes',
      index: 1,
      start: FIRST.length + 1,
      end: FIRST.length + 5,
      format: 'bold',
    });
    expect(shown(e.snapshot().notesFormatted!)[1]).toEqual(['Then[u]', ' show the chart.[bu]']);
  });

  test('a slide without notes has none', async () => {
    const e = await open();
    e.goToSlide(0);
    expect(e.snapshot().notesFormatted).toEqual([]);
  });
});

describe('formatNotes', () => {
  test('bold on part of a run splits it and keeps the other formatting', async () => {
    const e = await open();
    const before = canonicalPackageFingerprint(e.session!.store.package);
    const start = FIRST.length + 1; // "Then"
    expect(
      e.exec({ type: 'formatNotes', index: 1, start, end: start + 4, format: 'bold' })
    ).toEqual({ ok: true });
    const expected = [
      ['Open with the []', 'headline number[b]', ' and []', 'pause[i]', '.[]'],
      ['Then[bu]', ' show the chart.[u]'],
    ];
    expect(shown(e.snapshot().notesFormatted!)).toEqual(expected);
    expect(shown(notesParagraphs(reopened(e), 1))).toEqual(expected);
    e.exec({ type: 'undo' });
    expect(canonicalPackageFingerprint(e.session!.store.package)).toBe(before);
  });

  test('toggles off when the whole range has the format, on when part of it lacks it', async () => {
    const e = await open();
    const at = FIRST.indexOf('headline');
    e.exec({ type: 'formatNotes', index: 1, start: at, end: at + 15, format: 'bold' });
    expect(shown(e.snapshot().notesFormatted!)[0]).toEqual([
      'Open with the headline number and []',
      'pause[i]',
      '.[]',
    ]);
    // Across a plain and an italic run: italic goes on for all of it.
    const from = FIRST.indexOf('and');
    e.exec({ type: 'formatNotes', index: 1, start: from, end: FIRST.length - 1, format: 'italic' });
    expect(shown(e.snapshot().notesFormatted!)[0]).toEqual([
      'Open with the headline number []',
      'and pause[i]',
      '.[]',
    ]);
    // An explicit value sets it, whatever the range has.
    e.exec({ type: 'formatNotes', index: 1, start: 0, end: 4, format: 'strike', value: true });
    e.exec({ type: 'formatNotes', index: 1, start: 0, end: 4, format: 'strike', value: true });
    expect(shown(e.snapshot().notesFormatted!)[0]![0]).toBe('Open[s]');
  });

  test('a text edit keeps the formatting of the runs it does not touch', async () => {
    const e = await open();
    const text = e.snapshot().notes.replace('pause', 'stop').replace('chart', 'table');
    e.exec({ type: 'setNotes', index: 1, text });
    expect(shown(notesParagraphs(reopened(e), 1))).toEqual([
      ['Open with the []', 'headline number[b]', ' and []', 'stop[i]', '.[]'],
      ['Then show the table.[u]'],
    ]);
  });

  test('a field in the range takes the format whole, so the toggle turns it off again', async () => {
    const field =
      '<a:p><a:r><a:rPr lang="en-US"/><a:t>Due </a:t></a:r>' +
      '<a:fld id="{B2C3D4E5-0000-4000-8000-000000000001}" type="datetime1"><a:rPr lang="en-US"/><a:t>10/31/2026</a:t></a:fld>' +
      '<a:r><a:rPr lang="en-US"/><a:t> ok</a:t></a:r></a:p>';
    const e = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    expect((await e.load(buildPptx({ slides: [{ shapes: '', notesXml: field }] }))).ok).toBe(true);
    const all = { type: 'formatNotes', index: 0, start: 0, end: 17, format: 'bold' } as const;
    e.exec(all);
    expect(shown(e.snapshot().notesFormatted!)).toEqual([['Due 10/31/2026 ok[b]']]);
    e.exec(all);
    expect(shown(e.snapshot().notesFormatted!)).toEqual([['Due 10/31/2026 ok[]']]);
    // Only the field: it takes the format, the runs around do not.
    e.exec({ type: 'formatNotes', index: 0, start: 6, end: 8, format: 'italic' });
    expect(shown(e.snapshot().notesFormatted!)).toEqual([['Due []', '10/31/2026[i]', ' ok[]']]);
  });

  test('a line break in the range does not keep bold from turning off', async () => {
    const notes =
      '<a:p><a:r><a:rPr lang="en-US" b="1"/><a:t>ab</a:t></a:r><a:br><a:rPr lang="en-US"/></a:br>' +
      '<a:r><a:rPr lang="en-US" b="1"/><a:t>cd</a:t></a:r></a:p>';
    const e = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    expect((await e.load(buildPptx({ slides: [{ shapes: '', notesXml: notes }] }))).ok).toBe(true);
    e.exec({ type: 'formatNotes', index: 0, start: 0, end: 5, format: 'bold' });
    expect(e.snapshot().notesFormatted![0]!.runs.every((r) => !r.bold || r.text === '\n')).toBe(
      true
    );
    expect(
      e.can({ type: 'formatNotes', index: 0, start: 0, end: 5, format: 'toString' as 'bold' })
    ).toMatchObject({ ok: false, reason: 'errors.unsupported' });
  });

  test('an edit inside a field keeps the field text formatted as it was', async () => {
    const field =
      '<a:p><a:r><a:rPr lang="en-US" b="1"/><a:t>Date: </a:t></a:r>' +
      '<a:fld id="{B2C3D4E5-0000-4000-8000-000000000002}" type="datetime1"><a:rPr lang="en-US" i="1"/><a:t>10/31/2026</a:t></a:fld></a:p>';
    const e = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    expect((await e.load(buildPptx({ slides: [{ shapes: '', notesXml: field }] }))).ok).toBe(true);
    e.exec({ type: 'setNotes', index: 0, text: 'Date: 10/31/2027' });
    const r = readPptx(e.save()!);
    if (!r.ok) throw new Error('reopen');
    expect(shown(notesParagraphs(r.pkg, 0))).toEqual([['Date: [b]', '10/31/2027[i]']]);
    // The edited field is static text now.
    expect(JSON.stringify(r.pkg.parts.get('/ppt/notesSlides/notesSlide1.xml'))).not.toContain(
      'datetime1'
    );
  });

  test('refuses an empty range, notes that do not exist, and viewing mode', async () => {
    const e = await open();
    expect(e.can({ type: 'formatNotes', index: 1, start: 3, end: 3, format: 'bold' })).toEqual({
      ok: false,
      reason: 'disabled.selectNotesText',
    });
    expect(e.can({ type: 'formatNotes', index: 0, start: 0, end: 2, format: 'bold' })).toEqual({
      ok: false,
      reason: 'disabled.noNotes',
    });
    e.setMode('view');
    expect(e.can({ type: 'formatNotes', index: 1, start: 0, end: 2, format: 'bold' }).ok).toBe(
      false
    );
  });
});

describe('diffHunks', () => {
  const apply = (a: string, b: string): string => {
    let out = a;
    for (const h of [...diffHunks(a, b)].reverse())
      out = out.slice(0, h.aStart) + b.slice(h.bStart, h.bEnd) + out.slice(h.aEnd);
    return out;
  };

  test('finds each changed region and leaves the rest', () => {
    expect(diffHunks('one two three', 'one 2 three!')).toEqual([
      { aStart: 4, aEnd: 7, bStart: 4, bEnd: 5 },
      { aStart: 13, aEnd: 13, bStart: 11, bEnd: 12 },
    ]);
    expect(diffHunks('same', 'same')).toEqual([]);
    for (const [a, b] of [
      ['', 'new'],
      ['old', ''],
      ['abcabba', 'cbabac'],
      ['a\nb\nc', 'a\nB\nc\nd'],
      ['😀 x 😁', '😁 x 😀'],
    ])
      expect(apply(a!, b!)).toBe(b!);
  });

  test('gives one region over the changed middle past the edit limit', () => {
    expect(diffHunks('a1b2c3d', 'aXbYcZd', 2)).toEqual([
      { aStart: 1, aEnd: 6, bStart: 1, bEnd: 6 },
    ]);
  });
});
