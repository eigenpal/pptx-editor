import { describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { createPptxEditor, resolveEditorModules } from '@pptx-editor.dev/core/editor';
import { canonicalPackageFingerprint } from '@pptx-editor.dev/core/store';
import { buildPptx, para, run, shape } from '../../../../scripts/fixtures/builder.ts';
import { collaborationModule, reviewModule } from '../index.ts';
import { memoryRoom } from './memory-room.ts';

const FIXTURES = join(import.meta.dir, '../../../../fixtures');
const IN = 914_400;

describe('reviewModule', () => {
  test('reads legacy comments and their reply from the slide-structure fixture', async () => {
    const editor = createPptxEditor({
      fixedMeasurer: true,
      substitutes: false,
      modules: [reviewModule()],
    });
    const bytes = new Uint8Array(readFileSync(join(FIXTURES, 'slide-structure.pptx')));
    expect((await editor.load(bytes)).ok).toBe(true);
    const review = resolveEditorModules([reviewModule()]).review!;
    const items = review.collectReviewItems({ package: editor.session!.store.package });
    expect(items.map((i) => [i.kind, i.text, i.author])).toEqual([
      ['comment', 'Check the chart numbers.', 'Fixture Author'],
      ['comment', 'Checked.', 'Fixture Author'],
    ]);
    expect(items[1].parentId).toBe(items[0].id);
    expect(items[0].date).toBe('2026-01-01T09:00:00.000');
  });

  test('a deck without comments has an empty queue', async () => {
    const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    const bytes = new Uint8Array(readFileSync(join(FIXTURES, 'text-runs.pptx')));
    expect((await editor.load(bytes)).ok).toBe(true);
    const review = reviewModule().review!;
    expect(review.collectReviewItems({ package: editor.session!.store.package })).toEqual([]);
  });
});

describe('collaborationModule', () => {
  test('two editors in one room converge on each other edits', async () => {
    const room = memoryRoom();
    const bytes = buildPptx({
      slides: [
        {
          layout: 'title',
          shapes: shape({ id: 2, x: IN, y: IN, cx: 4 * IN, cy: IN, text: para(run('Hi')) }),
        },
      ],
    });
    const open = async (actorId: string) => {
      const editor = createPptxEditor({
        fixedMeasurer: true,
        substitutes: false,
        modules: [collaborationModule({ session: room.session({ actorId, name: actorId }) })],
      });
      expect((await editor.load(bytes)).ok).toBe(true);
      return editor;
    };
    const a = await open('a');
    const b = await open('b');
    const fingerprint = (e: typeof a) => canonicalPackageFingerprint(e.session!.store.package);

    expect(a.exec({ type: 'addSlide' }).ok).toBe(true);
    expect(fingerprint(b)).toBe(fingerprint(a));
    expect(b.exec({ type: 'insertTextBox' }).ok).toBe(true);
    expect(fingerprint(a)).toBe(fingerprint(b));
    expect(b.session!.info.slides.length).toBe(2);

    a.destroy();
    b.destroy();
  });
});
