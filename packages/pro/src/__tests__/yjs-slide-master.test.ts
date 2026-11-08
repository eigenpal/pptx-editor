/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Layout edits in the slide master view, in a collaboration room: a peer that edits a layout
// and a peer that edits a slide on it at the same time converge, and both see the layout edit.

import { describe, expect, test } from 'bun:test';
import type { PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import { serializeXmlPart, xmlPartOf } from '@pptx-editor.dev/core/store';
import { buildPptx, para, run, shape } from '../../../../scripts/fixtures/builder.ts';
import { openPeers } from '../../../../scripts/fidelity/collaboration-kit.ts';

const IN = 914_400;

/** One slide with shape A on a layout with a Header shape. */
const deck = (): Uint8Array =>
  buildPptx({
    layouts: [
      {
        name: 'Branded',
        shapes: shape({
          id: 9,
          name: 'Header',
          x: IN / 4,
          y: IN / 8,
          cx: 3 * IN,
          cy: IN / 2,
          text: para(run('Brand')),
        }),
      },
    ],
    slides: [
      {
        layoutName: 'Branded',
        shapes: shape({ id: 2, name: 'A', x: IN, y: 2 * IN, cx: IN, cy: IN }),
      },
    ],
  });

const select = (e: PptxEditorInstance, name: string): void =>
  e.setSelection({
    kind: 'shapes',
    nodeIds: [e.snapshot().shapes!.find((s) => s.name === name)!.nodeId],
  });

const partXml = (e: PptxEditorInstance, part: string): string =>
  serializeXmlPart(xmlPartOf(e.session!.store.package, part)!);

describe('the slide master view in a collaboration room', () => {
  test('a layout edit and a slide edit at the same time converge', async () => {
    const { editor: a, peer: b, link, destroy } = await openPeers(deck());
    try {
      expect(a.exec({ type: 'setPresentationView', view: 'slideMaster' }).ok).toBe(true);
      const layout = a.snapshot().slides.find((s) => s.name === 'Branded')!;
      a.goToSlide(layout.index);
      link.pause();
      select(a, 'Header');
      expect(a.exec({ type: 'nudge', dx: IN, dy: 0 }).ok).toBe(true);
      select(b, 'A');
      expect(b.exec({ type: 'nudge', dx: 0, dy: IN }).ok).toBe(true);
      link.flush();
      // The peer's slide edit leaves the view and the layout on the canvas.
      expect(a.snapshot().presentationView).toBe('slideMaster');
      expect(a.snapshot().currentSlide).toBe(layout.index);
      // An undo in the view reverts the layout edit, and the canvas stays on the layout.
      expect(a.exec({ type: 'undo' }).ok).toBe(true);
      expect(a.snapshot().currentSlide).toBe(layout.index);
      expect(a.exec({ type: 'redo' }).ok).toBe(true);
      for (const e of [a, b]) {
        expect(partXml(e, layout.id)).toContain(`<a:off x="${IN / 4 + IN}"`);
        expect(partXml(e, '/ppt/slides/slide1.xml')).toContain(`y="${3 * IN}"`);
      }
      // The peer on the slides shows the moved header from the layout.
      const header = b.session!.layout(0)!.items.find((i) => i.element.name === 'Header')!;
      expect(header.element.xfrm.x).toBe(IN / 4 + IN);
    } finally {
      destroy();
    }
  });

  test('Close master returns to the same slide after a peer deletes an earlier one', async () => {
    const two = buildPptx({ slides: [{ shapes: '' }, { shapes: '' }, { shapes: '' }] });
    const { editor: a, peer: b, destroy } = await openPeers(two);
    try {
      a.goToSlide(1);
      const part = a.session!.info.slides[1]!.partName;
      expect(a.exec({ type: 'setPresentationView', view: 'slideMaster' }).ok).toBe(true);
      expect(b.exec({ type: 'deleteSlide', index: 0 }).ok).toBe(true);
      expect(a.exec({ type: 'setPresentationView', view: 'normal' }).ok).toBe(true);
      expect(a.session!.info.slides[a.snapshot().currentSlide]!.partName).toBe(part);
    } finally {
      destroy();
    }
  });
});
