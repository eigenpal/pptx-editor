/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Two editors on linked Y.Docs that hold one package save the same bytes: with no edit, after
// an edit that adds an attribute out of sorted order, and after concurrent attribute writes on
// one element. The replica that loaded the file keeps the file's attribute order, so the room
// keeps that order for every peer (`ao` in the record shell), in both client orders.

import { describe, expect, test } from 'bun:test';
import type { PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import { buildPptx, para, run, shape } from '../../../../scripts/fixtures/builder.ts';
import { all, IN, kids, peers, root, TWO_ORDERS } from './lineage-support.ts';

/** One text box whose `a:off` and `a:rPr` keep their attributes out of sorted order. */
const deck = (): Uint8Array =>
  buildPptx({
    slides: [
      {
        shapes: shape({
          id: 2,
          name: 'Box',
          x: IN,
          y: IN,
          cx: 3 * IN,
          cy: IN,
          text: para(run('Hello world', 'sz="2400"')),
        }),
      },
    ],
  });

const box = (e: PptxEditorInstance): string =>
  kids(kids(root(e)).find((c) => c.local === 'cSld')!)
    .flatMap((c) => kids(c))
    .find((c) => c.local === 'sp')!.id;

/** The saved bytes of each editor, as one string each. */
function saves(editors: readonly PptxEditorInstance[]): string[] {
  return editors.map((e) => Buffer.from(e.save()!).toString('latin1'));
}

describe('two replicas of one package', () => {
  for (const ids of TWO_ORDERS) {
    test(`save the same bytes with no edit (client ids ${ids.join(', ')})`, async () => {
      const { editors } = await peers(deck(), ids);
      const [a, b] = saves(editors);
      expect(b === a).toBe(true);
    });

    test(`save the same bytes after an edit (client ids ${ids.join(', ')})`, async () => {
      const { editors } = await peers(deck(), ids);
      const [a, b] = editors as [PptxEditorInstance, PptxEditorInstance];
      a.setSelection({
        kind: 'text',
        nodeId: box(a),
        anchor: { paragraph: 0, offset: 0 },
        focus: { paragraph: 0, offset: 5 },
      });
      // Bold adds `b` after `lang`, `dirty`, and `sz` on the run properties.
      expect(a.exec({ type: 'toggleBold' }).ok).toBe(true);
      const bold = (e: PptxEditorInstance) =>
        all(root(e), 'rPr').map((r) => r.attrs.map((x) => x.local).join(' '));
      expect(bold(b)).toEqual(['lang dirty sz b', 'lang dirty sz']);
      const [sa, sb] = saves([a, b]);
      expect(sb === sa).toBe(true);
    });

    test(`save the same bytes after concurrent attribute writes (client ids ${ids.join(', ')})`, async () => {
      const { editors, sync } = await peers(deck(), ids);
      const [a, b] = editors as [PptxEditorInstance, PptxEditorInstance];
      sync.pause();
      a.setSelection({ kind: 'shapes', nodeIds: [box(a)] });
      b.setSelection({ kind: 'shapes', nodeIds: [box(b)] });
      expect(a.exec({ type: 'rotateObjects', degrees: 90 }).ok).toBe(true);
      expect(b.exec({ type: 'nudge', dx: IN / 10, dy: IN / 10 }).ok).toBe(true);
      sync.flush();
      const [sa, sb] = saves([a, b]);
      expect(sb === sa).toBe(true);
    });
  }
});
