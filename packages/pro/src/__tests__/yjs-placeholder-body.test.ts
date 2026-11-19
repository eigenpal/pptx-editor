/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// A text placeholder without `p:txBody` on two replicas (AGENTS.md, "Collaboration
// verification"): both peers type into it while the link is paused, so each adds a body. After
// the link delivers, both replicas show one body with the text of both peers, and save the same
// package, in both client orders.

import { describe, expect, test } from 'bun:test';
import type { PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import { canonicalPackageFingerprint, readPptx } from '@pptx-editor.dev/core/store';
import { buildPptx } from '../../../../scripts/fixtures/builder.ts';
import { kids, peers, root, textOf, TWO_ORDERS } from './lineage-support.ts';

const deck = (): Uint8Array =>
  buildPptx({
    slides: [
      {
        shapes:
          '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr><a:spLocks noGrp="1"/>' +
          '</p:cNvSpPr><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr/></p:sp>',
      },
    ],
  });

const title = (e: PptxEditorInstance) =>
  kids(kids(kids(root(e)).find((c) => c.local === 'cSld')!)[0]!).find((c) => c.local === 'sp')!;

const caretIn = (e: PptxEditorInstance) => {
  const at = { paragraph: 0, offset: 0 };
  e.setSelection({ kind: 'text', nodeId: title(e).id, anchor: at, focus: at });
};

describe('a placeholder without a text body on two replicas', () => {
  for (const ids of TWO_ORDERS)
    test(`both peers type at once (client ids ${ids.join(', ')})`, async () => {
      const { editors, sync } = await peers(deck(), ids);
      const [a, b] = editors as [PptxEditorInstance, PptxEditorInstance];
      sync.pause();
      caretIn(a);
      caretIn(b);
      for (const ch of 'Left') expect(a.input('text', ch).ok).toBe(true);
      for (const ch of 'Right') expect(b.input('text', ch).ok).toBe(true);
      sync.flush();
      const bodies = (e: PptxEditorInstance) => kids(title(e)).filter((c) => c.local === 'txBody');
      expect(bodies(a)).toHaveLength(1);
      expect(bodies(b)).toHaveLength(1);
      const text = textOf(bodies(a)[0]!);
      expect(textOf(bodies(b)[0]!)).toBe(text);
      expect(text).toContain('Left');
      expect(text).toContain('Right');
      const [ra, rb] = [readPptx(a.save()!), readPptx(b.save()!)];
      if (!ra.ok || !rb.ok) throw new Error('a saved copy does not open');
      expect(canonicalPackageFingerprint(ra.pkg)).toBe(canonicalPackageFingerprint(rb.pkg));
    });
});
