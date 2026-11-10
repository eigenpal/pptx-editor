/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
import { expect, test } from 'bun:test';
import type { PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import { converged, oneRun, peers, selectIn } from './lineage-support.ts';

const IDS = [
  [1_001, 2_002, 3_003],
  [3_003, 2_002, 1_001],
  [2_002, 1_001, 3_003],
  [1_001, 3_003, 2_002],
  [2_002, 3_003, 1_001],
  [3_003, 1_001, 2_002],
];
type Ed = PptxEditorInstance;
const STAGES = ['all', 'enter-author-first', 'undoer-last', 'bystander-first'] as const;

for (const ids of IDS)
  for (const stage of STAGES)
    test(`format undo vs enter ${ids} ${stage}`, async () => {
      const { editors, sync, docs } = await peers(oneRun('ABCDEFGH'), ids);
      const [a, b, c] = editors as [Ed, Ed, Ed];
      sync.pause();
      selectIn(c, 0, 1);
      c.input('paragraph');
      selectIn(b, 0, 2, 4);
      b.exec({ type: 'toggleItalic' });
      if (stage === 'enter-author-first') sync.deliverTo(docs[2]!);
      if (stage === 'bystander-first') sync.deliverTo(docs[0]!);
      b.exec({ type: 'undo' });
      if (stage === 'undoer-last') {
        sync.deliverTo(docs[0]!);
        sync.deliverTo(docs[2]!);
      }
      sync.flush();
      try {
        expect(converged(a, b, c)).toEqual(['A', 'BCDEFGH']);
      } finally {
        editors.forEach((e) => e.destroy());
      }
    });
