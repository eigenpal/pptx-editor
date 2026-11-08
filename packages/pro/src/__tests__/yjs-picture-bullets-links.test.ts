/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/

// Two peers converge on a picture bullet (`setBulletFormat` with a picture: a new media part and
// `a:buBlip`), a custom show link, and a link to another file. The gate's one fixture per
// command covers a character bullet and a web link only.

import { describe, test } from 'bun:test';
import type { PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import { buildPptx, para, run, shape } from '../../../../scripts/fixtures/builder.ts';
import { converge, png } from './convergence-harness.ts';

const IN = 914_400;

/** One text shape `A` and a custom show of the first slide. */
const deck = (): Uint8Array =>
  buildPptx({
    slides: [
      {
        shapes: shape({
          id: 2,
          name: 'A',
          x: IN,
          y: IN,
          cx: 4 * IN,
          cy: IN,
          text: para(run('Hello world')),
        }),
      },
    ],
    presentationLists:
      '<p:custShowLst><p:custShow name="Short" id="3"><p:sldLst><p:sld r:id="rId101"/>' +
      '</p:sldLst></p:custShow></p:custShowLst>',
  });

const selectA = (e: PptxEditorInstance): void => {
  const id = e.snapshot().shapes!.find((s) => s.name === 'A')!.nodeId;
  e.setSelection({
    kind: 'text',
    nodeId: id,
    anchor: { paragraph: 0, offset: 0 },
    focus: { paragraph: 0, offset: 5 },
  });
};

describe('collaboration: picture bullets and links', () => {
  test('a picture bullet converges', () =>
    converge('setBulletFormat', {
      deck,
      arrange: selectA,
      command: () => ({
        type: 'setBulletFormat',
        format: { kind: 'picture', picture: png(6, 6), sizePercent: 150 },
      }),
    }));

  test('a custom show link converges', () =>
    converge('setLink', {
      deck,
      arrange: selectA,
      command: () => ({ type: 'setLink', url: '#custom-show-3?return=true' }),
    }));

  test('a link to another file converges', () =>
    converge('setLink', {
      deck,
      arrange: selectA,
      command: () => ({ type: 'setLink', url: '../shared/report.pdf' }),
    }));
});
