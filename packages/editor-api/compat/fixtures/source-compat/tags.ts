/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
/*
Office.js sample portions: Copyright (c) Microsoft Corporation.
Licensed under the MIT License. See MICROSOFT-LICENSE.txt and README.md in this directory.
The EigenPal terms above apply to EigenPal additions, not to the Microsoft sample portions.
*/
/**
 * Modeled on representative Office.js presentation API samples ("add tags to a slide and to
 * the presentation", "read the tags", "get a tag by key", "delete a tag"). See README.md for
 * sources and license scope.
 */
import type * as PptxEditor from '../../../src/index.ts';

declare const runtime: PptxEditor.PptxEditorServerRuntime;

export async function addTags(): Promise<void> {
  await runtime.run(async (context) => {
    const slide = context.presentation.getSelectedSlides().getItemAt(0);
    slide.tags.add('OCEAN', 'Arctic');

    const presentationTags = context.presentation.tags;
    presentationTags.add('COLOR', 'blue');
    await context.sync();
  });
}

export async function logSlideTags(): Promise<void> {
  await runtime.run(async (context) => {
    const tags = context.presentation.slides.getItemAt(0).tags;
    tags.load('key, value');
    await context.sync();

    for (let i = 0; i < tags.items.length; i++) {
      console.log('Key: ' + tags.items[i]!.key + ', Value: ' + tags.items[i]!.value);
    }
  });
}

export async function readTag(): Promise<string | null> {
  let value: string | null = null;
  await runtime.run(async (context) => {
    const tag = context.presentation.tags.getItemOrNullObject('COLOR');
    tag.load('value');
    await context.sync();

    if (!tag.isNullObject) value = tag.value;
  });
  return value;
}

export async function deleteTag(): Promise<void> {
  await runtime.run(async (context) => {
    const presentationTags = context.presentation.tags;
    presentationTags.delete('COLOR');
    await context.sync();
  });
}
