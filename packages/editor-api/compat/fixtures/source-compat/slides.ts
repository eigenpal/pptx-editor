/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
/*
Office.js sample portions: Copyright (c) Microsoft Corporation.
Licensed under the MIT License. See MICROSOFT-LICENSE.txt and README.md in this directory.
The EigenPal terms above apply to EigenPal additions, not to the Microsoft sample portions.
*/
/**
 * Adapts Office.js presentation examples for compile-time compatibility checks.
 * See README.md for pinned sources, license scope, and adaptation details.
 *
 * This file is never executed. `compat/tsconfig.json` checks it against the public types.
 */
import type * as PptxEditor from '../../../src/index.ts';

declare const runtime: PptxEditor.PptxEditorServerRuntime;

export async function addSlide(): Promise<void> {
  await runtime.run(async (context) => {
    context.presentation.slides.add();
    await context.sync();
  });
}

export async function addSlideWithLayout(slideMasterId: string, layoutId: string): Promise<void> {
  await runtime.run(async (context) => {
    context.presentation.slides.add({
      slideMasterId: slideMasterId,
      layoutId: layoutId,
    });
    await context.sync();
  });
}

export async function logSlideMastersAndLayouts(): Promise<void> {
  await runtime.run(async (context) => {
    context.presentation.slideMasters.load('id, name, layouts/items/name, layouts/items/id');
    await context.sync();

    for (let i = 0; i < context.presentation.slideMasters.items.length; i++) {
      console.log('Master name: ' + context.presentation.slideMasters.items[i]!.name);
      console.log('Master ID: ' + context.presentation.slideMasters.items[i]!.id);
      const layoutsInMaster = context.presentation.slideMasters.items[i]!.layouts;
      for (let j = 0; j < layoutsInMaster.items.length; j++) {
        console.log('    Layout name: ' + layoutsInMaster.items[j]!.name);
        console.log('    Layout ID: ' + layoutsInMaster.items[j]!.id);
      }
    }
  });
}

export async function logShapeTextOfEachSlide(): Promise<void> {
  await runtime.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load('items/shapes/items/name, items/shapes/items/textFrame/textRange/text');
    const first = context.presentation.slides.getItemAt(0);
    first.load({ $all: true, layout: { name: true }, shapes: { $top: 5, name: true } });
    await context.sync();
    for (const slide of slides.items)
      for (const shape of slide.shapes.items)
        console.log(shape.name + ': ' + shape.textFrame.textRange.text);
    console.log(first.index, first.layout.name, first.shapes.items.length);
  });
}

export async function deleteSecondSlide(): Promise<void> {
  await runtime.run(async (context) => {
    const slide = context.presentation.slides.getItemAt(1);
    slide.delete();
    await context.sync();
  });
}

export async function moveLastSlideToFront(): Promise<void> {
  await runtime.run(async (context) => {
    const slideCount = context.presentation.slides.getCount();
    await context.sync();
    const lastSlide = context.presentation.slides.getItemAt(slideCount.value - 1);
    lastSlide.moveTo(0);
    await context.sync();
  });
}
