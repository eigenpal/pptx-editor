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
 * Modeled on representative Office.js presentation API samples ("get the selected slides",
 * "select slides", "get the selected shapes", "select shapes", "get the selected text",
 * "check the shape types of a placeholder"). See README.md for sources and license scope.
 */
import * as PptxEditor from '../../../src/index.ts';

declare const runtime: PptxEditor.PptxEditorServerRuntime;

export async function logSelectedSlides(): Promise<void> {
  await runtime.run(async (context) => {
    const selectedSlides = context.presentation.getSelectedSlides();
    const slideCount = selectedSlides.getCount();
    selectedSlides.load('items/id');
    await context.sync();

    console.log(`${slideCount.value} slides selected.`);
    selectedSlides.items.forEach((slide) => console.log(slide.id));
  });
}

export async function selectSlides(slideIds: string[]): Promise<void> {
  await runtime.run(async (context) => {
    context.presentation.setSelectedSlides(slideIds);
    await context.sync();
  });
}

export async function selectAllShapes(): Promise<void> {
  await runtime.run(async (context) => {
    const slide = context.presentation.getSelectedSlides().getItemAt(0);
    slide.shapes.load('items/id');
    await context.sync();

    const shapeIds = slide.shapes.items.map((shape) => shape.id);
    slide.setSelectedShapes(shapeIds);
    await context.sync();
  });
}

export async function fillSelectedShapes(): Promise<void> {
  await runtime.run(async (context) => {
    const shapes = context.presentation.getSelectedShapes();
    const shapeCount = shapes.getCount();
    shapes.load('items/fill/type');
    await context.sync();

    console.log(`${shapeCount.value} shapes selected.`);
    shapes.items.forEach((shape) => shape.fill.setSolidColor('#FFFF00'));
    await context.sync();
  });
}

export async function logSelectedText(): Promise<void> {
  await runtime.run(async (context) => {
    const textRange = context.presentation.getSelectedTextRangeOrNullObject();
    textRange.load('text');
    await context.sync();

    if (!textRange.isNullObject) console.log(`Selected text: ${textRange.text}`);
  });
}

export async function logPlaceholders(): Promise<void> {
  await runtime.run(async (context) => {
    const shapes = context.presentation.slides.getItemAt(0).shapes;
    shapes.load('items/type,items/name,items/placeholderFormat/type');
    await context.sync();

    for (const shape of shapes.items) {
      if (shape.type === PptxEditor.ShapeType.placeholder) {
        const type = shape.placeholderFormat.type;
        if (type === PptxEditor.PlaceholderType.cameo) {
          console.log(`${shape.name} holds a ${type} placeholder.`);
        } else {
          console.log(`${shape.name} is a ${type} placeholder.`);
        }
      }
    }
  });
}
