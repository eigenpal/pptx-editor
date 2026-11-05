/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
/*
Office.js sample portions: Copyright (c) Microsoft Corporation.
Licensed under the MIT License. See MICROSOFT-LICENSE.txt and README.md in this directory.
The EigenPal terms above apply to EigenPal additions, not to the Microsoft sample portions.
*/
/**
 * Modeled on representative Office.js presentation API samples ("add a text box", "format
 * the text of a shape", "read the text of every shape"). See README.md for sources and license scope.
 */
import * as PptxEditor from '../../../src/index.ts';

declare const runtime: PptxEditor.PptxEditorServerRuntime;

export async function addTextBox(): Promise<void> {
  await runtime.run(async (context) => {
    const shapes = context.presentation.slides.getItemAt(0).shapes;
    const textbox = shapes.addTextBox('Hello!', {
      left: 100,
      top: 300,
      height: 300,
      width: 450,
    });
    textbox.name = 'Textbox';
    await context.sync();
  });
}

export async function formatTextBox(): Promise<void> {
  await runtime.run(async (context) => {
    const shapes = context.presentation.slides.getItemAt(0).shapes;
    const textbox = shapes.addTextBox('Quarterly results', {
      left: 100,
      top: 100,
      height: 80,
      width: 450,
    });
    const textRange = textbox.textFrame.textRange;
    textRange.font.bold = true;
    textRange.font.italic = false;
    textRange.font.color = '#7030A0';
    textRange.font.name = 'Arial';
    textRange.font.size = 28;
    textRange.font.underline = PptxEditor.ShapeFontUnderlineStyle.single;
    textRange.paragraphFormat.horizontalAlignment = PptxEditor.ParagraphHorizontalAlignment.center;
    textbox.textFrame.verticalAlignment = PptxEditor.TextVerticalAlignment.middleCentered;
    textbox.textFrame.autoSizeSetting = 'AutoSizeShapeToFitText';
    textbox.textFrame.wordWrap = true;
    await context.sync();
  });
}

export async function replaceTextOfShapes(): Promise<void> {
  await runtime.run(async (context) => {
    const shapes = context.presentation.slides.getItemAt(0).shapes;
    shapes.load('items/name,items/textFrame/hasText,items/textFrame/textRange/text');
    await context.sync();

    shapes.items.forEach((shape) => {
      if (shape.textFrame.hasText) {
        console.log(shape.name + ': ' + shape.textFrame.textRange.text);
        shape.textFrame.textRange.text = shape.textFrame.textRange.text.toUpperCase();
      }
    });
    await context.sync();
  });
}

export async function boldFirstWord(): Promise<void> {
  await runtime.run(async (context) => {
    const shape = context.presentation.slides.getItemAt(0).shapes.getItemAt(0);
    const firstWord = shape.textFrame.textRange.getSubstring(0, 5);
    firstWord.font.bold = true;
    firstWord.font.color = '#FF0000';
    await context.sync();
  });
}
