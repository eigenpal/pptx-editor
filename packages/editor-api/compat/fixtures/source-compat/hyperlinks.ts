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
 * Modeled on representative Office.js presentation API samples ("list the hyperlinks of a
 * slide", "add a hyperlink to text", "add a hyperlink to a shape"). See README.md for sources and license scope.
 */
import * as PptxEditor from '../../../src/index.ts';

declare const runtime: PptxEditor.PptxEditorServerRuntime;

export async function logHyperlinks(): Promise<void> {
  await runtime.run(async (context) => {
    const hyperlinks = context.presentation.slides.getItemAt(0).hyperlinks;
    hyperlinks.load('address,screenTip,type');
    await context.sync();

    console.log(`${hyperlinks.items.length} hyperlinks found in the first slide.`);
    for (const link of hyperlinks.items) {
      const where = link.type === PptxEditor.HyperlinkType.shape ? 'shape' : 'text';
      console.log(`Address: "${link.address}" (${where}), screen tip: "${link.screenTip}"`);
    }
  });
}

export async function linkSelectedText(): Promise<void> {
  await runtime.run(async (context) => {
    const textRange = context.presentation.getSelectedTextRange();
    textRange.setHyperlink({
      address: 'https://example.com',
      screenTip: 'Open the example site',
    });
    await context.sync();
  });
}

export async function linkFirstShape(): Promise<void> {
  await runtime.run(async (context) => {
    const shape = context.presentation.slides.getItemAt(0).shapes.getItemAt(0);
    const link = shape.setHyperlink({ address: 'https://example.com/report' });
    link.screenTip = 'Read the report';
    await context.sync();
  });
}
