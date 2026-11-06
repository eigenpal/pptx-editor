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
 * Modeled on representative Office.js presentation API samples ("add a table with values",
 * "format a table cell", "add rows and columns", "read the table values"). See README.md for
 * sources and license scope.
 */
import * as PptxEditor from '../../../src/index.ts';

declare const runtime: PptxEditor.PptxEditorServerRuntime;

export async function addTableWithValues(): Promise<void> {
  await runtime.run(async (context) => {
    const shapes = context.presentation.slides.getItemAt(0).shapes;
    shapes.addTable(3, 4, {
      values: [
        ['A', 'B', 'C', 'D'],
        ['1', '2', '3', '4'],
        ['5', '6', '7', '8'],
      ],
      left: 100,
      top: 100,
      width: 400,
    });
    await context.sync();
  });
}

export async function formatHeaderCell(): Promise<void> {
  await runtime.run(async (context) => {
    const shapes = context.presentation.slides.getItemAt(0).shapes;
    const table = shapes.getItemAt(0).getTable();
    const cell = table.getCellOrNullObject(0, 0);
    cell.text = 'Region';
    cell.font.bold = true;
    cell.font.color = '#FFFFFF';
    cell.fill.setSolidColor('#1F4E79');
    cell.horizontalAlignment = PptxEditor.ParagraphHorizontalAlignment.center;
    cell.verticalAlignment = PptxEditor.TextVerticalAlignment.middle;
    await context.sync();
  });
}

export async function addRowAndColumn(): Promise<void> {
  await runtime.run(async (context) => {
    const table = context.presentation.slides.getItemAt(0).shapes.getItemAt(0).getTable();
    table.rows.add(1, 1);
    table.columns.add(1, 1);
    await context.sync();
  });
}

export async function logTableValues(): Promise<void> {
  await runtime.run(async (context) => {
    const table = context.presentation.slides.getItemAt(0).shapes.getItemAt(0).getTable();
    table.load('rowCount,columnCount,values');
    await context.sync();

    console.log(`Rows: ${table.rowCount}, columns: ${table.columnCount}`);
    table.values.forEach((row) => console.log(row.join(' | ')));
  });
}
