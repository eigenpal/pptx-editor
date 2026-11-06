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
 * Modeled on representative Office.js presentation API samples ("add a geometric shape",
 * "add a line", "move and resize a shape", "change the z-order", "group shapes", "delete all
 * shapes"). See README.md for sources and license scope.
 */
import * as PptxEditor from '../../../src/index.ts';

declare const runtime: PptxEditor.PptxEditorServerRuntime;

export async function addHexagon(): Promise<void> {
  await runtime.run(async (context) => {
    const shapes = context.presentation.slides.getItemAt(0).shapes;
    const hexagon = shapes.addGeometricShape(PptxEditor.GeometricShapeType.hexagon, {
      left: 100,
      top: 100,
      height: 150,
      width: 150,
    });
    hexagon.name = 'Hexagon';
    hexagon.fill.setSolidColor('#FFA500');
    hexagon.fill.transparency = 0.25;
    hexagon.lineFormat.color = '#0000FF';
    hexagon.lineFormat.weight = 3;
    hexagon.lineFormat.dashStyle = PptxEditor.ShapeLineDashStyle.dashDot;
    await context.sync();
  });
}

export async function addStraightLine(): Promise<void> {
  await runtime.run(async (context) => {
    const shapes = context.presentation.slides.getItemAt(0).shapes;
    const line = shapes.addLine(PptxEditor.ConnectorType.straight, {
      left: 200,
      top: 50,
      height: 300,
      width: 150,
    });
    line.name = 'StraightLine';
    line.lineFormat.style = 'Single';
    await context.sync();
  });
}

export async function moveResizeAndRaise(): Promise<void> {
  await runtime.run(async (context) => {
    const shapes = context.presentation.slides.getItemAt(0).shapes;
    const shape = shapes.getItem('Hexagon');
    shape.left = 50;
    shape.top = 50;
    shape.width = 200;
    shape.height = 200;
    shape.rotation = 45;
    shape.setZOrder(PptxEditor.ShapeZOrder.bringToFront);
    await context.sync();
  });
}

export async function groupShapes(): Promise<void> {
  await runtime.run(async (context) => {
    const shapes = context.presentation.slides.getItemAt(0).shapes;
    shapes.load('items/id');
    await context.sync();

    const shapeIds = shapes.items.map((shape) => shape.id);
    const group = shapes.addGroup(shapeIds);
    group.name = 'Group';
    await context.sync();
  });
}

export async function deleteAllShapes(): Promise<void> {
  await runtime.run(async (context) => {
    const shapes = context.presentation.slides.getItemAt(0).shapes;
    shapes.load('items/type,items/name');
    await context.sync();

    shapes.items.forEach((shape) => shape.delete());
    await context.sync();
  });
}
