/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Scenarios for text: text frames, text ranges, fonts, paragraph format, bullets, autofit,
// and hyperlinks.

import type { RequestContext, ShapeFont, TextRange } from '../../runtime/public.ts';
import { check, pt, same, type Json, type Scenario } from './types.ts';

const FONT = 'name,size,bold,italic,color,underline,allCaps,smallCaps,strikethrough,superscript';

function fontJson(f: ShapeFont): Json {
  return {
    name: f.name,
    size: f.size,
    bold: f.bold,
    italic: f.italic,
    color: f.color,
    underline: f.underline,
    allCaps: f.allCaps,
    smallCaps: f.smallCaps,
    strikethrough: f.strikethrough,
    superscript: f.superscript,
  };
}

/** The text box on the fourth slide of `slide-layouts`. */
const textBox = (context: RequestContext) =>
  context.presentation.slides.getItemAt(3).shapes.getItem('3');

/** The addresses of the text links in `textBox`, in text order. */
async function rangeLinksOf(context: RequestContext): Promise<Json> {
  const links = textBox(context).textFrame.textRange.hyperlinks.load('items/address');
  await context.sync();
  return links.items.map((l) => l.address);
}

/** The font of each named piece of a range. */
async function pieces(range: TextRange, spans: [number, number][]): Promise<Json> {
  const fonts = spans.map(([start, length]) => {
    const sub = range.getSubstring(start, length);
    sub.load('text');
    return { sub, font: sub.font.load(FONT) };
  });
  await range.context.sync();
  return fonts.map(({ sub, font }) => ({ text: sub.text, font: fontJson(font) }));
}

async function frameOf(context: RequestContext): Promise<Json> {
  const frame = textBox(context).textFrame;
  frame.load(
    'autoSizeSetting,leftMargin,rightMargin,topMargin,bottomMargin,verticalAlignment,wordWrap,hasText'
  );
  await context.sync();
  return {
    autoSize: frame.autoSizeSetting,
    margins: [frame.leftMargin, frame.rightMargin, frame.topMargin, frame.bottomMargin].map(pt),
    vertical: frame.verticalAlignment,
    wordWrap: frame.wordWrap,
    hasText: frame.hasText,
  };
}

async function paragraphsOf(context: RequestContext): Promise<Json> {
  const range = context.presentation.slides.getItemAt(1).shapes.getItem('3').textFrame.textRange;
  range.load('text');
  await context.sync();
  const lines = range.text.split('\r');
  let start = 0;
  const formats = lines.map((line) => {
    const sub = range.getSubstring(start, line.length);
    start += line.length + 1;
    return {
      line,
      format: sub.paragraphFormat.load('horizontalAlignment,indentLevel'),
      bullet: sub.paragraphFormat.bulletFormat.load('visible,type,style'),
    };
  });
  await context.sync();
  return formats.map(({ line, format, bullet }) => ({
    line,
    align: format.horizontalAlignment,
    level: format.indentLevel,
    bullet: { visible: bullet.visible, type: bullet.type, style: bullet.style },
  }));
}

async function linksOf(context: RequestContext): Promise<Json> {
  const links = context.presentation.slides.getItemAt(0).hyperlinks;
  links.load('items/address,items/screenTip,items/type');
  await context.sync();
  const targets = links.items.map((l) => ({
    shape: l.getLinkedShapeOrNullObject().load('name'),
    range: l.getLinkedTextRangeOrNullObject().load('text'),
  }));
  await context.sync();
  return links.items.map((l, i) => ({
    address: l.address,
    screenTip: l.screenTip,
    type: l.type,
    shape: targets[i]!.shape.isNullObject ? null : targets[i]!.shape.name,
    text: targets[i]!.range.isNullObject ? null : targets[i]!.range.text,
  }));
}

export const TEXT_SCENARIOS: readonly Scenario[] = [
  {
    name: 'text.frame-properties',
    area: 'text',
    deck: 'slide-layouts',
    writes: true,
    async script(context) {
      const before = await frameOf(context);
      const frame = textBox(context).textFrame;
      frame.leftMargin = 12;
      frame.rightMargin = 12;
      frame.topMargin = 6;
      frame.bottomMargin = 6;
      frame.verticalAlignment = 'Middle';
      frame.wordWrap = false;
      frame.autoSizeSetting = 'AutoSizeNone';
      await context.sync();
      return { before, after: await frameOf(context) };
    },
    readBack: frameOf,
  },
  {
    name: 'text.range-and-font',
    area: 'text',
    deck: 'slide-layouts',
    writes: true,
    async script(context) {
      const range = textBox(context).textFrame.textRange;
      range.text = 'Revenue grew 12% in Q3';
      await context.sync();
      const growth = range.getSubstring(13, 3);
      growth.font.bold = true;
      growth.font.color = '#00B050';
      growth.font.size = 28;
      const quarter = range.getSubstring(20, 2);
      quarter.font.italic = true;
      quarter.font.underline = 'Single';
      quarter.font.name = 'Georgia';
      range.getSubstring(0, 7).font.allCaps = true;
      await context.sync();
      range.load('text,start,length');
      const whole = range.font.load('bold,size');
      await context.sync();
      same(range.text, 'Revenue grew 12% in Q3', 'the new text');
      return {
        length: range.length,
        mixed: { bold: whole.bold, size: whole.size },
        pieces: await pieces(range, [
          [0, 7],
          [13, 3],
          [20, 2],
        ]),
      };
    },
    async readBack(context) {
      return pieces(textBox(context).textFrame.textRange, [
        [0, 7],
        [13, 3],
        [20, 2],
      ]);
    },
  },
  {
    name: 'text.font-effects',
    area: 'text',
    deck: 'slide-layouts',
    writes: true,
    async script(context) {
      const range = context.presentation.slides.getItemAt(0).shapes.getItem('3')
        .textFrame.textRange;
      range.text = 'H2O and E=mc2 struck small';
      await context.sync();
      range.getSubstring(1, 1).font.subscript = true;
      range.getSubstring(12, 1).font.superscript = true;
      range.getSubstring(14, 6).font.strikethrough = true;
      range.getSubstring(21, 5).font.smallCaps = true;
      await context.sync();
      const sub = range.getSubstring(1, 1).font.load('subscript,superscript');
      const dbl = range.getSubstring(14, 6);
      dbl.font.doubleStrikethrough = true;
      await context.sync();
      const strike = dbl.font.load('strikethrough,doubleStrikethrough');
      await context.sync();
      return {
        subscript: [sub.subscript, sub.superscript],
        strike: [strike.strikethrough, strike.doubleStrikethrough],
        read: await this.readBack!(context),
      };
    },
    async readBack(context) {
      const range = context.presentation.slides.getItemAt(0).shapes.getItem('3')
        .textFrame.textRange;
      return pieces(range, [
        [1, 1],
        [12, 1],
        [14, 6],
        [21, 5],
      ]);
    },
  },
  {
    name: 'text.paragraphs-and-bullets',
    area: 'text',
    deck: 'slide-layouts',
    writes: true,
    async script(context) {
      const before = await paragraphsOf(context);
      const range = context.presentation.slides.getItemAt(1).shapes.getItem('3')
        .textFrame.textRange;
      range.text = 'Goals\nGrow revenue\nCut costs\nHire';
      await context.sync();
      const goals = range.getSubstring(0, 5);
      goals.paragraphFormat.horizontalAlignment = 'Center';
      goals.paragraphFormat.bulletFormat.visible = false;
      const grow = range.getSubstring(6, 12);
      grow.paragraphFormat.indentLevel = 2;
      grow.paragraphFormat.bulletFormat.style = 'ArabicNumeralPeriod';
      const cut = range.getSubstring(19, 9);
      cut.paragraphFormat.indentLevel = 2;
      cut.paragraphFormat.bulletFormat.type = 'Unnumbered';
      range.getSubstring(29, 4).paragraphFormat.horizontalAlignment = 'Right';
      await context.sync();
      return { before, after: await paragraphsOf(context) };
    },
    readBack: paragraphsOf,
  },
  {
    name: 'text.autofit',
    area: 'text',
    deck: 'slide-layouts',
    writes: true,
    async script(context) {
      const shapes = context.presentation.slides.getItemAt(4).shapes;
      const grow = shapes.addTextBox('One line', { left: 40, top: 300, width: 200, height: 30 });
      grow.name = 'Grows';
      const shrink = shapes.addTextBox('Short', { left: 300, top: 300, width: 200, height: 60 });
      shrink.name = 'Shrinks';
      await context.sync();
      grow.textFrame.autoSizeSetting = 'AutoSizeShapeToFitText';
      grow.textFrame.wordWrap = true;
      grow.textFrame.textRange.text =
        'A longer paragraph that wraps over several lines in a narrow box';
      shrink.textFrame.autoSizeSetting = 'AutoSizeTextToFitShape';
      await context.sync();
      grow.load('height');
      const settings = [grow, shrink].map((s) => s.textFrame.load('autoSizeSetting'));
      await context.sync();
      check(grow.height > 30, 'a shape that fits its text grows with the text');
      return {
        settings: settings.map((f) => f.autoSizeSetting),
        grewTaller: grow.height > 30,
      };
    },
    async readBack(context) {
      const shapes = context.presentation.slides
        .getItemAt(4)
        .shapes.load('items/name,items/textFrame/autoSizeSetting');
      await context.sync();
      return shapes.items.map((s) => `${s.name}:${s.textFrame.autoSizeSetting}`);
    },
  },
  {
    name: 'text.hyperlinks',
    area: 'text',
    deck: 'slide-layouts',
    writes: true,
    async script(context) {
      const shapes = context.presentation.slides.getItemAt(0).shapes;
      const title = shapes.getItem('2');
      const range = shapes.getItem('3').textFrame.textRange;
      range.text = 'Read the full report online';
      await context.sync();
      const link = range.getSubstring(9, 11).setHyperlink({
        address: 'https://example.com/report',
        screenTip: 'Open the report',
      });
      title.setHyperlink({ address: 'mailto:team@example.com' });
      await context.sync();
      link.load('address,screenTip');
      const count = context.presentation.slides.getItemAt(0).hyperlinks.getCount();
      await context.sync();
      same(count.value, 2, 'two hyperlinks on the slide');
      link.screenTip = 'Opens the report';
      await context.sync();
      return { first: [link.address, link.screenTip], links: await linksOf(context) };
    },
    readBack: linksOf,
  },
  {
    name: 'text.range-hyperlinks',
    area: 'text',
    deck: 'slide-layouts',
    writes: true,
    async script(context) {
      const range = textBox(context).textFrame.textRange;
      range.text = 'See the annual report and the press kit';
      await context.sync();
      range.getSubstring(8, 13).setHyperlink({ address: 'https://example.com/report' });
      range.getSubstring(28, 11).setHyperlink({ address: 'https://example.com/press' });
      // No sync between: the lookup sees the links queued before it.
      const all = range.hyperlinks.load('items/address');
      const count = range.hyperlinks.getCount();
      // A range that starts inside the first link holds that whole link.
      const inner = range.getSubstring(12, 4).hyperlinks.load('items/address');
      const plain = range.getSubstring(0, 3).hyperlinks.getCount();
      await context.sync();
      same(count.value, 2, 'the whole text holds two links');
      same(plain.value, 0, 'the first word holds no link');
      return {
        all: all.items.map((l) => l.address),
        inner: inner.items.map((l) => l.address),
        links: await rangeLinksOf(context),
      };
    },
    readBack: rangeLinksOf,
  },
  {
    name: 'text.delete-text',
    area: 'text',
    deck: 'slide-layouts',
    writes: true,
    async script(context) {
      const frame = textBox(context).textFrame;
      frame.load('hasText');
      await context.sync();
      check(frame.hasText, 'the text box starts with text');
      frame.deleteText();
      await context.sync();
      const after = frame.load('hasText');
      const text = frame.textRange.load('text,length');
      await context.sync();
      return { hasText: after.hasText, text: text.text, length: text.length };
    },
    async readBack(context) {
      const frame = textBox(context).textFrame.load('hasText');
      await context.sync();
      return frame.hasText;
    },
  },
];
