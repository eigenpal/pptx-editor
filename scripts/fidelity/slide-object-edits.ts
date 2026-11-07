// The feature edits of the table, slide, theme, and collaboration rows, and of the kept-content
// rows of text, pictures, objects, presenting, and review. They join the other recipes in
// `FEATURE_EDITS` (`feature-edits.ts`), which says what a recipe proves. The kept-content rows
// of shapes and effects stay with their category there. A recipe of a kept-content row edits
// next to the content, or its box, and its check finds the content and its relationships as
// they were.

import {
  canonicalPackageFingerprint,
  find,
  presentationInfo,
  readPptx,
  serializeXmlPart,
  textOf,
  xmlPartOf,
  type PptxPackage,
} from '../../packages/core/src/store/index.ts';
import { PptxEditor } from '../../packages/editor-api/src/browser.ts';
import { openPeers } from './collaboration-kit.ts';
import {
  added,
  all,
  has,
  kept,
  keptChild,
  keptRels,
  lacks,
  nameOf,
  NEW_SLIDES,
  paragraphOf,
  PRESENTATION,
  removed,
  runsOf,
  shapeElement,
  shapesIn,
  shapeXmlOf,
  slidePart,
  type FeatureEdit,
} from './feature-edit-kit.ts';

/** The recipes of this file, by row id. */
export const SLIDE_AND_OBJECT_EDITS: Readonly<Record<string, FeatureEdit>> = {
  'tables.render': {
    fixture: 'tables',
    edit: (t) => {
      t.selectCell('Table 1', 0, 0);
      t.run({ type: 'toggleTableStyleOption', option: 'bandRow' });
      t.run({ type: 'toggleTableStyleOption', option: 'firstCol' });
    },
    check: (d) => {
      const options = (xml: string) => /<a:tblPr[^>]*>/.exec(xml)?.[0] ?? '';
      const before = options(shapeXmlOf(d.original, 0, 'Table 1'));
      const after = options(d.shapeXml(0, 'Table 1'));
      const flag = (tag: string, name: string) => tag.includes(`${name}="1"`);
      return flag(after, 'bandRow') !== flag(before, 'bandRow') &&
        flag(after, 'firstCol') !== flag(before, 'firstCol')
        ? undefined
        : `table options ${after}, were ${before}`;
    },
  },
  'tables.cell-text': {
    fixture: 'table-editing',
    edit: (t) => {
      t.selectCell('Table 1', 1, 1);
      t.type('Cell & <text> ');
    },
    check: (d) => has(d.shapeXml(0, 'Table 1'), 'Cell &amp; &lt;text&gt; '),
  },
  'tables.structure': {
    fixture: 'table-editing',
    edit: (t) => {
      t.selectCell('Table 1', 0, 0);
      t.run({ type: 'insertRow', where: 'below' });
      t.run({ type: 'insertColumn', where: 'right' });
      t.run({ type: 'insertTable', rows: 2, cols: 3 });
    },
    check: (d) => {
      const count = (xml: string, re: RegExp) => xml.match(re)?.length ?? 0;
      const before = shapeXmlOf(d.original, 0, 'Table 1');
      const after = d.shapeXml(0, 'Table 1');
      return all(
        count(d.slideXml(0), /<a:tbl>/g) === 2 ? undefined : 'no new table',
        count(after, /<a:tr /g) === count(before, /<a:tr /g) + 1 ? undefined : 'no new row',
        count(after, /<a:gridCol /g) === count(before, /<a:gridCol /g) + 1
          ? undefined
          : 'no new column'
      );
    },
  },
  'tables.merge-split': {
    fixture: 'table-editing',
    edit: (t) => {
      const table = t.node('Table 1');
      t.editor.setSelection({
        kind: 'cells',
        nodeId: table,
        anchor: { row: 3, column: 0 },
        focus: { row: 3, column: 1 },
      });
      t.run({ type: 'mergeCells' });
      // The fixture merges rows 1 and 2 of columns 0 and 1: split that block into four cells.
      t.selectCell('Table 1', 1, 0);
      t.run({ type: 'splitCell', rows: 2, cols: 2 });
    },
    check: (d) => {
      const rows = d.shapeXml(0, 'Table 1').split('<a:tr ');
      return all(
        has(rows[rows.length - 1], 'gridSpan="2"', 'hMerge="1"'),
        lacks(rows[2] + rows[3], 'gridSpan=', 'rowSpan=', 'hMerge=', 'vMerge=')
      );
    },
  },
  'tables.styles': {
    fixture: 'tables',
    edit: (t) => {
      t.selectCell('Table 1', 0, 0);
      const current = t.editor.snapshot().table?.styleId;
      const style = t.editor.tableStyles().find((s) => s.id !== current);
      if (!style) throw new Error('no other table style');
      t.run({ type: 'setTableStyle', styleId: style.id });
    },
    check: (d) => {
      const styleOf = (xml: string) => /<a:tableStyleId>([^<]*)</.exec(xml)?.[1];
      const before = styleOf(shapeXmlOf(d.original, 0, 'Table 1'));
      const after = styleOf(d.shapeXml(0, 'Table 1'));
      return after && after !== before ? undefined : `table style ${after}, was ${before}`;
    },
    parts: [/^\/ppt\/tableStyles\.xml$/],
  },
  'tables.cell-format': {
    fixture: 'table-editing',
    edit: (t) => {
      t.selectCell('Table 1', 0, 1);
      t.run({ type: 'setCellFill', color: '#92D050' });
      t.run({ type: 'setCellMargins', margins: { left: 182880, top: 45720 } });
      t.run({
        type: 'setTableBorders',
        scope: 'outside',
        spec: { color: '#FF0000', width: 2, dash: 'solid' },
      });
    },
    check: (d) =>
      has(
        d.shapeXml(0, 'Table 1'),
        'val="92D050"',
        'marL="182880"',
        'marT="45720"',
        'val="FF0000"'
      ),
  },
  'tables.sizes': {
    fixture: 'table-editing',
    edit: (t) => {
      t.selectCell('Table 1', 0, 0);
      t.run({ type: 'setRowHeight', height: 640080 });
      t.run({ type: 'setColumnWidth', width: 2011680 });
    },
    check: (d) => has(d.shapeXml(0, 'Table 1'), 'h="640080"', 'w="2011680"'),
  },
  'slides.structure': {
    fixture: 'slide-structure',
    edit: (t) => {
      t.run({ type: 'duplicateSlide', index: 1 });
      t.run({ type: 'moveSlide', from: 0, to: 3 });
      t.run({ type: 'deleteSlide', index: 4 });
      t.run({ type: 'addSlide' });
    },
    check: (d) => {
      const titleOf = (pkg: PptxPackage, i: number) =>
        textOf(find(xmlPartOf(pkg, slidePart(pkg, i))!.root, 'p:spTree')).match(
          /^(Intro|Body) \d/
        )?.[0] ?? '';
      const titles = presentationInfo(d.pkg).slides.map((_, i) => titleOf(d.pkg, i));
      const kept = titles.filter(Boolean).join(', ');
      const expected = 'Intro 2, Intro 2, Body 1, Intro 1';
      // The duplicate keeps every shape of its source.
      const shapes = (i: number) =>
        shapesIn(xmlPartOf(d.pkg, slidePart(d.pkg, i))!.root)
          .map(nameOf)
          .join(', ');
      return all(
        titles.length === 5 ? undefined : `${titles.length} slides`,
        kept === expected ? undefined : `slide order ${kept}, expected ${expected}`,
        shapes(0) === shapes(1) ? undefined : `the duplicate has ${shapes(1)}`
      );
    },
    parts: [...NEW_SLIDES, removed(/^\/ppt\/(slides|notesSlides|comments|embeddings)\//)],
  },
  'slides.hidden': {
    fixture: 'slide-layouts',
    edit: (t) => t.run({ type: 'setSlidesHidden', indices: [1, 3], hidden: true }),
    check: (d) =>
      all(
        has(d.slideXml(1), 'show="0"'),
        has(d.slideXml(3), 'show="0"'),
        lacks(d.slideXml(0), 'show="0"')
      ),
    parts: [/^\/ppt\/slides\/slide[24]\.xml$/],
  },
  'slides.placeholder-prompts': {
    fixture: 'slide-layouts',
    edit: (t) => {
      const layout = t.editor.snapshot().layouts?.find((l) => /content/i.test(l.name));
      if (!layout) throw new Error('the fixture has no content layout');
      t.run({ type: 'addSlide', layoutId: layout.id });
    },
    check: (d) => {
      const xml = d.slideXml(1);
      return all(has(xml, '<p:ph'), /<a:t>/.test(xml) ? 'the new slide has text' : undefined);
    },
    parts: NEW_SLIDES,
  },
  'slides.notes': {
    fixture: 'slide-layouts',
    edit: (t) => {
      t.run({ type: 'setNotes', index: 0, text: 'Notes & <kept>\nSecond line' });
      t.run({ type: 'formatNotes', index: 0, start: 0, end: 5, format: 'bold', value: true });
    },
    check: (d) => {
      const rels = d.partXml(slidePart(d.pkg, 0).replace(/\/([^/]+)$/, '/_rels/$1.rels'));
      const target = /Target="\.\.\/notesSlides\/([^"]+)"/.exec(rels)?.[1];
      if (!target) return 'slide 1 has no notes slide';
      const part = `/ppt/notesSlides/${target}`;
      const text = textOf(xmlPartOf(d.pkg, part)!.root);
      return all(
        text.includes('Notes & <kept>') && text.includes('Second line')
          ? undefined
          : `notes text "${text}"`,
        has(d.partXml(part), 'b="1"/><a:t>Notes</a:t>')
      );
    },
    parts: [
      added(/^\/ppt\/notesSlides\//),
      added(/^\/ppt\/notesMasters\//),
      added(/^\/ppt\/theme\//),
      /^\/\[Content_Types\]\.xml$/,
      ...PRESENTATION,
    ],
  },
  'slides.sections': {
    fixture: 'slide-layouts',
    edit: (t) => {
      t.run({ type: 'addSection', name: 'Intro', index: 0 });
      t.run({ type: 'addSection', name: 'Body & <rest>', index: 2 });
    },
    check: (d) =>
      has(
        d.partXml('/ppt/presentation.xml'),
        'sectionLst',
        'name="Intro"',
        'name="Body &amp; &lt;rest&gt;"'
      ),
    parts: PRESENTATION,
  },
  'slides.size': {
    fixture: 'slide-size',
    // A4 to 16:9: the height ratio (6858000 / 7562850) is the smaller one, so every box
    // scales by it and the content is centered across the width.
    edit: (t) => t.run({ type: 'setSlideSize', width: 12_192_000, height: 6_858_000 }),
    check: (d) => {
      const s = 6_858_000 / 7_562_850;
      const dx = Math.round((12_192_000 - 10_693_400 * s) / 2);
      return all(
        has(d.partXml('/ppt/presentation.xml'), '<p:sldSz cx="12192000" cy="6858000"/>'),
        has(
          d.shapeXml(0, 'Shape 2'),
          `<a:off x="${dx}" y="${Math.round(7_242_810 * s)}"/>`,
          `<a:ext cx="${Math.round(10_693_400 * s)}" cy="${Math.round(320_040 * s)}"/>`
        )
      );
    },
    parts: [
      /^\/ppt\/presentation\.xml$/,
      /^\/ppt\/slides\/slide\d+\.xml$/,
      /^\/ppt\/slideLayouts\/slideLayout\d+\.xml$/,
      /^\/ppt\/slideMasters\/slideMaster\d+\.xml$/,
    ],
  },
  'themes.backgrounds': {
    fixture: 'placeholders-background',
    edit: (t) =>
      t.run({
        type: 'setSlideBackground',
        background: { kind: 'solid', color: { theme: 'accent6', shift: 60 } },
      }),
    check: (d) => has(d.slideXml(0), '<p:bg><p:bgPr><a:solidFill><a:schemeClr val="accent6">'),
  },
  'themes.change-layout-theme': {
    fixture: 'slide-layouts',
    slide: 1,
    edit: (t) => {
      const other = t.editor.snapshot().layouts?.find((l) => l.name === 'Title Only');
      if (!other) throw new Error('the fixture has no Title Only layout');
      t.run({ type: 'applyLayout', layoutId: other.id, indices: [1] });
      // Move the title away from its layout box, then Reset puts it back.
      t.run({
        type: 'setTransform',
        nodeId: t.node('Title 1'),
        x: 914400,
        y: 3657600,
        cx: 4572000,
        cy: 914400,
        rot: 0,
      });
      t.run({ type: 'resetSlides', indices: [1] });
    },
    check: (d) => {
      const after = presentationInfo(d.pkg).slides[1]?.layoutPart;
      const before = presentationInfo(d.original).slides[1]?.layoutPart;
      const text = (pkg: PptxPackage) =>
        textOf(find(xmlPartOf(pkg, slidePart(pkg, 1))!.root, 'p:spTree'));
      return all(
        after && after !== before ? undefined : `layout ${after}, was ${before}`,
        lacks(d.shapeXml(1, 'Title 1'), '<a:off'),
        text(d.pkg) === text(d.original) ? undefined : 'the slide text changed'
      );
    },
  },
  'themes.placeholder-inheritance': {
    fixture: 'placeholders-background',
    edit: (t) => {
      t.selectText('Title 1', 0, 0);
      t.type('New: ');
      t.selectText('Content 2', 0, 0, 1);
      t.type('Typed ');
    },
    check: (d) => {
      const title = d.shapeXml(0, 'Title 1');
      const body = d.shapeXml(0, 'Content 2');
      return all(
        textOf(shapeElement(d.pkg, 0, 'Title 1')).startsWith('New: ') ? undefined : 'no title text',
        has(paragraphOf(body, 1), '<a:t>Typed '),
        // Typing writes no box, fill, or format of its own: the placeholders still inherit them.
        has(title, '<p:spPr/>'),
        has(body, '<p:spPr/>'),
        lacks(title + body, '<a:xfrm', 'sz="', 'typeface=')
      );
    },
  },
  'collaboration.realtime': {
    fixture: 'text-runs',
    collaborate: openPeers,
    edit: (t) => {
      const peer = t.peer!;
      // Concurrent edits of one paragraph: one peer types at its start, the other bolds its
      // first word and types at its end. The other peer also moves Shape 3.
      t.link!.pause();
      t.selectText('Shape 2', 0, 0);
      t.type('Peer one: ');
      peer.selectText('Shape 2', 0, 7);
      peer.run({ type: 'toggleItalic' });
      peer.run({
        type: 'setTransform',
        nodeId: peer.node('Shape 3'),
        x: 6400800,
        y: 914400,
        cx: 4572000,
        cy: 3657600,
        rot: 0,
      });
      t.link!.flush();
      // Both replicas converge before the save.
      const own = readPptx(t.editor.save()!);
      const other = readPptx(peer.editor.save()!);
      if (!own.ok || !other.ok) throw new Error('a replica did not save');
      if (canonicalPackageFingerprint(own.pkg) !== canonicalPackageFingerprint(other.pkg))
        throw new Error('the replicas differ');
    },
    check: (d) => {
      const runs = runsOf(d.shapeXml(0, 'Shape 2'));
      const typed = runs.find((r) => r.text.includes('Peer one: '));
      const italic = runs.find((r) => r.text.includes('Regular'));
      return all(
        typed ? undefined : 'the typed text is gone',
        italic && /i="1"/.test(italic.rPr) ? undefined : 'the italic of the other peer is gone',
        has(d.shapeXml(0, 'Shape 3'), '<a:off x="6400800" y="914400"/>')
      );
    },
  },
  'collaboration.automation-api': {
    fixture: 'text-runs',
    edit: async (t) => {
      const runtime = PptxEditor.createBrowser(t.editor);
      try {
        await runtime.run(async (context) => {
          const shapes = context.presentation.slides.getItemAt(0).shapes;
          const box = shapes.addTextBox('Added by code & <kept>', {
            left: 72,
            top: 400,
            width: 360,
            height: 60,
          });
          box.name = 'Automation box';
          box.textFrame.textRange.font.bold = true;
          box.fill.setSolidColor('#FFF2CC');
          await context.sync();
        });
      } finally {
        runtime.dispose();
      }
    },
    check: (d) =>
      has(
        d.shapeXml(0, 'Automation box'),
        'Added by code &amp; &lt;kept&gt;',
        'b="1"',
        'val="FFF2CC"'
      ),
  },
  // --- Kept content: an edit next to it, or of its box, keeps it as it was ---------------------
  'paragraphs.text-warp': {
    fixture: 'preserved-content',
    edit: (t) => {
      t.selectText('Warped text', 0, 0);
      t.type('New ');
      t.run({
        type: 'setTransform',
        nodeId: t.node('Warped text'),
        x: 457200,
        y: 1188720,
        cx: 5029200,
        cy: 1463040,
        rot: 0,
      });
    },
    check: (d) => {
      const warp = (xml: string) => /<a:prstTxWarp.*?<\/a:prstTxWarp>/.exec(xml)?.[0];
      const xml = d.shapeXml(0, 'Warped text');
      return all(
        textOf(shapeElement(d.pkg, 0, 'Warped text')).startsWith('New ') ? undefined : 'no text',
        has(xml, '<a:off x="457200" y="1188720"/>'),
        warp(xml) === warp(shapeXmlOf(d.original, 0, 'Warped text'))
          ? undefined
          : 'the warp changed'
      );
    },
  },
  'pictures.emf-wmf': {
    fixture: 'preserved-content',
    edit: (t) => {
      t.run({
        type: 'setTransform',
        nodeId: t.node('EMF picture'),
        x: 5029200,
        y: 3200400,
        cx: 2743200,
        cy: 1371600,
        rot: 0,
      });
      t.select('WMF picture');
      t.run({ type: 'nudge', dx: 0, dy: 91440 });
    },
    check: (d) => {
      const fill = (xml: string) => /<p:blipFill>.*?<\/p:blipFill>/.exec(xml)?.[0];
      const same = (name: string) =>
        fill(d.shapeXml(0, name)) === fill(shapeXmlOf(d.original, 0, name))
          ? undefined
          : `the fill of ${name} changed`;
      return all(
        has(
          d.shapeXml(0, 'EMF picture'),
          '<a:off x="5029200" y="3200400"/>',
          '<a:ext cx="2743200" cy="1371600"/>'
        ),
        lacks(d.shapeXml(0, 'WMF picture'), shapeXmlOf(d.original, 0, 'WMF picture')),
        same('EMF picture'),
        same('WMF picture'),
        keptRels(d, 0, 'rIdEmf', 'rIdWmf')
      );
    },
  },
  'objects.equations': {
    fixture: 'preserved-content',
    edit: (t) => {
      t.selectText('Animated text', 0, 0);
      t.type('Edited. ');
    },
    check: (d) => {
      const equation = (pkg: PptxPackage) =>
        /<mc:AlternateContent.*?<\/mc:AlternateContent>/.exec(
          serializeXmlPart(xmlPartOf(pkg, slidePart(pkg, 0))!)
        )?.[0];
      return all(
        textOf(find(shapeElement(d.pkg, 0, 'Animated text'), 'p:txBody')).startsWith('Edited. ')
          ? undefined
          : 'no text',
        equation(d.pkg) === equation(d.original) ? undefined : 'the equation changed',
        // The picture that the equation's fallback shows.
        keptRels(d, 0, 'rIdImg')
      );
    },
  },
  'objects.media-ole': {
    fixture: 'preserved-content',
    edit: (t) => {
      t.run({
        type: 'setTransform',
        nodeId: t.node('Audio clip'),
        x: 7589520,
        y: 3017520,
        cx: 548640,
        cy: 411480,
        rot: 0,
      });
      t.select('Video clip', 'Embedded workbook');
      t.run({ type: 'nudge', dx: 0, dy: 182880 });
    },
    check: (d) => {
      const media = (xml: string) => /<p:nvPicPr>.*?<\/p:blipFill>/.exec(xml)?.[0];
      const same = (name: string) =>
        media(d.shapeXml(0, name)) === media(shapeXmlOf(d.original, 0, name))
          ? undefined
          : `the link, media, or poster of ${name} changed`;
      const ole = (xml: string) => /<a:graphic>.*<\/a:graphic>/.exec(xml)?.[0];
      const workbook = d.shapeXml(0, 'Embedded workbook');
      return all(
        has(d.shapeXml(0, 'Audio clip'), '<a:off x="7589520" y="3017520"/>'),
        lacks(d.shapeXml(0, 'Video clip'), shapeXmlOf(d.original, 0, 'Video clip')),
        lacks(workbook, shapeXmlOf(d.original, 0, 'Embedded workbook')),
        same('Audio clip'),
        same('Video clip'),
        ole(workbook) === ole(shapeXmlOf(d.original, 0, 'Embedded workbook'))
          ? undefined
          : 'the embedded workbook changed',
        keptChild(d, 0, 'p:timing'),
        keptRels(d, 0, 'rIdAudio', 'rIdMedia', 'rIdVideo', 'rIdVideoMedia', 'rIdOle', 'rIdImg')
      );
    },
  },
  'presenting.animations-transitions': {
    fixture: 'preserved-content',
    edit: (t) => {
      t.selectText('Animated text', 0, 4);
      t.run({ type: 'toggleBold' });
      t.selectText('Animated text', 0, 0);
      t.type('Typed. ');
    },
    check: (d) => {
      // The shape keeps its id, which the animation names, and its unknown extension.
      const nonVisual = (xml: string) => /<p:cNvPr .*?<\/p:cNvPr>/.exec(xml)?.[0];
      return all(
        textOf(find(shapeElement(d.pkg, 0, 'Animated text'), 'p:txBody')).startsWith('Typed. ')
          ? undefined
          : 'no text',
        keptChild(d, 0, 'p:transition'),
        keptChild(d, 0, 'p:timing'),
        keptChild(d, 0, 'p:extLst'),
        nonVisual(d.shapeXml(0, 'Animated text')) ===
          nonVisual(shapeXmlOf(d.original, 0, 'Animated text'))
          ? undefined
          : 'the id or the extension of the animated shape changed'
      );
    },
  },
  'objects.charts': {
    fixture: 'charts',
    edit: (t) => {
      t.selectText('Shape 2', 0, 0);
      t.type('Edited: ');
      t.run({
        type: 'setTransform',
        nodeId: t.node('Chart 1'),
        x: 914400,
        y: 1371600,
        cx: 6400800,
        cy: 4114800,
        rot: 0,
      });
    },
    check: (d) => {
      const data = (xml: string) => /<a:graphic>.*<\/a:graphic>/.exec(xml)?.[0];
      const xml = d.shapeXml(0, 'Chart 1');
      return all(
        has(xml, '<a:off x="914400" y="1371600"/>', '<a:ext cx="6400800" cy="4114800"/>'),
        data(xml) === data(shapeXmlOf(d.original, 0, 'Chart 1')) ? undefined : 'the chart changed'
      );
    },
  },
  'objects.diagrams': {
    fixture: 'diagrams',
    edit: (t) => {
      t.selectText('Shape 2', 0, 0);
      t.type('Edited: ');
      t.run({
        type: 'setTransform',
        nodeId: t.node('Diagram 1'),
        x: 1828800,
        y: 1371600,
        cx: 6400800,
        cy: 4114800,
        rot: 0,
      });
    },
    check: (d) => {
      const data = (xml: string) => /<a:graphic>.*<\/a:graphic>/.exec(xml)?.[0];
      const xml = d.shapeXml(0, 'Diagram 1');
      return all(
        has(xml, '<a:off x="1828800" y="1371600"/>'),
        data(xml) === data(shapeXmlOf(d.original, 0, 'Diagram 1'))
          ? undefined
          : 'the diagram changed'
      );
    },
  },
  'review.comments': {
    fixture: 'slide-structure',
    edit: (t) => {
      t.selectText('Title', 0, 0);
      t.type('Edited: ');
    },
    check: (d) =>
      all(
        textOf(shapeElement(d.pkg, 0, 'Title')).startsWith('Edited: ') ? undefined : 'no text',
        has(d.partXml('/ppt/slides/_rels/slide1.xml.rels'), 'Target="../comments/comment1.xml"')
      ),
  },
  'slides.header-footer': {
    fixture: 'header-footer',
    slide: 1,
    // Apply to All: the footer of slide 1 takes the new text.
    parts: [/^\/ppt\/slides\/slide1\.xml$/],
    edit: (t) => {
      t.run({ type: 'setSlideNumbers', show: true });
      t.run({ type: 'setFooter', text: 'Quarterly results' });
      t.run({ type: 'setDateTime', show: true, text: '11/7/2026' });
    },
    check: (d) => {
      const footer = d.shapeXml(0, 'Footer Placeholder 3');
      return all(
        has(d.slideXml(1), 'type="sldNum"', 'type="slidenum"', 'type="ftr"', 'type="dt"'),
        has(d.slideXml(1), 'Quarterly results', 'type="datetime1"', '11/7/2026'),
        // Apply to All: slide 1 keeps its number, and its footer keeps its box and format.
        kept(d, 0, 'Slide Number Placeholder 4'),
        has(footer, 'Quarterly results', '<a:pPr algn="ctr"/>', 'sz="1200"'),
        lacks(footer, 'Footer text'),
        /<p:spPr>.*?<\/p:spPr>/.exec(footer)?.[0] ===
          /<p:spPr>.*?<\/p:spPr>/.exec(shapeXmlOf(d.original, 0, 'Footer Placeholder 3'))?.[0]
          ? undefined
          : 'the footer box changed'
      );
    },
  },
  'slides.sorter': {
    fixture: 'slide-layouts',
    // Slides A to F; sections Opening (A, B) and Details (C to F). A is "Quarterly review", B is
    // "Agenda", C is "Before and after", and D is "Backup slide".
    edit: (t) => {
      const details = t.editor.snapshot().sections?.find((x) => x.name === 'Details');
      if (!details) throw new Error('the fixture has no Details section');
      t.run({ type: 'moveSlides', indices: [0], to: 3 }); // B C D A E F
      t.run({ type: 'duplicateSlides', indices: [1] }); // B C C' D A E F
      t.run({ type: 'deleteSlides', indices: [6] }); // B C C' D A E
      t.run({ type: 'setSlidesHidden', indices: [3], hidden: true }); // D hidden
      t.run({ type: 'moveSection', sectionId: details.id, direction: -1 }); // C C' D A E B
    },
    check: (d) => {
      const slides = presentationInfo(d.pkg).slides;
      const text = (i: number) =>
        textOf(find(xmlPartOf(d.pkg, slides[i]!.partName)!.root, 'p:spTree'));
      const starts = ['Before and after', 'Before and after', 'Backup slide', 'Quarterly review'];
      const presentation = d.partXml('/ppt/presentation.xml');
      const sections = [
        ...presentation.matchAll(/<p14:section name="([^"]*)".*?<\/p14:section>/g),
      ].map((m) => `${m[1]} ${m[0].match(/<p14:sldId /g)?.length ?? 0}`);
      return all(
        slides.length === 6 ? undefined : `${slides.length} slides`,
        starts.every((s, i) => text(i).startsWith(s)) && text(5).startsWith('Agenda')
          ? undefined
          : `slide order ${slides.map((_, i) => text(i).slice(0, 16)).join(' | ')}`,
        slides[0]!.partName !== slides[1]!.partName ? undefined : 'the duplicate is the same part',
        slides.every((_, i) => (i === 2) === /<p:sld [^>]*show="0"/.test(d.slideXml(i)))
          ? undefined
          : 'a slide other than D is hidden, or D is not',
        sections.join(', ') === 'Details 5, Opening 1'
          ? undefined
          : `sections ${sections.join(', ')}`
      );
    },
    parts: [
      ...NEW_SLIDES,
      removed(/^\/ppt\/(slides|notesSlides)\//),
      /^\/ppt\/slides\/slide\d+\.xml$/,
    ],
  },
};
