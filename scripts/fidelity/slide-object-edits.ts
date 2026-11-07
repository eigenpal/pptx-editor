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
  textOf,
  xmlPartOf,
  type PptxPackage,
} from '../../packages/core/src/store/index.ts';
import { PptxEditor } from '../../packages/editor-api/src/browser.ts';
import { openPeers } from './collaboration-kit.ts';
import { bytesText, CHART_DATA_PARTS, editChart, workbookText } from './chart-data-edits.ts';
import { REVIEW_COMMENTS_EDIT } from './comment-edits.ts';
import {
  caretInDiagram,
  consistent,
  diagramCommand,
  DIAGRAM_PARTS,
  topTexts,
} from './diagram-edits.ts';
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
    fixture: 'equations',
    slide: 1,
    edit: (t) => {
      // Insert Equation from a linear form, then Edit Equation on a fraction of the deck.
      t.run({ type: 'insertEquation', linear: 'A=\\pi r^2' });
      t.select('Equation 3');
      const state = t.editor.snapshot().equation;
      if (!state) throw new Error('no equation is selected');
      t.run({ type: 'setEquation', nodeId: state.nodeId, linear: 'y=(a+1)/b' });
    },
    check: (d) => {
      // The Markup Compatibility wrapper of the equation named `name`.
      const wrapper = (name: string) =>
        d
          .slideXml(1)
          .split('<mc:AlternateContent')
          .find((piece) => piece.includes(`name="${name}"`)) ?? '';
      const inserted = wrapper('Equation 9');
      const edited = wrapper('Equation 3');
      return all(
        has(
          inserted,
          'Requires="a14"',
          '<m:t>π</m:t>',
          '<m:sSup><m:e>',
          'typeface="Cambria Math"',
          '<mc:Fallback>',
          '<a:t>A=πr^2</a:t>'
        ),
        has(edited, '<m:f><m:num>', '<m:t>1</m:t>', '<m:den>'),
        // The fallback showed the old equation, so the edit removes it.
        lacks(edited, '<mc:Fallback>'),
        kept(d, 1, 'Equation 4', 'Equation 5', 'Equation 6', 'Equation 7', 'Equation 8')
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
      // An effect of each class, one on a shape in markup-compatibility content, a reorder, new
      // timing, the removal of an effect of the file, and a new transition.
      t.select('Title');
      t.run({ type: 'addAnimation', preset: 'flyIn' });
      t.select('Warped text');
      t.run({ type: 'addAnimation', preset: 'pulse', trigger: 'afterPrevious' });
      t.select('EMF picture');
      t.run({ type: 'addAnimation', preset: 'fadeOut' });
      t.select('Equation');
      t.run({ type: 'addAnimation', preset: 'zoom', trigger: 'withPrevious' });
      const effects = () => t.editor.snapshot().animations!.effects;
      const of = (name: string) => effects().find((e) => e.shapeName === name)!.id;
      t.run({ type: 'moveAnimation', effectId: of('EMF picture'), to: 0 });
      t.run({ type: 'setAnimationTiming', effectId: of('Title'), delayMs: 250, durationMs: 1000 });
      t.run({ type: 'removeAnimation', effectId: of('3-D shape') });
      t.run({ type: 'setTransition', kind: 'push', durationMs: 1200, advanceAfterMs: 5000 });
    },
    check: (d) => {
      const animations = d.editor.snapshot().animations!;
      const list = animations.effects.map((e) =>
        [e.shapeName, e.preset, e.trigger, e.step, e.delayMs, e.durationMs].join(' ')
      );
      const want = [
        'EMF picture fadeOut onClick 1 0 500',
        'Animated text appear onClick 2 0 ',
        'Title flyIn onClick 3 250 1000',
        'Warped text pulse afterPrevious 3 0 500',
        'Equation zoom withPrevious 3 0 500',
      ];
      const xml = d.slideXml(0);
      const ids = [...xml.matchAll(/<p:cTn id="(\d+)"/g)].map((m) => m[1]);
      const transition = animations.transition;
      return all(
        list.join('\n') === want.join('\n') ? undefined : `effects ${list.join('; ')}`,
        transition.kind === 'push' &&
          transition.durationMs === 1200 &&
          transition.advanceOnClick &&
          transition.advanceAfterMs === 5000
          ? undefined
          : `transition ${JSON.stringify(transition)}`,
        new Set(ids).size === ids.length ? undefined : 'two time nodes share an id',
        lacks(xml, '<p:spTgt spid="4"/>', '<p:bldP spid="4"'),
        has(xml, '<p:bldP spid="3" grpId="0" animBg="1"/>', '<p:bldP spid="5" grpId="0"'),
        keptChild(d, 0, 'p:extLst')
      );
    },
  },
  // The chart's own edit: values, a renamed series, a series and a category added and removed.
  // `chart-data-edits.ts` makes each edit alone.
  'objects.charts': {
    fixture: 'chart-data',
    parts: CHART_DATA_PARTS,
    edit: (t) =>
      editChart(t, (d) => ({
        kind: d.kind,
        categories: [d.categories[0]!, d.categories[2]!, d.categories[3]!, { label: 'Central' }],
        series: [
          { ...d.series[0]!, name: 'Budget', values: [40, 50, 30, 20] },
          { ...d.series[1]!, values: [38, 47, 33, null] },
          { name: 'Stretch', values: [5, 6, 7, 8] },
        ],
      })),
    check: (d) => {
      const chart = bytesText(d.pkg, '/ppt/charts/chart1.xml');
      const sheet = workbookText(d.pkg, 'Workbook1.xlsx', '/xl/worksheets/sheet1.xml');
      const table = workbookText(d.pkg, 'Workbook1.xlsx', '/xl/tables/table1.xml');
      return all(
        has(
          chart,
          '<c:f>Sheet1!$A$2:$A$5</c:f><c:strCache><c:ptCount val="4"/>',
          '<c:pt idx="3"><c:v>Central</c:v></c:pt>',
          '<c:pt idx="0"><c:v>Budget</c:v></c:pt>',
          '<c:f>Sheet1!$D$1</c:f>',
          '<c:f>Sheet1!$D$2:$D$5</c:f>',
          '<c:dPt><c:idx val="2"/>'
        ),
        lacks(chart, 'Forecast', 'South'),
        has(
          sheet,
          '<c r="B1" t="inlineStr"><is><t>Budget</t></is></c>',
          '<c r="A5" t="inlineStr"><is><t>Central</t></is></c>',
          '<c r="D5"><v>8</v></c>'
        ),
        lacks(sheet, 'r="C5"'),
        has(table, 'ref="A1:D5"', '<tableColumn id="4" name="Stretch"/>')
      );
    },
  },
  // The diagram's own edits: text typed in a shape, a shape added, and a shape deleted. Each
  // writes the data model and the saved drawing; `diagram-edits.ts` makes them one at a time.
  'objects.diagrams': {
    fixture: 'diagram-editing',
    parts: DIAGRAM_PARTS,
    edit: (t) => {
      caretInDiagram(t, 'Build');
      t.type(' & review');
      caretInDiagram(t, 'Test');
      diagramCommand(t, { type: 'addDiagramShape', position: 'after' }, 'Launch');
      caretInDiagram(t, 'Plan');
      diagramCommand(t, { type: 'deleteDiagramShape' });
    },
    check: (d) =>
      all(
        topTexts(d, 1).join('|') === 'Build &amp; review|Test|Launch|Ship'
          ? undefined
          : `the top-level nodes are ${topTexts(d, 1).join('|')}`,
        has(bytesText(d.pkg, '/ppt/diagrams/drawing1.xml'), 'Build &amp; review', 'Launch'),
        lacks(bytesText(d.pkg, '/ppt/diagrams/data1.xml'), '<a:t>Plan</a:t>'),
        consistent(d, 1)
      ),
  },
  'review.comments': REVIEW_COMMENTS_EDIT,
  'slides.header-footer': {
    fixture: 'header-footer',
    slide: 1,
    // Apply to All: the footer of slide 1 takes the new text.
    parts: [/^\/ppt\/slides\/slide1\.xml$/],
    edit: (t) => {
      t.run({ type: 'setSlideNumbers', show: true });
      t.run({ type: 'setFooter', show: true, text: 'Quarterly results' });
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
