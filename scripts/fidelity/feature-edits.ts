// The save evidence of each editable feature row (`docs/site/data/pptx-features.ts`, "Evidence
// rule"). A recipe opens a fixture of its row, makes the row's own edit through the public
// editor API, saves, and opens the saved copy again. The edit must be in the reopened copy, a
// second save must give the same package, and only the parts the edit owns may change.
//
// `feature-edits.test.ts` runs every recipe. `bun scripts/fidelity/feature-edits.ts --out <dir>`
// writes each edited copy as `<row id>.pptx`, so the copies can be opened in other
// presentation applications to check for a repair prompt (CONTRIBUTING.md, "Fidelity harness").

import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { parseArgs } from 'node:util';
import {
  attr,
  elements,
  find,
  presentationInfo,
  readPptx,
  textOf,
  xmlPartOf,
} from '../../packages/core/src/store/index.ts';
import {
  all,
  FIXTURES,
  firstRun,
  has,
  lacks,
  MEDIA,
  nameOf,
  NEW_SLIDES,
  PNG,
  PRESENTATION,
  runFeatureEdit,
  shapeElement,
  shapesIn,
  shapeXmlOf,
  SHAPE_KINDS,
  slidePart,
  type FeatureEdit,
} from './feature-edit-kit.ts';

/** The recipes, by row id. Rows without a recipe cannot claim a full round trip. */
export const FEATURE_EDITS: Readonly<Record<string, FeatureEdit>> = {
  'text.basic-formatting': {
    fixture: 'text-runs',
    edit: (t) => {
      t.selectText('Shape 2', 0, 4);
      t.run({
        type: 'setFontFormat',
        bold: true,
        italic: true,
        underline: 'double',
        strike: 'single',
      });
    },
    check: (d) => has(d.shapeXml(0, 'Shape 2'), 'b="1"', 'i="1"', 'u="dbl"', 'strike="sngStrike"'),
  },
  'text.font-family-size': {
    fixture: 'open-font-families',
    edit: (t) => {
      t.selectText('Shape 2', 0, 3);
      t.run({ type: 'setFontFamily', family: 'Lato' });
      t.run({ type: 'setFontSize', points: 31 });
    },
    check: (d) => has(d.shapeXml(0, 'Shape 2'), 'typeface="Lato"', 'sz="3100"'),
  },
  'text.color-highlight': {
    fixture: 'text-runs',
    edit: (t) => {
      t.selectText('Shape 2', 0, 4);
      t.run({ type: 'setTextColor', color: '#C00000' });
      t.run({ type: 'setHighlight', color: '#FFFF00' });
    },
    check: (d) =>
      has(firstRun(d.shapeXml(0, 'Shape 2')), 'val="C00000"', '<a:highlight>', 'val="FFFF00"'),
  },
  'text.caps-baseline': {
    fixture: 'font-effects',
    edit: (t) => {
      t.selectText('Shape 2', 0, 3);
      t.run({ type: 'toggleSuperscript' });
      t.run({ type: 'setFontFormat', caps: 'small' });
    },
    check: (d) => has(firstRun(d.shapeXml(0, 'Shape 2')), 'baseline="30000"', 'cap="small"'),
  },
  'text.character-spacing': {
    fixture: 'font-effects',
    edit: (t) => {
      t.selectText('Shape 2', 0, 6);
      t.run({ type: 'setCharacterSpacing', points: 3 });
    },
    check: (d) => has(firstRun(d.shapeXml(0, 'Shape 2')), 'spc="300"'),
  },
  'text.change-case': {
    fixture: 'font-effects',
    edit: (t) => {
      t.selectText('Shape 2', 0, 6);
      t.run({ type: 'changeCase', case: 'upper' });
    },
    check: (d) => {
      const text = textOf(shapeElement(d.pkg, 0, 'Shape 2'));
      return text.slice(0, 6) === text.slice(0, 6).toUpperCase()
        ? undefined
        : `text "${text.slice(0, 6)}"`;
    },
  },
  'text.input-method': {
    fixture: 'text-runs',
    edit: (t) => {
      t.selectText('Shape 2', 0, 0);
      t.type('Typed é 日本 \u{1F600} & <x> ');
    },
    check: (d) =>
      textOf(shapeElement(d.pkg, 0, 'Shape 2')).startsWith('Typed é 日本 \u{1F600} & <x> ')
        ? undefined
        : 'typed text is missing',
  },
  'text.complex-scripts': {
    fixture: 'rtl-text',
    edit: (t) => {
      t.selectText('Shape 2', 0, 0);
      t.type('שלום ');
    },
    check: (d) =>
      all(
        textOf(shapeElement(d.pkg, 0, 'Shape 2')).startsWith('שלום ')
          ? undefined
          : 'typed text is missing',
        has(d.shapeXml(0, 'Shape 2'), 'rtl="1"')
      ),
  },
  'paragraphs.alignment': {
    fixture: 'alignment-anchor',
    edit: (t) => {
      t.selectText('Shape 2', 0, 1);
      t.run({ type: 'setAlign', align: 'r' });
      t.run({ type: 'setVerticalAlign', anchor: 'b' });
    },
    check: (d) => has(d.shapeXml(0, 'Shape 2'), 'algn="r"', 'anchor="b"'),
  },
  'paragraphs.spacing': {
    fixture: 'spacing',
    edit: (t) => {
      t.selectText('Shape 2', 0, 1);
      t.run({ type: 'setLineSpacing', multiple: 1.5 });
      t.run({ type: 'setParagraphFormat', spaceBeforePt: 9, spaceAfterPt: 7 });
    },
    check: (d) =>
      has(
        d.shapeXml(0, 'Shape 2'),
        '<a:spcPct val="150000"/>',
        '<a:spcPts val="900"/>',
        '<a:spcPts val="700"/>'
      ),
  },
  'paragraphs.indent-tabs': {
    fixture: 'paragraph-format',
    edit: (t) => {
      t.selectText('Shape 2', 0, 1, 1);
      t.run({ type: 'indent', direction: 1 });
    },
    check: (d) => {
      const p = elements(
        elements(shapeElement(d.pkg, 0, 'Shape 2')).find((c) => c.local === 'txBody')!
      ).filter((c) => c.local === 'p')[1];
      const pPr = p && elements(p).find((c) => c.local === 'pPr');
      return attr(pPr, 'lvl') || attr(pPr, 'marL')
        ? undefined
        : 'the second paragraph has no indent';
    },
  },
  'paragraphs.autofit': {
    fixture: 'autofit',
    edit: (t) => {
      t.select('Shape 2');
      t.run({ type: 'setAutofit', mode: 'shape' });
      t.select('Shape 3');
      t.run({ type: 'setAutofit', mode: 'none' });
    },
    check: (d) =>
      all(
        has(d.shapeXml(0, 'Shape 2'), '<a:spAutoFit/>'),
        has(d.shapeXml(0, 'Shape 3'), '<a:noAutofit/>')
      ),
  },
  'paragraphs.text-direction-columns': {
    fixture: 'vertical-columns',
    edit: (t) => {
      t.select('Shape 2');
      t.run({ type: 'setTextDirection', direction: 'vertical270' });
      t.select('Shape 3');
      t.run({ type: 'setColumns', count: 3, gapEmu: 91440 });
    },
    check: (d) =>
      all(
        has(d.shapeXml(0, 'Shape 2'), 'vert="vert270"'),
        has(d.shapeXml(0, 'Shape 3'), 'numCol="3"', 'spcCol="91440"')
      ),
  },
  'lists.bullets-numbering': {
    fixture: 'bullets-numbering',
    edit: (t) => {
      t.selectText('Shape 2', 0, 1);
      t.run({ type: 'toggleNumbering' });
    },
    check: (d) => has(d.shapeXml(0, 'Shape 2'), '<a:buAutoNum'),
  },
  'shapes.preset-geometry': {
    fixture: 'preset-shapes',
    edit: (t) => {
      t.run({ type: 'insertShape', preset: 'star5' });
      t.select('Shape 2');
      const node = t.node('Shape 2');
      t.run({
        type: 'setTransform',
        nodeId: node,
        x: 914400,
        y: 914400,
        cx: 1828800,
        cy: 914400,
        rot: 0,
      });
    },
    check: (d) =>
      all(
        has(d.slideXml(0), 'prst="star5"'),
        has(
          d.shapeXml(0, 'Shape 2'),
          '<a:off x="914400" y="914400"/>',
          '<a:ext cx="1828800" cy="914400"/>'
        )
      ),
  },
  'shapes.groups': {
    fixture: 'groups-transforms',
    edit: (t) => {
      t.select('Shape 5', 'Shape 6');
      t.run({ type: 'group' });
    },
    check: (d) => {
      const root = xmlPartOf(d.pkg, slidePart(d.pkg, 0))!.root;
      const groups = shapesIn(root).filter((s) => s.local === 'grpSp');
      const grouped = groups.some((g) => {
        const names = shapesIn(g).map(nameOf);
        return names.includes('Shape 5') && names.includes('Shape 6');
      });
      return grouped ? undefined : 'Shape 5 and Shape 6 are not in one group';
    },
  },
  'shapes.connectors': {
    fixture: 'lines-arrows',
    edit: (t) => {
      t.select('Line 0');
      t.run({ type: 'setLineEnds', head: 'oval', tail: 'stealth' });
      t.run({ type: 'insertLine', kind: 'elbow' });
    },
    check: (d) =>
      all(
        has(d.shapeXml(0, 'Line 0'), '<a:headEnd type="oval"', '<a:tailEnd type="stealth"'),
        has(d.slideXml(0), 'prst="bentConnector3"')
      ),
  },
  'shapes.arrange': {
    fixture: 'arrange-flip-hidden',
    edit: (t) => {
      t.select('Arrow');
      t.run({ type: 'order', move: 'front' });
      t.select('Distributed 1', 'Distributed 2', 'Distributed 3');
      t.run({ type: 'alignObjects', align: 'top' });
    },
    check: (d) => {
      const tree = find(xmlPartOf(d.pkg, slidePart(d.pkg, 0))!.root, 'p:spTree');
      const top = elements(tree).filter((c) => SHAPE_KINDS.has(c.local));
      const last = nameOf(top[top.length - 1]);
      const ys = ['Distributed 1', 'Distributed 2', 'Distributed 3'].map((n) =>
        attr(find(shapeElement(d.pkg, 0, n), 'a:off'), 'y')
      );
      return all(
        last === 'Arrow' ? undefined : `the front shape is "${last}"`,
        new Set(ys).size === 1 ? undefined : `tops differ: ${ys.join(', ')}`
      );
    },
  },
  'shapes.flip-rotate': {
    fixture: 'arrange-flip-hidden',
    edit: (t) => {
      t.select('Arrow');
      t.run({ type: 'flipObjects', axis: 'horizontal' });
      t.run({ type: 'rotateObjects', degrees: 90 });
      t.run({ type: 'setLockAspectRatio', lock: true });
    },
    check: (d) => has(d.shapeXml(0, 'Arrow'), 'flipH="1"', 'rot="5400000"', 'noChangeAspect="1"'),
  },
  'shapes.selection-pane': {
    fixture: 'arrange-flip-hidden',
    edit: (t) => {
      t.run({ type: 'setShapesVisible', nodeIds: [t.node('Arrow')], visible: false });
      t.run({ type: 'setShapesVisible', nodeIds: [t.node('Hidden shape')], visible: true });
      t.run({ type: 'renameShape', nodeId: t.node('Distributed 1'), name: 'Renamed & <kept>' });
    },
    check: (d) =>
      all(
        has(d.shapeXml(0, 'Arrow'), 'hidden="1"'),
        lacks(d.shapeXml(0, 'Hidden shape'), 'hidden="1"'),
        has(d.shapeXml(0, 'Renamed & <kept>'), 'name="Renamed &amp; &lt;kept&gt;"')
      ),
  },
  'shapes.styles-gallery': {
    fixture: 'shape-styles',
    edit: (t) => {
      t.select('Intense effect dk1');
      t.run({ type: 'setShapeStyle', style: { row: 'coloredFill', color: 'accent4' } });
    },
    check: (d) => has(d.shapeXml(0, 'Intense effect dk1'), '<p:style>', 'val="accent4"'),
  },
  'fills-lines.solid-gradient': {
    fixture: 'fills-lines',
    edit: (t) => {
      t.select('Shape 2');
      t.run({ type: 'setShapeFill', color: '#00B050' });
      t.select('Shape 3');
      t.run({
        type: 'setShapeGradient',
        gradient: {
          kind: 'linear',
          angle: 45,
          rotWithShape: true,
          stops: [
            { pos: 0, color: '#1F4E79' },
            { pos: 0.5, color: '#FFFFFF' },
            { pos: 1, color: '#C55A11' },
          ],
        },
      });
    },
    check: (d) =>
      all(
        has(d.shapeXml(0, 'Shape 2'), '<a:solidFill><a:srgbClr val="00B050"/>'),
        has(
          d.shapeXml(0, 'Shape 3'),
          '<a:gradFill',
          'val="1F4E79"',
          'pos="50000"',
          'val="C55A11"',
          'ang="2700000"'
        )
      ),
  },
  'fills-lines.lines': {
    fixture: 'lines-arrows',
    edit: (t) => {
      t.select('Line 1');
      t.run({ type: 'setLineWidth', points: 4.5 });
      t.run({ type: 'setLineDash', dash: 'lgDashDot' });
      t.run({ type: 'setLineColor', color: '#7030A0' });
    },
    check: (d) => has(d.shapeXml(0, 'Line 1'), 'w="57150"', 'val="lgDashDot"', 'val="7030A0"'),
  },
  'fills-lines.theme-colors': {
    fixture: 'colors',
    edit: (t) => {
      t.select('Shape 2');
      t.run({ type: 'setShapeFill', color: { theme: 'accent2', shift: 40 } });
      t.select('Shape 3');
      t.run({ type: 'setLineColor', color: { theme: 'tx1', shift: -25 } });
    },
    check: (d) =>
      all(
        has(
          d.shapeXml(0, 'Shape 2'),
          '<a:schemeClr val="accent2"><a:lumMod val="60000"/><a:lumOff val="40000"/>'
        ),
        has(d.shapeXml(0, 'Shape 3'), '<a:schemeClr val="tx1"><a:lumMod val="75000"/>')
      ),
  },
  'effects.outer-shadow': {
    fixture: 'effects',
    edit: (t) => {
      t.select('Shape 7');
      t.run({ type: 'setShapeEffect', effect: { kind: 'shadow', preset: 'outerBottomLeft' } });
    },
    check: (d) => has(d.shapeXml(0, 'Shape 7'), '<a:outerShdw'),
  },
  'effects.glow-reflection': {
    fixture: 'effects',
    edit: (t) => {
      t.select('Shape 3');
      t.run({ type: 'setShapeEffect', effect: { kind: 'glow', size: 8, color: '#FFC000' } });
      t.run({ type: 'setShapeEffect', effect: { kind: 'reflection', preset: 'halfTouching' } });
      t.run({ type: 'setShapeEffect', effect: { kind: 'softEdge', size: 5 } });
    },
    check: (d) =>
      has(
        d.shapeXml(0, 'Shape 3'),
        '<a:glow rad="101600"',
        'val="FFC000"',
        '<a:reflection',
        '<a:softEdge rad="63500"'
      ),
  },
  'pictures.raster-svg': {
    fixture: 'pictures',
    edit: (t) => {
      t.select('Picture 0');
      t.run({ type: 'replaceImage', bytes: t.png, mime: 'image/png' });
      t.run({ type: 'insertImage', bytes: t.png, mime: 'image/png', width: 2, height: 2 });
    },
    check: (d) => {
      const pictures = shapesIn(xmlPartOf(d.pkg, slidePart(d.pkg, 0))!.root).filter(
        (s) => s.local === 'pic'
      );
      const media = [...d.pkg.parts.keys()].filter((p) => p.startsWith('/ppt/media/'));
      const same = media.filter((p) => {
        const part = d.pkg.parts.get(p);
        return (
          part?.kind === 'binary' &&
          part.bytes.length === PNG.length &&
          part.bytes.every((b, i) => b === PNG[i])
        );
      });
      return all(
        pictures.length === 7 ? undefined : `${pictures.length} pictures`,
        same.length >= 1 ? undefined : 'the new picture bytes are missing'
      );
    },
    parts: MEDIA,
  },
  'pictures.adjustments': {
    fixture: 'picture-editing',
    edit: (t) => {
      t.select('Picture 0');
      t.run({
        type: 'setImageAdjustments',
        adjustments: { brightness: 0.2, contrast: -0.1, transparency: 0.3, recolor: 'grayscale' },
      });
    },
    check: (d) => has(d.shapeXml(0, 'Picture 0'), '<a:lum', '<a:grayscl/>', '<a:alphaModFix'),
  },
  'pictures.crop': {
    fixture: 'picture-editing',
    edit: (t) => {
      t.select('Picture 1');
      t.run({ type: 'cropImage', crop: { kind: 'rect', left: 10, top: 5, right: 15, bottom: 0 } });
      t.select('Picture 2');
      t.run({ type: 'cropToShape', preset: 'ellipse' });
    },
    check: (d) =>
      all(
        has(d.shapeXml(0, 'Picture 1'), '<a:srcRect l="10000" t="5000" r="15000"'),
        has(d.shapeXml(0, 'Picture 2'), 'prst="ellipse"')
      ),
  },
  'pictures.styles-border': {
    fixture: 'picture-editing',
    edit: (t) => {
      t.select('Picture 3');
      t.run({ type: 'setImageStyle', style: 'simpleFrameBlack' });
    },
    check: (d) => has(d.shapeXml(0, 'Picture 3'), '<a:ln'),
  },
  'pictures.alt-text': {
    fixture: 'picture-editing',
    edit: (t) => {
      t.select('Picture 4');
      t.run({ type: 'setAltText', description: 'A chart & <its> "legend"', title: 'Sales' });
      t.select('Picture 5');
      t.run({ type: 'setAltText', decorative: true });
    },
    check: (d) =>
      all(
        has(
          d.shapeXml(0, 'Picture 4'),
          'descr="A chart &amp; &lt;its&gt; &quot;legend&quot;"',
          'title="Sales"'
        ),
        has(d.shapeXml(0, 'Picture 5'), 'decorative')
      ),
  },
  'tables.render': {
    fixture: 'tables',
    edit: (t) => {
      t.selectCell('Table 1', 0, 0);
      t.run({ type: 'toggleTableStyleOption', option: 'bandRow' });
      t.run({ type: 'toggleTableStyleOption', option: 'firstCol' });
    },
    check: (d) => {
      const tblPr = d.shapeXml(0, 'Table 1').match(/<a:tblPr[^>]*>/)?.[0] ?? '';
      return tblPr.includes('firstCol="1"') ? undefined : `table options ${tblPr}`;
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
      const tables = d.slideXml(0).match(/<a:tbl>/g)?.length ?? 0;
      const rows = d.shapeXml(0, 'Table 1').match(/<a:tr /g)?.length ?? 0;
      const original = readPptx(new Uint8Array(readFileSync(join(FIXTURES, 'table-editing.pptx'))));
      const before = original.ok
        ? (shapeXmlOf(original.pkg, 0, 'Table 1').match(/<a:tr /g)?.length ?? 0)
        : 0;
      return all(
        tables === 2 ? undefined : `${tables} tables`,
        rows === before + 1 ? undefined : `${rows} rows, was ${before}`
      );
    },
  },
  'tables.merge-split': {
    fixture: 'table-editing',
    edit: (t) => {
      t.editor.setSelection({
        kind: 'cells',
        nodeId: t.node('Table 1'),
        anchor: { row: 3, column: 0 },
        focus: { row: 3, column: 1 },
      });
      t.run({ type: 'mergeCells' });
    },
    check: (d) => {
      const rows = d.shapeXml(0, 'Table 1').split('<a:tr ');
      return has(rows[rows.length - 1], 'gridSpan="2"', 'hMerge="1"');
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
      const original = readPptx(new Uint8Array(readFileSync(join(FIXTURES, 'tables.pptx'))));
      const before = original.ok ? styleOf(shapeXmlOf(original.pkg, 0, 'Table 1')) : undefined;
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
      const n = presentationInfo(d.pkg).slides.length;
      return n === 5 ? undefined : `${n} slides`;
    },
    parts: [
      ...PRESENTATION,
      ...NEW_SLIDES,
      /^\/ppt\/media\//,
      /^\/ppt\/charts\//,
      /^\/ppt\/embeddings\//,
    ],
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
      t.run({ type: 'addSlide', ...(layout ? { layoutId: layout.id } : {}) });
    },
    check: (d) => {
      const xml = d.slideXml(1);
      return all(has(xml, '<p:ph'), /<a:t>/.test(xml) ? 'the new slide has text' : undefined);
    },
    parts: [...PRESENTATION, ...NEW_SLIDES],
  },
  'slides.notes': {
    fixture: 'slide-layouts',
    edit: (t) => {
      t.run({ type: 'setNotes', index: 0, text: 'Notes & <kept>\nSecond line' });
      t.run({ type: 'formatNotes', index: 0, start: 0, end: 5, format: 'bold', value: true });
    },
    check: (d) => {
      const notes = [...d.pkg.parts.keys()].filter((p) =>
        /^\/ppt\/notesSlides\/[^/]+\.xml$/.test(p)
      );
      const xml = notes.map((p) => d.partXml(p)).find((x) => x.includes('&lt;kept&gt;'));
      if (!xml) return 'no notes slide has the text';
      const text = textOf(
        xmlPartOf(
          d.pkg,
          notes.find((p) => d.partXml(p) === xml)
        )!.root
      );
      return all(
        text.includes('Notes & <kept>') && text.includes('Second line')
          ? undefined
          : `notes text "${text}"`,
        has(xml, 'b="1"/><a:t>Notes</a:t>')
      );
    },
    parts: [
      /^\/ppt\/notesSlides\//,
      /^\/ppt\/notesMasters\//,
      /^\/ppt\/theme\//,
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
      const s = t.editor.snapshot();
      const current = s.slides[1]?.layoutId;
      const other = s.layouts?.find((l) => l.id !== current);
      if (!other) throw new Error('no other layout');
      t.run({ type: 'applyLayout', layoutId: other.id, indices: [1] });
    },
    check: (d) => {
      const ref = presentationInfo(d.pkg).slides[1];
      const original = readPptx(new Uint8Array(readFileSync(join(FIXTURES, 'slide-layouts.pptx'))));
      const before = original.ok ? presentationInfo(original.pkg).slides[1]?.layoutPart : undefined;
      return ref?.layoutPart && ref.layoutPart !== before ? undefined : `layout ${ref?.layoutPart}`;
    },
  },
};

if (import.meta.main) {
  const { values } = parseArgs({ options: { out: { type: 'string' } } });
  if (values.out) mkdirSync(values.out, { recursive: true });
  let failed = 0;
  for (const [id, recipe] of Object.entries(FEATURE_EDITS)) {
    const result = await runFeatureEdit(id, recipe);
    if (!result.ok) failed++;
    console.log(`${id}: ${result.ok ? 'ok' : `FAIL ${result.failure}`}`);
    if (values.out && result.bytes) writeFileSync(join(values.out, `${id}.pptx`), result.bytes);
  }
  process.exitCode = failed ? 1 : 0;
}
