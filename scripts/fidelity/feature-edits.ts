// The save evidence of each editable feature row (`docs/site/data/pptx-features.ts`, "Evidence
// rule"). A recipe opens a fixture of its row, makes the row's own edit through the public
// editor API, saves, and opens the saved copy again. The edit must be in the reopened copy, a
// second save must give the same package, and only the parts the edit owns may change.
// The recipes of the table, slide, theme, collaboration, and kept-content rows are in
// `slide-object-edits.ts`.
//
// `feature-edits.test.ts` runs every recipe. `bun scripts/fidelity/feature-edits.ts --out <dir>`
// writes each edited copy as `<row id>.pptx`, so the copies can be opened in other
// presentation applications to check for a repair prompt (CONTRIBUTING.md, "Fidelity harness").

import { mkdirSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { parseArgs } from 'node:util';
import {
  attr,
  elements,
  find,
  imagePartFor,
  textOf,
  xmlPartOf,
} from '../../packages/core/src/store/index.ts';
import {
  all,
  firstRun,
  has,
  kept,
  keptRels,
  lacks,
  runsOf,
  MEDIA,
  nameOf,
  paragraphOf,
  PNG,
  runFeatureEdit,
  shapeElement,
  shapesIn,
  shapeXmlOf,
  SHAPE_KINDS,
  slidePart,
  NEW_SLIDES,
  type FeatureEdit,
} from './feature-edit-kit.ts';
import { SLIDE_AND_OBJECT_EDITS } from './slide-object-edits.ts';
import {
  paragraphSpecs,
  richFromHtml,
  richFromText,
} from '../../packages/core/src/editor/clipboard-rich.ts';

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
        underlineColor: '#00B0F0',
        strike: 'single',
      });
    },
    check: (d) =>
      has(
        firstRun(d.shapeXml(0, 'Shape 2')),
        'b="1"',
        'i="1"',
        'u="dbl"',
        'strike="sngStrike"',
        '<a:uFill><a:solidFill><a:srgbClr val="00B0F0"/>'
      ),
  },
  'text.font-family-size': {
    fixture: 'open-font-families',
    edit: (t) => {
      t.selectText('Shape 2', 0, 3);
      t.run({ type: 'setFontFamily', family: 'Lato' });
      t.run({ type: 'setFontSize', points: 31 });
    },
    check: (d) => has(firstRun(d.shapeXml(0, 'Shape 2')), 'typeface="Lato"', 'sz="3100"'),
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
  'text.hyperlinks': {
    fixture: 'hyperlinks',
    edit: (t) => {
      // Paragraph 0 is "Plain text to link. More plain text."
      t.selectText('Links', 0, 5);
      t.run({ type: 'setLink', url: 'https://example.com/new?a=1&b=2' });
      t.selectText('Links', 6, 10);
      t.run({ type: 'setLink', url: 'mailto:a@example.com?subject=Hi%20there' });
      t.selectText('Links', 20, 24);
      t.run({ type: 'setLink', url: '#slide-2' });
      t.selectText('Links', 0, 16, 7);
      t.run({ type: 'setLink', url: null });
      // "plain" links to the custom show "Short", and "to" to a file next to the deck.
      t.selectText('Links', 25, 30);
      t.run({ type: 'setLink', url: '#custom-show-0' });
      t.selectText('Links', 11, 13);
      t.run({ type: 'setLink', url: './docs/Q3 report.pdf' });
    },
    check: (d) => {
      const xml = d.shapeXml(0, 'Links');
      const rels = d.partXml('/ppt/slides/_rels/slide1.xml.rels');
      const target = (label: string): string | undefined => {
        const rPr = runsOf(xml).find((r) => r.text === label)?.rPr ?? '';
        const id = /r:id="([^"]+)"/.exec(rPr)?.[1];
        const rel = id && new RegExp(`<Relationship [^>]*Id="${id}"[^>]*>`).exec(rels)?.[0];
        return rel ? /Target="([^"]*)"/.exec(rel)?.[1] : undefined;
      };
      const slideJump = runsOf(xml).find((r) => r.text === 'More')?.rPr ?? '';
      const paragraphs = (x: string) => x.match(/<a:p>.*?<\/a:p>/g) ?? [];
      const before = paragraphs(shapeXmlOf(d.original, 0, 'Links'));
      const after = paragraphs(xml);
      return all(
        target('Plain') === 'https://example.com/new?a=1&amp;b=2'
          ? undefined
          : `the web link targets ${target('Plain')}`,
        target('text') === 'mailto:a@example.com?subject=Hi%20there'
          ? undefined
          : `the email link targets ${target('text')}`,
        has(slideJump, 'action="ppaction://hlinksldjump"'),
        has(
          runsOf(xml).find((r) => r.text === 'plain')?.rPr ?? '',
          '<a:hlinkClick r:id="" action="ppaction://customshow?id=0"/>'
        ),
        target('to') === 'docs/Q3%20report.pdf'
          ? undefined
          : `the file link targets ${target('to')}`,
        has(rels, 'Target="docs/Q3%20report.pdf" TargetMode="External"'),
        target('More') === 'slide2.xml' ? undefined : `the slide link targets ${target('More')}`,
        lacks(after[7] ?? '', 'hlinkClick'),
        has(after[7] ?? '', '<a:t>A link to remove</a:t>'),
        // The links the edit did not touch stay as they were: web, email, slide, the show jumps,
        // the custom show, the file, the program, and the shape link, with their targets.
        [1, 2, 3, 4, 5, 6, 8].every((i) => after[i] === before[i])
          ? undefined
          : 'a link the edit did not touch changed',
        kept(d, 0, 'Linked shape'),
        keptRels(d, 0, 'rIdWeb', 'rIdMail', 'rIdSlide', 'rIdFile', 'rIdShape', 'rIdProgram')
      );
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
  'text.find-replace': {
    fixture: 'slide-layouts',
    slide: 1,
    edit: (t) => {
      t.run({ type: 'replaceAll', query: 'Results', replacement: 'Outcomes & <totals>' });
      t.run({ type: 'replaceAll', query: 'headline', replacement: 'lead' });
    },
    check: (d) => {
      const notes = d.partXml('/ppt/notesSlides/notesSlide2.xml');
      return all(
        has(d.slideXml(1), 'Outcomes &amp; &lt;totals&gt;'),
        lacks(d.slideXml(1), '>Results'),
        // The replacement keeps the bold of the text it replaces.
        /b="1"[^>]*\/?>(?:<\/a:rPr>)?<a:t>lead number<\/a:t>/.test(notes)
          ? undefined
          : 'the notes replacement lost its bold'
      );
    },
    parts: [/^\/ppt\/notesSlides\/notesSlide2\.xml$/],
  },
  'text.clipboard': {
    fixture: 'text-runs',
    edit: (t) => {
      // HTML from another application at a caret: a bold red run, a bulleted paragraph, and a
      // last line that joins the text after the caret.
      const html =
        '<p><b style="color:#C00000;font-size:20pt">Pasted</b> text</p>' +
        '<ul><li>Second <i>item</i></li></ul><p>Last line</p>';
      t.selectText('Shape 2', 0, 0, 1);
      t.run({ type: 'pasteText', paragraphs: paragraphSpecs(richFromHtml(html)!) });
      // Plain text with a shape selected goes in a new text box.
      t.select('Shape 2');
      t.run({ type: 'pasteText', paragraphs: paragraphSpecs(richFromText('Box line')) });
      // Copy a shape and paste it; copy the slide and paste it after itself.
      t.select('Shape 2');
      const payload = t.editor.copy();
      if (!payload) throw new Error('nothing to copy');
      t.run({ type: 'paste', payload });
      t.run({ type: 'pasteSlides', file: t.editor.save()!, slides: [0] });
    },
    check: (d) => {
      const slide = d.slideXml(0);
      return all(
        has(slide, '<a:t>Pasted</a:t>'),
        has(slide, 'sz="2000" b="1"'),
        has(slide, '<a:buChar char="•"/>'),
        has(slide, '<a:t>Box line</a:t>'),
        has(d.slideXml(1), '<a:t>Box line</a:t>')
      );
    },
    parts: [...NEW_SLIDES],
  },
  'text.format-painter': {
    fixture: 'text-runs',
    edit: (t) => {
      // "Regular 18 Bold Italic …": copy the format of "Bold", paint "Regular".
      t.selectText('Shape 2', 11, 15);
      t.run({ type: 'copyFormat' });
      t.selectText('Shape 2', 0, 7);
      t.run({ type: 'pasteFormat' });
      // The Format Painter button: copy "Italic", paint "18".
      t.selectText('Shape 2', 16, 22);
      t.run({ type: 'toggleFormatPainter' });
      t.selectText('Shape 2', 8, 10);
      t.run({ type: 'pasteFormat' });
    },
    check: (d) => {
      const runs = runsOf(d.shapeXml(0, 'Shape 2'));
      const rPr = (text: string) => runs.find((r) => r.text === text)?.rPr ?? '';
      return all(has(rPr('Regular'), 'b="1"'), has(rPr('18'), 'i="1"'), lacks(rPr('18'), 'b="1"'));
    },
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
        paragraphOf(d.shapeXml(0, 'Shape 2'), 0),
        '<a:lnSpc><a:spcPct val="150000"/></a:lnSpc>',
        '<a:spcBef><a:spcPts val="900"/></a:spcBef>',
        '<a:spcAft><a:spcPts val="700"/></a:spcAft>'
      ),
  },
  'paragraphs.indent-tabs': {
    fixture: 'svg-tab-stops',
    edit: (t) => {
      // The Paragraph dialog indents of a paragraph with tab stops, and Increase List Level.
      t.selectText('Tab stops', 0, 1, 0);
      t.run({ type: 'setParagraphFormat', indentLeftEmu: 228600, indentFirstLineEmu: 114300 });
      t.selectText('Tab stops', 0, 1, 1);
      t.run({ type: 'indent', direction: 1 });
      t.selectText('Tab stops', 0, 1, 3);
      t.run({ type: 'setParagraphFormat', indentLeftEmu: 685800, indentFirstLineEmu: -342900 });
    },
    check: (d) => {
      const xml = d.shapeXml(0, 'Tab stops');
      const tabs = (x: string) => /<a:tabLst>.*?<\/a:tabLst>/.exec(x)?.[0];
      const committed = tabs(shapeXmlOf(d.original, 0, 'Tab stops'));
      return all(
        has(paragraphOf(xml, 0), 'marL="228600"', 'indent="114300"'),
        has(paragraphOf(xml, 1), 'lvl="1"'),
        has(paragraphOf(xml, 3), 'marL="685800"', 'indent="-342900"'),
        // The tab stops, which the editor does not edit, stay on every paragraph that has them.
        [0, 1, 2].every((i) => tabs(paragraphOf(xml, i)) === committed)
          ? undefined
          : 'a tab stop list changed'
      );
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
  'paragraphs.text-warp': {
    fixture: 'text-warp',
    edit: (t) => {
      t.select('textArchDown 3');
      t.run({ type: 'setTextWarp', preset: 'textWave1', adjust: { adj1: 15000 } });
      t.select('textArchUp 2');
      t.run({ type: 'setTextWarp', preset: null });
    },
    check: (d) =>
      all(
        has(
          d.shapeXml(0, 'textArchDown 3'),
          '<a:bodyPr wrap="none" rtlCol="0" anchor="t"><a:prstTxWarp prst="textWave1"><a:avLst><a:gd name="adj1" fmla="val 15000"/></a:avLst></a:prstTxWarp><a:noAutofit/></a:bodyPr>'
        ),
        has(
          d.shapeXml(0, 'textArchUp 2'),
          '<a:bodyPr wrap="none" rtlCol="0" anchor="t"><a:noAutofit/>'
        ),
        lacks(d.shapeXml(0, 'textArchUp 2'), 'prstTxWarp')
      ),
  },
  'lists.bullets-numbering': {
    fixture: 'bullets-numbering',
    edit: (t) => {
      t.selectText('Shape 2', 0, 1);
      t.run({ type: 'toggleNumbering' });
      // A picture bullet on "Second level": one media part, related to the slide.
      t.selectText('Shape 3', 0, 6, 1);
      t.run({
        type: 'setBulletFormat',
        format: { kind: 'picture', picture: t.png, sizePercent: 120 },
      });
    },
    check: (d) => {
      const second = paragraphOf(d.shapeXml(0, 'Shape 3'), 1);
      const rId = /<a:buBlip><a:blip r:embed="([^"]+)"\/><\/a:buBlip>/.exec(second)?.[1];
      const rels = d.partXml('/ppt/slides/_rels/slide1.xml.rels');
      const rel = rId && new RegExp(`<Relationship [^>]*Id="${rId}"[^>]*>`).exec(rels)?.[0];
      const media = rel ? /Target="\.\.\/media\/([^"]+)"/.exec(rel)?.[1] : undefined;
      return all(
        has(d.shapeXml(0, 'Shape 2'), '<a:buAutoNum'),
        has(second, '<a:buSzPct val="120000"/>', '<a:buBlip>'),
        media && d.pkg.parts.get(`/ppt/media/${media}`)?.kind === 'binary'
          ? undefined
          : 'the picture bullet names no media part',
        d.editor
          .session!.layout(0)!
          .items.some(
            (i) => i.kind === 'shape' && !!i.text?.lines.some((l) => l.runs.some((r) => r.picture))
          )
          ? undefined
          : 'the reopened copy draws no picture bullet'
      );
    },
    parts: MEDIA,
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
  'shapes.custom-geometry': {
    fixture: 'custom-geometry',
    edit: (t) => {
      t.run({
        type: 'setTransform',
        nodeId: t.node('Freeform'),
        x: 1371600,
        y: 685800,
        cx: 5486400,
        cy: 2743200,
        rot: 30,
      });
      t.select('Freeform');
      t.run({ type: 'setShapeFill', color: '#70AD47' });
      t.run({ type: 'flipObjects', axis: 'vertical' });
    },
    check: (d) => {
      const geometry = (xml: string) => /<a:custGeom>.*?<\/a:custGeom>/.exec(xml)?.[0];
      const xml = d.shapeXml(0, 'Freeform');
      return all(
        has(xml, '<a:off x="1371600" y="685800"/>', '<a:ext cx="5486400" cy="2743200"/>'),
        // A vertical flip turns the 30° rotation into 330°.
        has(xml, 'rot="19800000"', 'flipV="1"', 'val="70AD47"'),
        geometry(xml) === geometry(shapeXmlOf(d.original, 0, 'Freeform'))
          ? undefined
          : 'the custom geometry changed'
      );
    },
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
  'fills-lines.pattern-picture': {
    fixture: 'fills-lines',
    edit: (t) => {
      t.select('Shape 2');
      t.run({
        type: 'setShapePattern',
        pattern: { preset: 'dkDnDiag', foreground: '#1F4E79', background: '#DEEBF7' },
      });
      t.select('Shape 3');
      t.run({ type: 'setShapePicture', bytes: t.png, mime: 'image/png' });
      t.select('Shape 4');
      t.run({ type: 'setShapePicture', bytes: t.png, mime: 'image/png', tile: true });
    },
    check: (d) =>
      all(
        has(
          d.shapeXml(0, 'Shape 2'),
          '<a:pattFill prst="dkDnDiag"><a:fgClr><a:srgbClr val="1F4E79"/>',
          '<a:bgClr><a:srgbClr val="DEEBF7"/>'
        ),
        has(d.shapeXml(0, 'Shape 3'), '<a:blipFill', '<a:stretch>'),
        has(d.shapeXml(0, 'Shape 4'), '<a:blipFill', '<a:tile'),
        // The picture and pattern fills the edit did not touch stay as they were.
        kept(d, 0, 'Shape 10', 'Shape 30', 'Shape 31')
      ),
    parts: MEDIA,
  },
  'effects.fill-overlay': {
    fixture: 'fill-fidelity',
    slide: 2,
    edit: (t) => {
      t.run({
        type: 'setTransform',
        nodeId: t.node('Shape 5'),
        x: 457200,
        y: 457200,
        cx: 2286000,
        cy: 1371600,
        rot: 0,
      });
      t.select('Shape 6');
      t.run({ type: 'setLineColor', color: '#C00000' });
      t.run({ type: 'setLineWidth', points: 3 });
    },
    check: (d) => {
      const overlay = (xml: string) => /<a:fillOverlay.*?<\/a:fillOverlay>/.exec(xml)?.[0];
      const same = (name: string) =>
        overlay(d.shapeXml(2, name)) === overlay(shapeXmlOf(d.original, 2, name))
          ? undefined
          : `the overlay of ${name} changed`;
      return all(
        has(d.shapeXml(2, 'Shape 5'), '<a:off x="457200" y="457200"/>'),
        has(d.shapeXml(2, 'Shape 6'), 'val="C00000"', 'w="38100"'),
        same('Shape 5'),
        same('Shape 6'),
        kept(d, 2, 'Shape 7')
      );
    },
  },
  'effects.3d-and-other': {
    fixture: 'preserved-content',
    edit: (t) => {
      t.select('3-D shape');
      t.run({ type: 'setShapeFill', color: '#5B9BD5' });
      t.run({ type: 'nudge', dx: 91440, dy: 0 });
    },
    check: (d) => {
      const threeD = (xml: string) => /<a:scene3d>.*<\/a:sp3d>/.exec(xml)?.[0];
      const xml = d.shapeXml(0, '3-D shape');
      return all(
        has(xml, 'val="5B9BD5"'),
        threeD(xml) === threeD(shapeXmlOf(d.original, 0, '3-D shape'))
          ? undefined
          : 'the 3-D scene or shape changed'
      );
    },
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
      t.select('Shape 7');
      t.run({ type: 'setShapeEffect', effect: { kind: 'shadow', preset: 'innerTopLeft' } });
    },
    check: (d) =>
      all(
        has(
          d.shapeXml(0, 'Shape 3'),
          '<a:glow rad="101600"',
          'val="FFC000"',
          '<a:reflection',
          '<a:softEdge rad="63500"'
        ),
        has(d.shapeXml(0, 'Shape 7'), '<a:innerShdw')
      ),
  },
  'pictures.raster-svg': {
    fixture: 'svg-tab-stops',
    edit: (t) => {
      t.select('PNG picture');
      t.run({ type: 'replaceImage', bytes: t.png, mime: 'image/png' });
      t.run({ type: 'insertImage', bytes: t.png, mime: 'image/png', width: 2, height: 2 });
      // Move, resize, and rotate the SVG picture.
      t.run({
        type: 'setTransform',
        nodeId: t.node('SVG picture'),
        x: 7315200,
        y: 457200,
        cx: 1828800,
        cy: 1828800,
        rot: 15,
      });
    },
    check: (d) => {
      const pictures = shapesIn(xmlPartOf(d.pkg, slidePart(d.pkg, 0))!.root).filter(
        (s) => s.local === 'pic'
      );
      const embed = /r:embed="([^"]+)"/.exec(d.shapeXml(0, 'PNG picture'))?.[1];
      const target = embed && imagePartFor(d.pkg, slidePart(d.pkg, 0), embed);
      const part = target ? d.pkg.parts.get(target) : undefined;
      const replaced =
        part?.kind === 'binary' &&
        part.bytes.length === PNG.length &&
        part.bytes.every((b, i) => b === PNG[i]);
      const fill = (xml: string) => /<p:blipFill>.*?<\/p:blipFill>/.exec(xml)?.[0];
      const svg = d.shapeXml(0, 'SVG picture');
      return all(
        pictures.length === 3 ? undefined : `${pictures.length} pictures`,
        replaced ? undefined : `the PNG picture shows ${target}, not the new picture`,
        has(svg, 'rot="900000"', '<a:off x="7315200" y="457200"/>'),
        // The SVG picture keeps its SVG part and its PNG fallback, which it shared with the PNG
        // picture: the new picture takes a relationship of its own.
        fill(svg) === fill(shapeXmlOf(d.original, 0, 'SVG picture'))
          ? undefined
          : 'the SVG picture fill changed',
        embed !== 'rIdPng' ? undefined : 'the new picture reuses the shared relationship',
        keptRels(d, 0, 'rIdPng', 'rIdSvg')
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
    check: (d) =>
      has(
        d.shapeXml(0, 'Picture 3'),
        '<a:ln w="88900" cap="sq"><a:solidFill><a:srgbClr val="000000"/>',
        '<a:outerShdw blurRad="55000"'
      ),
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
  ...SLIDE_AND_OBJECT_EDITS,
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
