// The diagram editing fixture (ECMA-376 Part 1, 21.4): a process of boxes joined by arrows, and
// a vertical list of boxes, for the diagram text and structure edits. The layout, quick style,
// and colors definitions are written here, with `urn:pptx-editor.dev/diagram/` ids. The data
// and drawing parts in `diagrams/` (`data-process.xml`, `drawing-process.xml`, ...) are what the
// presentation app saves after it lays each diagram out with these definitions, as for the
// `diagrams` fixture (`drawingml.ts`). `bun scripts/fixtures/diagram-editing.ts --seed <dir>`
// writes the deck from the data written here, with no saved drawing, for that app step.

import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { parseArgs } from 'node:util';
import { buildPptx, type FixtureOptions, type FixtureSlide } from './builder.ts';
import { DIAGRAM_PARTS, frame, title } from './drawingml.ts';

const DGM = 'http://schemas.openxmlformats.org/drawingml/2006/diagram';
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const IN = 914_400;
const DIR = join(import.meta.dirname, 'diagrams');

const margins = ['l', 'r', 't', 'b']
  .map((m) => `<dgm:constr type="${m}Marg" refType="primFontSz" fact="0.3"/>`)
  .join('');

/** A text box layout node with one shape. */
const textNode = (name: string, styleLbl: string, shapeType: string, adj: string): string =>
  `<dgm:layoutNode name="${name}" styleLbl="${styleLbl}"><dgm:varLst><dgm:bulletEnabled val="0"/></dgm:varLst><dgm:alg type="tx"/>` +
  `<dgm:shape type="${shapeType}"><dgm:adjLst>${adj}</dgm:adjLst></dgm:shape><dgm:presOf axis="self"/>` +
  `<dgm:constrLst>${margins}</dgm:constrLst><dgm:ruleLst><dgm:rule type="primFontSz" val="5" fact="NaN" max="NaN"/></dgm:ruleLst></dgm:layoutNode>`;

const layoutDef = (id: string, category: string, body: string): string =>
  `${XML}<dgm:layoutDef xmlns:dgm="${DGM}" xmlns:a="${A}" uniqueId="urn:pptx-editor.dev/diagram/layout/${id}"><dgm:title val=""/><dgm:desc val=""/>` +
  `<dgm:catLst><dgm:cat type="${category}" pri="1000"/></dgm:catLst>${body}</dgm:layoutDef>`;

/** Boxes from left to right, an arrow after each box but the last. */
const PROCESS_LAYOUT = layoutDef(
  'arrow-process',
  'process',
  `<dgm:layoutNode name="diagram"><dgm:varLst><dgm:dir/><dgm:resizeHandles val="exact"/></dgm:varLst>` +
    `<dgm:alg type="lin"><dgm:param type="linDir" val="fromL"/><dgm:param type="nodeVertAlign" val="mid"/></dgm:alg>` +
    `<dgm:shape><dgm:adjLst/></dgm:shape><dgm:presOf/><dgm:constrLst>` +
    `<dgm:constr type="w" for="ch" forName="node" refType="w"/>` +
    `<dgm:constr type="h" for="ch" forName="node" refType="w" refFor="ch" refForName="node" fact="0.6"/>` +
    `<dgm:constr type="w" for="ch" forName="arrow" refType="w" refFor="ch" refForName="node" fact="0.3"/>` +
    `<dgm:constr type="h" for="ch" forName="arrow" refType="h" refFor="ch" refForName="node" fact="0.35"/>` +
    `<dgm:constr type="w" for="ch" forName="space" refType="w" refFor="ch" refForName="node" fact="0.1"/>` +
    `<dgm:constr type="primFontSz" for="ch" forName="node" op="equ" val="20"/></dgm:constrLst><dgm:ruleLst/>` +
    `<dgm:forEach name="nodes" axis="ch" ptType="node">` +
    textNode('node', 'node1', 'roundRect', '<dgm:adj idx="1" val="0.1"/>') +
    `<dgm:forEach name="arrows" axis="followSib" ptType="sibTrans" cnt="1">` +
    `<dgm:layoutNode name="spaceBefore"><dgm:alg type="sp"/><dgm:shape><dgm:adjLst/></dgm:shape><dgm:presOf/><dgm:constrLst/><dgm:ruleLst/></dgm:layoutNode>` +
    `<dgm:layoutNode name="arrow" styleLbl="sibTrans2D1"><dgm:alg type="sp"/><dgm:shape type="rightArrow"><dgm:adjLst><dgm:adj idx="1" val="0.6"/><dgm:adj idx="2" val="0.5"/></dgm:adjLst></dgm:shape><dgm:presOf axis="self"/><dgm:constrLst/><dgm:ruleLst/></dgm:layoutNode>` +
    `<dgm:layoutNode name="space"><dgm:alg type="sp"/><dgm:shape><dgm:adjLst/></dgm:shape><dgm:presOf/><dgm:constrLst/><dgm:ruleLst/></dgm:layoutNode>` +
    `</dgm:forEach></dgm:forEach></dgm:layoutNode>`
);

/** Boxes from top to bottom, with a gap between each two. */
const LIST_LAYOUT = layoutDef(
  'vertical-boxes',
  'list',
  `<dgm:layoutNode name="diagram"><dgm:varLst><dgm:dir/><dgm:resizeHandles val="exact"/></dgm:varLst>` +
    `<dgm:alg type="lin"><dgm:param type="linDir" val="fromT"/></dgm:alg>` +
    `<dgm:shape><dgm:adjLst/></dgm:shape><dgm:presOf/><dgm:constrLst>` +
    `<dgm:constr type="w" for="ch" forName="item" refType="w"/>` +
    `<dgm:constr type="h" for="ch" forName="item" refType="h"/>` +
    `<dgm:constr type="h" for="ch" forName="gap" refType="h" refFor="ch" refForName="item" fact="0.15"/>` +
    `<dgm:constr type="primFontSz" for="ch" forName="item" op="equ" val="18"/></dgm:constrLst><dgm:ruleLst/>` +
    `<dgm:forEach name="items" axis="ch" ptType="node">` +
    textNode('item', 'node1', 'rect', '') +
    `<dgm:forEach name="gaps" axis="followSib" ptType="sibTrans" cnt="1">` +
    `<dgm:layoutNode name="gap"><dgm:alg type="sp"/><dgm:shape><dgm:adjLst/></dgm:shape><dgm:presOf/><dgm:constrLst/><dgm:ruleLst/></dgm:layoutNode>` +
    `</dgm:forEach></dgm:forEach></dgm:layoutNode>`
);

const styleRef = (ln: number, font: string): string =>
  `<dgm:style><a:lnRef idx="${ln}"><a:scrgbClr r="0" g="0" b="0"/></a:lnRef><a:fillRef idx="1"><a:scrgbClr r="0" g="0" b="0"/></a:fillRef>` +
  `<a:effectRef idx="0"><a:scrgbClr r="0" g="0" b="0"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="${font}"/></a:fontRef></dgm:style>`;

const QUICK_STYLE =
  `${XML}<dgm:styleDef xmlns:dgm="${DGM}" xmlns:a="${A}" uniqueId="urn:pptx-editor.dev/diagram/style/flat"><dgm:title val=""/><dgm:desc val=""/>` +
  `<dgm:catLst><dgm:cat type="simple" pri="1000"/></dgm:catLst>` +
  `<dgm:styleLbl name="node1">${styleRef(2, 'lt1')}</dgm:styleLbl>` +
  `<dgm:styleLbl name="sibTrans2D1">${styleRef(0, 'lt1')}</dgm:styleLbl></dgm:styleDef>`;

const colorLbl = (name: string, fill: string, line: string, text: string): string =>
  `<dgm:styleLbl name="${name}"><dgm:fillClrLst meth="repeat">${fill}</dgm:fillClrLst><dgm:linClrLst meth="repeat">${line}</dgm:linClrLst>` +
  `<dgm:effectClrLst/><dgm:txLinClrLst/><dgm:txFillClrLst meth="repeat">${text}</dgm:txFillClrLst><dgm:txEffectClrLst/></dgm:styleLbl>`;
const scheme = (v: string) => `<a:schemeClr val="${v}"/>`;

const COLORS = (fill: string) =>
  `${XML}<dgm:colorsDef xmlns:dgm="${DGM}" xmlns:a="${A}" uniqueId="urn:pptx-editor.dev/diagram/colors/${fill}"><dgm:title val=""/><dgm:desc val=""/>` +
  `<dgm:catLst><dgm:cat type="accent1" pri="1000"/></dgm:catLst>` +
  colorLbl('node1', scheme(fill), scheme('lt1'), scheme('lt1')) +
  colorLbl('sibTrans2D1', `<a:schemeClr val="${fill}"><a:tint val="60000"/></a:schemeClr>`, scheme(fill), scheme('lt1')) +
  `</dgm:colorsDef>`;

/** A data model of top-level nodes with these texts, as an author writes it: no layout yet. */
function seedData(n: number, layout: string, colors: string, texts: readonly string[]): string {
  let next = 1;
  const id = () =>
    `{0000000${n}-0000-4000-8000-${(next++).toString(16).toUpperCase().padStart(12, '0')}}`;
  const body = (text?: string) =>
    `<dgm:t><a:bodyPr/><a:lstStyle/><a:p>${text ? `<a:r><a:rPr lang="en-US"/><a:t>${text}</a:t></a:r>` : '<a:endParaRPr lang="en-US"/>'}</a:p></dgm:t>`;
  const doc = id();
  const pts = [
    `<dgm:pt modelId="${doc}" type="doc"><dgm:prSet loTypeId="urn:pptx-editor.dev/diagram/layout/${layout}" loCatId="list" qsTypeId="urn:pptx-editor.dev/diagram/style/flat#1" qsCatId="simple" csTypeId="urn:pptx-editor.dev/diagram/colors/${colors}#1" csCatId="accent1"/><dgm:spPr/>${body()}</dgm:pt>`,
  ];
  const cxns: string[] = [];
  texts.forEach((text, i) => {
    const [node, par, sib, cxn] = [id(), id(), id(), id()];
    pts.push(
      `<dgm:pt modelId="${node}"><dgm:prSet/><dgm:spPr/>${body(text)}</dgm:pt>`,
      `<dgm:pt modelId="${par}" type="parTrans" cxnId="${cxn}"><dgm:prSet/><dgm:spPr/>${body()}</dgm:pt>`,
      `<dgm:pt modelId="${sib}" type="sibTrans" cxnId="${cxn}"><dgm:prSet/><dgm:spPr/>${body()}</dgm:pt>`
    );
    cxns.push(
      `<dgm:cxn modelId="${cxn}" srcId="${doc}" destId="${node}" srcOrd="${i}" destOrd="0" parTransId="${par}" sibTransId="${sib}"/>`
    );
  });
  return (
    `${XML}<dgm:dataModel xmlns:dgm="${DGM}" xmlns:a="${A}"><dgm:ptLst>${pts.join('')}</dgm:ptLst>` +
    `<dgm:cxnLst>${cxns.join('')}</dgm:cxnLst><dgm:bg/><dgm:whole/></dgm:dataModel>`
  );
}

interface Diagram {
  readonly key: string;
  readonly title: string;
  readonly layout: string;
  readonly colors: string;
  readonly box: readonly [number, number, number, number];
  readonly seed: string;
}

const DIAGRAMS: readonly Diagram[] = [
  {
    key: 'process',
    title: 'Arrow process',
    layout: PROCESS_LAYOUT,
    colors: COLORS('accent1'),
    box: [1 * IN, 2 * IN, 10.33 * IN, 2.5 * IN],
    seed: seedData(3, 'arrow-process', 'accent1', ['Plan', 'Build', 'Test', 'Ship']),
  },
  {
    key: 'list',
    title: 'Vertical list',
    layout: LIST_LAYOUT,
    colors: COLORS('accent2'),
    box: [3.5 * IN, 1.4 * IN, 5.33 * IN, 5 * IN],
    seed: seedData(4, 'vertical-boxes', 'accent2', ['Goals', 'Risks', 'Owners']),
  },
];

const saved = (file: string): string | undefined => {
  const path = join(DIR, file);
  return existsSync(path) ? readFileSync(path, 'utf8') : undefined;
};

/** The deck. `seed`: the data written here, with no saved drawing. */
export function diagramEditingDeck(seed = false): FixtureOptions {
  const slides: FixtureSlide[] = [];
  const extraParts: [string, string, string][] = [];
  DIAGRAMS.forEach((d, i) => {
    const n = i + 1;
    const data = seed ? undefined : saved(`data-${d.key}.xml`);
    const drawing = seed ? undefined : saved(`drawing-${d.key}.xml`);
    if (!seed && (!data || !drawing)) throw new Error(`no saved drawing for ${d.key}`);
    const content: Record<string, string> = {
      data: data ?? d.seed,
      layout: d.layout,
      quickStyle: QUICK_STYLE,
      colors: d.colors,
      ...(drawing ? { drawing } : {}),
    };
    const parts = DIAGRAM_PARTS.filter(([file]) => content[file] !== undefined);
    slides.push({
      layout: 'blank',
      shapes:
        title(d.title) +
        frame(
          3,
          `Diagram ${n}`,
          [...d.box],
          DGM,
          `<dgm:relIds xmlns:dgm="${DGM}" r:dm="rIdDiagramData${n}" r:lo="rIdDiagramLayout${n}" r:qs="rIdDiagramQuickStyle${n}" r:cs="rIdDiagramColors${n}"/>`
        ),
      rels: parts.map(([file, rel, type]) => [
        rel === 'diagramDrawing'
          ? `rIdEditDiagramDrawing${n}`
          : `rId${rel[0]!.toUpperCase()}${rel.slice(1)}${n}`,
        type,
        `../diagrams/${file}${n}.xml`,
      ]) as FixtureSlide['rels'],
    });
    for (const [file, , , contentType] of parts)
      extraParts.push([`ppt/diagrams/${file}${n}.xml`, contentType, content[file]!]);
  });
  return { slides, extraParts };
}

if (import.meta.main) {
  const { values } = parseArgs({ options: { seed: { type: 'string' } } });
  if (values.seed) {
    mkdirSync(values.seed, { recursive: true });
    writeFileSync(join(values.seed, 'diagram-editing-seed.pptx'), buildPptx(diagramEditingDeck(true)));
  }
}
