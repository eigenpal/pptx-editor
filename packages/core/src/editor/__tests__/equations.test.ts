// Equations: Office Math (`a14:m`) draws from the choice that holds it, Insert Equation writes
// math and a plain-text fallback, Edit Equation replaces the math, and the caret steps over an
// equation as one character.

import { describe, expect, test } from 'bun:test';
import { buildPptx } from '../../../../../scripts/fixtures/builder.ts';
import { equationFixtures } from '../../../../../scripts/fixtures/equations.ts';
import { presentationInfo, serializeXmlPart, xmlPartOf } from '../../store/index.ts';
import { createPptxEditor, type PptxEditorInstance } from '../index.ts';

async function open(bytes: Uint8Array): Promise<PptxEditorInstance> {
  const e = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await e.load(bytes)).ok).toBe(true);
  return e;
}

const slideXml = (e: PptxEditorInstance, index = 0): string => {
  const pkg = e.session!.store.package;
  return serializeXmlPart(xmlPartOf(pkg, presentationInfo(pkg).slides[index]!.partName)!);
};

/** The texts and marks that the layout of slide `index` paints, by shape name. */
function painted(e: PptxEditorInstance, index = 0): Map<string, { text: string; marks: number }> {
  const out = new Map<string, { text: string; marks: number }>();
  for (const item of e.session!.layout(index)!.items) {
    if (item.kind !== 'shape' || !item.text) continue;
    const text = item.text.lines.map((l) => l.runs.map((r) => r.text).join('')).join('\n');
    const marks = item.text.lines.reduce((n, l) => n + (l.marks?.length ?? 0), 0);
    out.set(item.element.name, { text, marks });
  }
  return out;
}

const blankDeck = (): Uint8Array => buildPptx({ slides: [{ shapes: '' }] });

describe('equations draw from their math', () => {
  test('the math choice draws, with its rules and delimiters as marks', async () => {
    const e = await open(buildPptx({ slides: equationFixtures.equations! }));
    e.selectSlide(1);
    const fractions = painted(e, 1);
    // y = a/b: the letters in math italic, and one fraction rule.
    const frac = fractions.get('Equation 3')!;
    expect(frac.text).toBe('𝑦=𝑎𝑏');
    expect(frac.marks).toBe(1);
    // The radical draws its sign and its rule.
    expect(fractions.get('Equation 6')!.marks).toBe(2);
    // No fallback text draws.
    expect([...fractions.values()].some((v) => v.text.includes('√'))).toBe(false);
  });

  test('an equation with a construct the layout does not draw keeps its fallback', async () => {
    const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
    const A14 = 'http://schemas.microsoft.com/office/drawing/2010/main';
    const M = 'http://schemas.openxmlformats.org/officeDocument/2006/math';
    const sp = (body: string): string =>
      '<p:sp><p:nvSpPr><p:cNvPr id="3" name="Eq"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>' +
      '<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>' +
      `<p:txBody><a:bodyPr/><a:lstStyle/><a:p>${body}</a:p></p:txBody></p:sp>`;
    const math = `<a14:m><m:oMath xmlns:m="${M}"><m:unknown/></m:oMath></a14:m>`;
    const deck = buildPptx({
      slides: [
        {
          shapes:
            `<mc:AlternateContent xmlns:mc="${MC}"><mc:Choice xmlns:a14="${A14}" Requires="a14">${sp(math)}</mc:Choice>` +
            `<mc:Fallback>${sp('<a:r><a:rPr lang="en-US"/><a:t>fallback</a:t></a:r>')}</mc:Fallback></mc:AlternateContent>`,
        },
      ],
    });
    const e = await open(deck);
    expect(painted(e).get('Eq')!.text.replace(/\n/g, '')).toBe('fallback');
    expect(slideXml(await open(e.save()!))).toContain('<m:unknown/>');
  });
});

describe('lines that hold equations', () => {
  /** Each line of shape `name`: its box, and the ink of its runs and marks (EMU, y down). */
  function inkLines(e: PptxEditorInstance, name: string) {
    const item = e
      .session!.layout(0)!
      .items.find((i) => i.kind === 'shape' && i.element.name === name);
    if (!item || item.kind !== 'shape' || !item.text) throw new Error(name);
    return item.text.lines.map((l) => {
      let top = Infinity;
      let bottom = -Infinity;
      for (const r of l.runs) {
        // The fixed measurer draws each glyph as a box from the baseline up to 0.7 em.
        const em = r.sizePt * 12_700;
        top = Math.min(top, l.baseline - r.rise - 0.7 * em);
        bottom = Math.max(bottom, l.baseline - r.rise);
      }
      for (const m of l.marks ?? [])
        for (const c of m.commands)
          if (c.op !== 'Z') {
            top = Math.min(top, l.baseline - c.y);
            bottom = Math.max(bottom, l.baseline - c.y);
          }
      return { top: l.top, bottom: l.top + l.height, inkTop: top, inkBottom: bottom };
    });
  }

  test('each line holds the ink of its equation, so stacked lines do not overlap', async () => {
    const e = await open(buildPptx({ slides: equationFixtures['equation-lines']! }));
    // Baselines round to whole points: the ink may pass its line by up to half a point.
    const half = 6_350;
    const gaps = (name: string) => {
      const lines = inkLines(e, name);
      for (const l of lines) {
        expect(l.inkTop).toBeGreaterThanOrEqual(l.top - half);
        expect(l.inkBottom).toBeLessThanOrEqual(l.bottom + half);
      }
      return lines.slice(1).map((l, i) => (l.inkTop - lines[i]!.inkBottom) / 12_700);
    };
    // Two display fractions: the line of the second starts where the ink of the first ends,
    // with no face descent between them (0.7 pt with the fixed measurer, 7.7 pt with the face's
    // descent).
    const [fractions] = gaps('Stack 3');
    expect(fractions).toBeGreaterThanOrEqual(-0.5);
    expect(fractions).toBeLessThan(1);
    // Text, a fraction, and text: the text below starts at the fraction's ink, plus the room
    // that the text line keeps above its own ink (1.4 pt with the fixed measurer).
    const [, below] = gaps('Stack 4');
    expect(below).toBeGreaterThanOrEqual(-0.5);
    expect(below).toBeLessThan(2);
    // At 150 % spacing, the line of an equation keeps its extent and the spacing adds half a
    // single line on top (10 pt with the fixed measurer), not only the room the ink leaves.
    const [spaced] = gaps('Stack 6');
    expect(spaced! - fractions!).toBeGreaterThan(9);
    expect(spaced! - fractions!).toBeLessThan(11);
  });
});

describe('equation lines at exact spacing', () => {
  /** Text, two display fractions, and text at 18 pt, with exact line spacing `pts` (1/100 pt). */
  async function spaced(pts: number) {
    const A14 = 'http://schemas.microsoft.com/office/drawing/2010/main';
    const M = 'http://schemas.openxmlformats.org/officeDocument/2006/math';
    const pPr = `<a:pPr><a:lnSpc><a:spcPts val="${pts}"/></a:lnSpc></a:pPr>`;
    const run = '<a:rPr lang="en-US" sz="1800"/>';
    const frac =
      `<m:f><m:num><m:r>${run}<m:t>a</m:t></m:r></m:num>` +
      `<m:den><m:r>${run}<m:t>b</m:t></m:r></m:den></m:f>`;
    const math = `<a:p>${pPr}<a14:m><m:oMathPara xmlns:m="${M}"><m:oMath>${frac}</m:oMath></m:oMathPara></a14:m></a:p>`;
    const text = `<a:p>${pPr}<a:r>${run}<a:t>Text</a:t></a:r></a:p>`;
    const shape =
      `<p:sp xmlns:a14="${A14}"><p:nvSpPr><p:cNvPr id="3" name="Eq"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>` +
      '<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="3657600" cy="3657600"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>' +
      `<p:txBody><a:bodyPr anchor="t"><a:noAutofit/></a:bodyPr><a:lstStyle/>${text}${math}${math}${text}</p:txBody></p:sp>`;
    const e = await open(buildPptx({ slides: [{ shapes: shape }] }));
    const item = e.session!.layout(0)!.items.find((i) => i.kind === 'shape');
    if (!item || item.kind !== 'shape' || !item.text) throw new Error('Eq');
    return item.text.lines.map((l) => l.height / 12_700);
  }

  test('a line taller than the equation keeps its exact height', async () => {
    for (const h of await spaced(6000)) expect(h).toBeCloseTo(60, 1);
  });

  test('a line shorter than the equation grows by what the equation needs past it', async () => {
    const [first, eq1, eq2, last] = await spaced(4000);
    expect(first).toBeCloseTo(40, 1);
    expect(last).toBeCloseTo(40, 1);
    // The equation fills the line's own 40 pt first: the line grows by 2 pt, not by the
    // single-line box plus the exact height.
    for (const h of [eq1!, eq2!]) {
      expect(h).toBeGreaterThan(40);
      expect(h).toBeLessThan(45);
    }
  });
});

describe('Insert > Equation', () => {
  test('writes the math, a plain-text fallback, and selects the new text box', async () => {
    const e = await open(blankDeck());
    const r = e.exec({ type: 'insertEquation', linear: 'x=(-b\\pm\\sqrt(b^2-4ac))/(2a)' });
    expect(r.ok).toBe(true);
    const xml = slideXml(e);
    expect(xml).toContain('Requires="a14"');
    expect(xml).toContain('<m:f><m:num>');
    expect(xml).toContain('<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/>');
    expect(xml).toContain('<a:latin typeface="Cambria Math"/>');
    expect(xml).toContain('<mc:Fallback>');
    expect(xml).toContain('x=(−b±√(b^2−4ac))/(2a)');
    const sel = e.snapshot().selection;
    expect(sel.kind).toBe('shapes');
    expect(e.snapshot().equation?.linear).toBe('x=(-b±\\sqrt(b^2-4ac))/(2a)');
    // The save reopens with the same equation.
    const again = await open(e.save()!);
    expect(slideXml(again)).toContain('<m:f><m:num>');
  });

  test('refuses a linear form it does not read, with no undo step', async () => {
    const e = await open(blankDeck());
    expect(e.exec({ type: 'insertEquation', linear: 'x^' }).ok).toBe(false);
    expect(e.exec({ type: 'insertEquation', linear: '\\unknown' }).ok).toBe(false);
    expect(e.exec({ type: 'insertEquation', linear: '' }).ok).toBe(false);
    expect(e.snapshot().canUndo).toBe(false);
  });

  test('escapes the text of the fallback and the math', async () => {
    const e = await open(blankDeck());
    expect(e.exec({ type: 'insertEquation', linear: 'a<b' }).ok).toBe(true);
    const xml = slideXml(e);
    expect(xml).toContain('<m:t>&lt;</m:t>');
    expect(xml).toContain('<a:t>a&lt;b</a:t>');
  });
});

describe('Edit Equation', () => {
  test('replaces the math, keeps the run font, and drops the stale fallback', async () => {
    const e = await open(blankDeck());
    e.exec({ type: 'insertEquation', linear: 'E=mc^2', font: 'STIX Two Math' });
    const state = e.snapshot().equation!;
    expect(state.linear).toBe('E=mc^2');
    expect(e.exec({ type: 'setEquation', nodeId: state.nodeId, linear: 'a^2+b^2=c^2' }).ok).toBe(
      true
    );
    const xml = slideXml(e);
    expect(xml).toContain('<a:latin typeface="STIX Two Math"/>');
    expect(xml).not.toContain('<mc:Fallback>');
    expect(e.snapshot().equation?.linear).toBe('a^2+b^2=c^2');
    e.exec({ type: 'undo' });
    expect(e.snapshot().equation?.linear).toBe('E=mc^2');
    expect(slideXml(e)).toContain('<mc:Fallback>');
  });

  test('fits the box of the edited equation, also when it is not selected', async () => {
    const e = await open(blankDeck());
    e.exec({ type: 'insertEquation', linear: 'x' });
    const { nodeId } = e.snapshot().equation!;
    const height = () =>
      Number(/name="Equation 1".*?<a:ext cx="\d+" cy="(\d+)"/.exec(slideXml(e))![1]);
    const before = height();
    e.setSelection({ kind: 'none' });
    const tall = '(\\matrix(a&b@c&d@e&f@g&h))';
    expect(e.exec({ type: 'setEquation', nodeId, linear: tall }).ok).toBe(true);
    expect(height()).toBeGreaterThan(before);
  });

  test('a format change of the selected box keeps the math whole and drawn', async () => {
    const e = await open(blankDeck());
    e.exec({ type: 'insertEquation', linear: 'E=mc^2' });
    e.exec({ type: 'toggleBold' });
    const xml = slideXml(e);
    // The math runs keep their own format; the paragraph mark takes the bold.
    expect(xml).toMatch(/<a14:m><m:oMathPara/);
    expect(xml).toContain('<a:endParaRPr lang="en-US" dirty="0" b="1"/>');
    expect(e.snapshot().equation?.linear).toBe('E=mc^2');
  });

  test('a hidden radical degree or n-ary limit with content has no linear form', async () => {
    const M = 'http://schemas.openxmlformats.org/officeDocument/2006/math';
    const run = (t: string) => `<m:r><m:t>${t}</m:t></m:r>`;
    const zone = (math: string) =>
      `<p:sp><p:nvSpPr><p:cNvPr id="2" name="Eq"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>` +
      `<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="3000000" cy="700000"/></a:xfrm>` +
      `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/>` +
      `<a:p><a14:m xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main">` +
      `<m:oMathPara xmlns:m="${M}"><m:oMath>${math}</m:oMath></m:oMathPara></a14:m></a:p></p:txBody></p:sp>`;
    const linearOf = async (math: string) => {
      const e = await open(buildPptx({ slides: [{ shapes: zone(math) }] }));
      const id = e.snapshot().shapes!.find((s) => s.name === 'Eq')!.nodeId;
      e.setSelection({ kind: 'shapes', nodeIds: [id] });
      return e.snapshot().equation?.linear;
    };
    const rad = (pr: string) =>
      `<m:rad><m:radPr>${pr}</m:radPr><m:deg>${run('3')}</m:deg><m:e>${run('x')}</m:e></m:rad>`;
    expect(await linearOf(rad('<m:degHide/>'))).toBeNull();
    expect(await linearOf(rad('<m:degHide m:val="off"/>'))).toBe('\\sqrt[3](x)');
    const sum = (pr: string) =>
      `<m:nary><m:naryPr><m:chr m:val="∑"/>${pr}</m:naryPr><m:sub>${run('k')}</m:sub>` +
      `<m:sup/><m:e>${run('k')}</m:e></m:nary>`;
    expect(await linearOf(sum('<m:subHide m:val="false"/>'))).toBe('\\sum_k k');
    expect(await linearOf(sum('<m:subHide/>'))).toBeNull();
  });

  test('an equation with a construct the linear form does not cover has no linear form', async () => {
    const e = await open(buildPptx({ slides: equationFixtures.equations! }));
    e.selectSlide(3);
    const id = e.snapshot().shapes!.find((s) => s.name === 'Equation 8')!.nodeId;
    e.setSelection({ kind: 'shapes', nodeIds: [id] });
    // A group character (an underbrace) has no linear form.
    expect(e.snapshot().equation).toEqual({ nodeId: expect.any(String), linear: null });
  });
});

describe('the caret and the text model', () => {
  test('an equation is one character between the text around it', async () => {
    const e = await open(blankDeck());
    e.exec({ type: 'insertEquation', linear: 'x^2' });
    const sel = e.snapshot().selection;
    const id = sel.kind === 'shapes' ? sel.nodeIds[0]! : '';
    const caret = { paragraph: 0, offset: 1 };
    e.setSelection({ kind: 'text', nodeId: id, anchor: caret, focus: caret });
    expect(e.input('text', ' is a square').ok).toBe(true);
    const xml = slideXml(e);
    expect(xml).toMatch(/<\/a14:m><a:r><a:rPr[^>]*\/?>(<\/a:rPr>)?<a:t> is a square<\/a:t><\/a:r>/);
    // Backspace after the equation removes the whole equation.
    e.setSelection({ kind: 'text', nodeId: id, anchor: caret, focus: caret });
    expect(e.input('deleteBackward').ok).toBe(true);
    expect(slideXml(e)).not.toContain('<a14:m>');
  });
});
