// Contextual RTL glyphs, measured with the same faces and features as line layout.
import type { TextMeasurer } from './measurer.ts';
import type { GlyphRun } from './text-layout.ts';

/** One glyph relative to the run's left edge; y is positive down. */
export interface PositionedGlyph {
  readonly id: number;
  readonly text: string;
  readonly x: number;
  readonly y: number;
  /** Horizontal scale used to fit the word to its laid-out width. */
  readonly scaleX: number;
}

/** Shape each RTL word as SVG paint does, preserving justification between words. */
export function shapeRtlRun(run: GlyphRun, measurer: TextMeasurer): GlyphRun {
  if (!run.rtl || !measurer.shape || !run.text || run.picture) return run;
  const glyphs: PositionedGlyph[] = [];
  const size = run.sizePt * 12700;
  const n = run.text.length;
  const kerning = run.props.kern !== undefined && run.props.size >= run.props.kern;
  // Include spaces for mixed-direction PDF text. Tabs have no ink or text glyph.
  for (const match of [...run.text.matchAll(/[^ \t\u3000]+|[ \u3000]/gu)].reverse()) {
    const word = match[0];
    const shaped = measurer.shape(run.face, word, { rtl: true, kerning });
    const clusters = [...new Set(shaped.map((g) => g.cluster))].sort((a, b) => a - b);
    const ends = new Map(clusters.map((start, i) => [start, clusters[i + 1] ?? word.length]));
    const carriers = new Map<number, number>();
    shaped.forEach((g, i) => {
      const old = carriers.get(g.cluster);
      if (old === undefined || g.advance > shaped[old].advance) carriers.set(g.cluster, i);
    });
    const width = shaped.reduce((sum, g) => sum + g.advance, 0) * size;
    const left = run.x[n - match.index - word.length];
    const right = run.x[n - match.index];
    const scaleX = width > 0 ? (right - left) / width : 1;
    let pen = 0;
    shaped.forEach((g, i) => {
      glyphs.push({
        id: g.id,
        text: carriers.get(g.cluster) === i ? word.slice(g.cluster, ends.get(g.cluster)) : '',
        x: left - run.x[0] + (pen + g.offsetX) * size * scaleX,
        y: -g.offsetY * size,
        scaleX,
      });
      pen += g.advance;
    });
  }
  return { ...run, glyphs };
}
