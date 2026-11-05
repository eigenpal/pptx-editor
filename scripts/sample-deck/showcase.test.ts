// The showcase deck (examples/vite/public/showcase.pptx) opens in the editor, every slide lays
// out, and no slide uses an element the engine does not render or an approximation that
// the public fidelity page (`docs/site/content/pptx-fidelity.mdx`) lists, except the ones named below.

import { describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { unzipSync } from 'fflate';
import { PNG } from 'pngjs';
import { createPptxEditor } from '../../packages/core/src/editor/index.ts';
import { notesText } from '../../packages/core/src/automation/commands/notes-commands.ts';
import { buildShowcase } from './showcase.ts';

const FILE = join(import.meta.dir, '../../examples/vite/public/showcase.pptx');

/** Approximations a slide may report. Every slide with text uses the line model. */
const ALLOWED = new Set(['line-model']);
/**
 * The fills slide (index 6) has a two-stop gradient, the pictures slide (index 9) recolors
 * pictures, and the links slide (index 13) shows hyperlink text; all are listed approximations.
 */
const ALLOWED_ON: Record<number, readonly string[]> = {
  6: ['gradient-colors'],
  9: ['picture-recolor'],
  13: ['hyperlink-text'],
};

describe('showcase deck', () => {
  test('the committed file holds what the builder writes', () => {
    // Zip entry times depend on the time zone, so compare the parts, not the zip bytes.
    const committed = unzipSync(new Uint8Array(readFileSync(FILE)));
    const built = unzipSync(buildShowcase());
    expect(Object.keys(committed).sort()).toEqual(Object.keys(built).sort());
    // PNG bytes depend on the zlib build of the runtime, so compare PNG parts by their pixels.
    const content = (name: string, bytes: Uint8Array): Buffer =>
      name.endsWith('.png') ? PNG.sync.read(Buffer.from(bytes)).data : Buffer.from(bytes);
    for (const [name, bytes] of Object.entries(built))
      expect({
        name,
        same: content(name, bytes).equals(content(name, committed[name])),
      }).toEqual({ name, same: true });
  });

  test('opens, lays out every slide, and reports no unsupported element', async () => {
    const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    expect((await editor.load(new Uint8Array(readFileSync(FILE)))).ok).toBe(true);
    const session = editor.session!;
    const slides = session.info.slides;
    expect(slides.length).toBe(17);
    expect(slides.map((s) => s.hidden)).toEqual(slides.map((_, i) => i === 15));
    for (let i = 0; i < slides.length; i += 1) {
      const layout = session.layout(i);
      expect(layout).toBeDefined();
      expect(layout!.items.length).toBeGreaterThan(0);
      const problems = layout!.diagnostics.filter((d) => {
        if (d.startsWith('font substituted:')) return false;
        if (!d.startsWith('approximation: ')) return true;
        const id = d.slice('approximation: '.length);
        return !ALLOWED.has(id) && !(ALLOWED_ON[i] ?? []).includes(id);
      });
      expect({ slide: i + 1, problems }).toEqual({ slide: i + 1, problems: [] });
    }
  });

  test('a text box added on a dark slide gets light text, not the background color', async () => {
    const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    expect((await editor.load(new Uint8Array(readFileSync(FILE)))).ok).toBe(true);
    for (const index of [0, 16]) {
      editor.selectSlide(index);
      expect(editor.exec({ type: 'insertTextBox' }).ok).toBe(true);
      const sel = editor.snapshot().selection;
      const id = sel.kind === 'text' ? sel.nodeId : sel.kind === 'shapes' ? sel.nodeIds[0] : '';
      const item = editor.session!.layout(index)!.items.find((i) => i.element.nodeId === id)!;
      const fill = item.kind === 'shape' ? item.element.text?.paragraphs[0].endProps.fill : null;
      const light = fill?.type === 'solid' && fill.color.r + fill.color.g + fill.color.b > 600;
      expect({ slide: index + 1, light }).toEqual({ slide: index + 1, light: true });
    }
  });

  test('every slide has speaker notes that say what to try', async () => {
    const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    expect((await editor.load(new Uint8Array(readFileSync(FILE)))).ok).toBe(true);
    const pkg = editor.session!.store.package;
    for (let i = 0; i < 17; i += 1) expect(notesText(pkg, i)).toContain('Try:');
  });
});
