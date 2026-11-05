import { describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';

// The pane text in both color modes meets WCAG 2.1 AA 1.4.3 (4.5:1), computed from the token
// values in `editor.css`: the rail slide numbers and the notes placeholder use
// `--doc-text-muted`, on the surface and on the notes field.

const css = readFileSync(new URL('../../styles/editor.css', import.meta.url), 'utf8');

function block(selector: string): string {
  const start = css.indexOf(selector);
  expect(start).toBeGreaterThanOrEqual(0);
  return css.slice(start, css.indexOf('}', start));
}

function token(scope: string, name: string): string {
  const m = new RegExp(`${name}:\\s*(#[0-9a-fA-F]{6})`).exec(scope);
  expect(m).not.toBeNull();
  return m![1]!;
}

function luminance(hex: string): number {
  const [r, g, b] = [1, 3, 5].map((i) => {
    const c = parseInt(hex.slice(i, i + 2), 16) / 255;
    return c <= 0.03928 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4;
  });
  return 0.2126 * r! + 0.7152 * g! + 0.0722 * b!;
}

const contrast = (a: string, b: string): number => {
  const [hi, lo] = [luminance(a), luminance(b)].sort((x, y) => y - x);
  return (hi! + 0.05) / (lo! + 0.05);
};

describe('pane text contrast', () => {
  test('the rail numbers and the notes placeholder use the muted token', () => {
    expect(block('.pptx-rail__number {')).toContain('var(--doc-text-muted)');
    expect(block('.pptx-notes__text:empty::before {')).toContain('var(--doc-text-muted)');
  });

  test('muted text is 4.5:1 or more on the surface and the notes field, light and dark', () => {
    const dark = block('.pptx-editor.dark,\n.pptx-editor.dark .pptx-editor {');
    const light = css.slice(0, css.indexOf('.pptx-editor.dark,\n.pptx-editor.dark .pptx-editor {'));
    for (const scope of [light, dark]) {
      const muted = token(scope, '--doc-text-muted');
      for (const bg of ['--doc-surface', '--doc-bg-input'])
        expect(contrast(muted, token(scope, bg))).toBeGreaterThanOrEqual(4.5);
    }
  });
});
