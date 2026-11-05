import { describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';

// The toolbar's picker fields (the font family field and the font size box) use the surface,
// border, and hover tokens, as DOCX Editor's font pickers do. The toggle tokens
// (`--doc-toggle-active-*`) are for pressed toggles only: the field keeps its text color, so a
// toggle slab under it (near-black in light mode) makes the label unreadable (deleted archive issue 76).
// The React and Vue adapter stylesheets only import this file, so both adapters get these rules.
//
// The checks read declarations, not raw text, so a reformatted stylesheet still passes.

const css = readFileSync(new URL('../../styles/editor.css', import.meta.url), 'utf8').replace(
  /\/\*[\s\S]*?\*\//g,
  ''
);

type Rule = { selectors: string[]; decls: Map<string, string> };

/** Every rule in the stylesheet, with its selector list and its declarations. */
function rules(): Rule[] {
  const out: Rule[] = [];
  const re = /([^{}]+)\{([^{}]*)\}/g;
  for (let m = re.exec(css); m; m = re.exec(css)) {
    const selectors = m[1]!.split(',').map((s) => s.replace(/\s+/g, ' ').trim());
    const decls = new Map<string, string>();
    for (const d of m[2]!.split(';')) {
      const i = d.indexOf(':');
      if (i > 0)
        decls.set(
          d.slice(0, i).trim(),
          d
            .slice(i + 1)
            .replace(/\s+/g, ' ')
            .trim()
        );
    }
    out.push({ selectors, decls });
  }
  return out;
}

const fieldRules = rules().filter((r) =>
  r.selectors.some((s) => /\.pptx-toolbar__(field|size)\b/.test(s))
);

/** The value of `property` in the rule that lists `selector`. Fails with the selector's name. */
function value(selector: string, property: string): string | undefined {
  const found = fieldRules.find((r) => r.selectors.includes(selector));
  if (!found) throw new Error(`editor.css has no rule for \`${selector}\``);
  return found.decls.get(property);
}

/** The hex value of a token in the light (root) or dark scope. */
function token(mode: 'light' | 'dark', name: string): string {
  const darkStart = css.indexOf('.pptx-editor.dark,');
  const scope = mode === 'light' ? css.slice(0, darkStart) : css.slice(darkStart);
  const m = new RegExp(`${name}:\\s*(#[0-9a-fA-F]{6})`).exec(scope);
  if (!m) throw new Error(`no ${mode} hex value for ${name}`);
  return m[1]!;
}

function luminance(hex: string): number {
  const [r, g, b] = [1, 3, 5].map((i) => {
    const c = parseInt(hex.slice(i, i + 2), 16) / 255;
    return c <= 0.03928 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4;
  });
  return 0.2126 * r! + 0.7152 * g! + 0.0722 * b!;
}

function contrast(a: string, b: string): number {
  const [hi, lo] = [luminance(a), luminance(b)].sort((x, y) => y - x);
  return (hi! + 0.05) / (lo! + 0.05);
}

const HOVER = ".pptx-toolbar__field:hover:not(:disabled):not([aria-disabled='true'])";

describe('toolbar field tokens', () => {
  test('no field state uses the toggle tokens', () => {
    expect(fieldRules.length).toBeGreaterThan(0);
    for (const r of fieldRules)
      for (const [prop, v] of r.decls)
        expect(`${r.selectors.join(', ')} { ${prop}: ${v} }`).not.toContain('--doc-toggle-active');
  });

  test('the family field rests on the surface and takes the hover token on hover and when open', () => {
    expect(value('.pptx-toolbar__field', 'background')).toBe('var(--doc-surface)');
    expect(value('.pptx-toolbar__field', 'border')).toBe('1px solid var(--doc-border)');
    expect(value(HOVER, 'background')).toBe('var(--doc-bg-hover)');
    expect(value(".pptx-toolbar__field[aria-expanded='true']", 'background')).toBe(
      'var(--doc-bg-hover)'
    );
  });

  test('the field label meets 4.5:1 on its rest and hover fills, light and dark', () => {
    expect(value('.pptx-toolbar__field', 'color')).toBe('var(--doc-text)');
    for (const mode of ['light', 'dark'] as const)
      for (const bg of ['--doc-surface', '--doc-bg-hover'])
        expect(contrast(token(mode, '--doc-text'), token(mode, bg))).toBeGreaterThanOrEqual(4.5);
  });

  test('a field disabled by aria-disabled is dimmed like a disabled one', () => {
    const dim = value('.pptx-toolbar__field:disabled', 'opacity');
    expect(dim).toBeDefined();
    expect(value(".pptx-toolbar__field[aria-disabled='true']", 'opacity')).toBe(dim);
    expect(value(".pptx-toolbar__size[aria-disabled='true']", 'opacity')).toBe(dim);
  });

  test('the size box shows the focus ring', () => {
    expect(value('.pptx-toolbar__size:focus', 'outline')).toBe('2px solid var(--doc-focus-ring)');
  });

  test('the field rules use tokens, not literal colors', () => {
    for (const r of fieldRules)
      for (const v of r.decls.values()) expect(v).not.toMatch(/#[0-9a-f]{3,8}\b|rgba?\(|hsla?\(/i);
  });
});
