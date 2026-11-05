import { describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { cssScopeErrors, nestedSelectorScoped, selectorScoped } from './check-css-scope.mjs';

const root = join(import.meta.dir, '..');

describe('check-css-scope', () => {
  test('the shipped chrome stylesheet styles nothing outside the editor', () => {
    const css = readFileSync(join(root, 'packages/core/src/styles/editor.css'), 'utf8');
    expect(cssScopeErrors(css)).toEqual([]);
  });

  test('a selector is scoped when it starts with a pptx- class and stays inside', () => {
    for (const selector of [
      '.pptx-editor',
      '.pptx-editor.dark .pptx-editor',
      'fieldset.pptx-dialog__section',
      '.pptx-menu:is(.a, .b) > button',
      '.pptx-button:not(:disabled):active',
      '.pptx-canvas:has(~ .pptx-find:not([hidden]))',
      '.pptx-swatches__section + .pptx-swatches__section',
      '.pptx-menu li + li',
      '.pptx-a > * + *',
      '*.pptx-x',
    ])
      expect([selector, selectorScoped(selector)]).toEqual([selector, true]);
    for (const selector of [
      'button',
      '*',
      ':root',
      '.dark .pptx-editor',
      'div > .pptx-menu',
      'body:has(.pptx-dialog[open])',
      ':not(.pptx-x)',
      ':is(body, .pptx-x)',
      ':where(.pptx-x)',
      'div[data-x=".pptx-y"]',
      '.pptx-editor ~ div',
      '.pptx-editor + *',
      '.pptx-editor +',
      '> .pptx-x',
    ])
      expect([selector, selectorScoped(selector)]).toEqual([selector, false]);
  });

  test('a nested selector may not step to a host sibling', () => {
    expect(nestedSelectorScoped('&:hover')).toBe(true);
    expect(nestedSelectorScoped('.pptx-a + .pptx-b')).toBe(true);
    expect(nestedSelectorScoped('button')).toBe(true);
    expect(nestedSelectorScoped('& + p')).toBe(false);
    expect(nestedSelectorScoped('+ p')).toBe(false);
    expect(nestedSelectorScoped('~ .pptx-x')).toBe(true);
    expect(nestedSelectorScoped('&.pptx-on > span')).toBe(true);
    expect(nestedSelectorScoped('body:has(&)')).toBe(false);
    expect(nestedSelectorScoped(':is(body, &)')).toBe(false);
    expect(nestedSelectorScoped('div:has(> &)')).toBe(false);
    expect(nestedSelectorScoped('.host &')).toBe(false);
  });

  test('host-wide rules, names, layers, and imports fail with their line', () => {
    const css = [
      '.pptx-editor { color: red; }',
      'body, .pptx-x { margin: 0; }',
      '@media (min-width: 1px) { html { color: red; } .pptx-y { color: red; } }',
      '@keyframes spin { to { rotate: 1turn; } }',
      '@keyframes pptx-spin { to { rotate: 1turn; } }',
      "@import 'host.css';",
      "@font-face { font-family: X; src: url('x.woff2'); }",
      '.pptx-z { color: red; &:hover { color: red; } @starting-style { opacity: 0; } }',
      '.pptx-z { & + p { color: red; } }',
      '.pptx-a { @media (x) { body:has(&) { margin: 0; } } }',
      '@layer base, components;',
      '@layer pptx-chrome { .pptx-a { color: red; } }',
      '@scope (.pptx-editor) { button { color: red; } }',
      '@scope (body) { .pptx-a { color: red; } }',
      ".pptx-q[data-a='}'] { content: '{'; }",
      '/* body { color: red; } */',
      '.pptx-s::after { content: "/*"; } body { color: red; } /* */',
    ].join('\n');
    expect(cssScopeErrors(css)).toEqual([
      '2: "body" is not scoped to a pptx- class',
      '3: "html" is not scoped to a pptx- class',
      '4: @keyframes spin needs a pptx- name',
      '6: @import is not allowed here',
      '7: @font-face is not allowed',
      '9: "& + p" is not scoped to a pptx- class',
      '10: "body:has(&)" is not scoped to a pptx- class',
      '11: @layer base needs a pptx- name',
      '11: @layer components needs a pptx- name',
      '14: @scope (body) needs a pptx- root',
      '17: "body" is not scoped to a pptx- class',
    ]);
  });
});
