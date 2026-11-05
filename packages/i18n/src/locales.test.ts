// The Japanese, Korean, and Simplified Chinese catalogs: they translate, fall back to English where an
// entry is `null`, and keep each English placeholder.

import { describe, expect, test } from 'bun:test';
import { createTranslator, englishCatalog, type TranslationKey } from './index.ts';
import ja from './ja.ts';
import ko from './ko.ts';
import zhCN from './zh-CN.ts';

function leaves(node: unknown, path = '', out: [string, unknown][] = []): [string, unknown][] {
  for (const [k, v] of Object.entries(node as Record<string, unknown>)) {
    if (v !== null && typeof v === 'object') leaves(v, `${path}${k}.`, out);
    else out.push([`${path}${k}`, v]);
  }
  return out;
}

describe.each([
  ['ja', ja, '太字 (Ctrl+B)'],
  ['ko', ko, '굵게(Ctrl+B)'],
  ['zh-CN', zhCN, '加粗(Ctrl+B)'],
] as const)('the %s catalog', (_name, catalog, bold) => {
  const t = createTranslator(catalog);

  test('translates, and falls back to English for a null entry', () => {
    expect(t('toolbar.bold')).toBe(bold);
    const untranslated = leaves(catalog).find(([, v]) => v === null)![0] as TranslationKey;
    expect(t(untranslated)).toBe(createTranslator()(untranslated));
  });

  test('has the English shape, and each string keeps the English placeholders', () => {
    const english = new Map(leaves(englishCatalog));
    const own = leaves(catalog);
    expect(own.map(([k]) => k)).toEqual([...english.keys()]);
    const params = (s: string) => [...s.matchAll(/\{(\w+)\}/g)].map((m) => m[1]).sort();
    for (const [k, v] of own)
      if (typeof v === 'string') expect(params(v)).toEqual(params(String(english.get(k))));
  });
});

describe('East Asian punctuation', () => {
  const HAN = '[\\u3040-\\u30ff\\u3400-\\u9fff]';

  test('zh-CN puts full-width punctuation next to Chinese text', () => {
    // ASCII stays in formats and identifiers: "1,000", "16:9", "PNG, .zip".
    const ascii = new RegExp(`${HAN}\\s*[,:;?!]|[,:;?!]\\s*${HAN}`, 'u');
    const bad = leaves(zhCN).filter(([, v]) => typeof v === 'string' && ascii.test(v));
    expect(bad).toEqual([]);
  });

  test('ja uses the ideographic comma and full-width question marks in Japanese text', () => {
    const ascii = new RegExp(`${HAN}\\s*,\\s*${HAN}|${HAN}[?!]`, 'u');
    const bad = leaves(ja).filter(([, v]) => typeof v === 'string' && ascii.test(v));
    expect(bad).toEqual([]);
  });
});
