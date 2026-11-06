import { describe, expect, test } from 'bun:test';
import { readFileSync, readdirSync } from 'node:fs';
import { dirname, join } from 'node:path';

const CONTENT = join(import.meta.dir, '../content');
function files(dir = CONTENT, prefix = ''): string[] {
  return readdirSync(dir, { withFileTypes: true }).flatMap((entry) => {
    const path = `${prefix}${entry.name}`;
    return entry.isDirectory() ? files(join(dir, entry.name), `${path}/`) : [path];
  });
}
const all = files();
const pages = all.filter((path) => path.endsWith('.mdx'));
const read = (path: string) => readFileSync(join(CONTENT, path), 'utf8');
const slug = (path: string) => path.replace(/(?:\/index)?\.mdx$/, '').replace(/^index$/, '');
const slugs = new Set(pages.map(slug));
const normalize = (path: string) => path.replace(/\/index$/, '').replace(/^index$/, '');
function entries(path: string): string[] {
  const meta = JSON.parse(read(path)) as { pages: string[] };
  return meta.pages.filter((entry) => !entry.startsWith('---'));
}

describe('public documentation navigation', () => {
  test('every page has search and navigation metadata', () => {
    const failures = pages.flatMap((page) => {
      const frontmatter = read(page).match(/^---\n([\s\S]*?)\n---/u)?.[1] ?? '';
      return ['title', 'description', 'category']
        .filter((key) => !new RegExp(`^${key}:\\s*\\S`, 'm').test(frontmatter))
        .map((key) => `${page}: missing ${key}`);
    });
    expect(failures).toEqual([]);
  });

  test('root navigation lists every page exactly once', () => {
    const root = entries('meta.json').map(normalize);
    expect(root.filter((entry) => !slugs.has(entry))).toEqual([]);
    expect([...slugs].filter((entry) => !root.includes(entry))).toEqual([]);
    expect(root.filter((entry, index) => root.indexOf(entry) !== index)).toEqual([]);
  });

  test('folder navigation resolves and includes each direct page', () => {
    const failures: string[] = [];
    for (const meta of all.filter((path) => path.endsWith('/meta.json'))) {
      const folder = dirname(meta);
      const local = entries(meta).map((entry) => normalize(`${folder}/${entry}`));
      for (const entry of local) {
        if (!slugs.has(entry) && !all.includes(`${entry}/meta.json`)) {
          failures.push(`${meta}: unknown ${entry}`);
        }
      }
      for (const page of pages.filter((page) => dirname(page) === folder)) {
        if (!local.includes(slug(page))) failures.push(`${meta}: missing ${page}`);
      }
    }
    for (const page of pages.filter((page) => dirname(page) !== '.')) {
      if (!all.includes(`${dirname(page)}/meta.json`))
        failures.push(`${page}: missing folder meta`);
    }
    expect(failures).toEqual([]);
  });

  test('internal page links resolve', () => {
    const failures = pages.flatMap((page) => {
      const prose = read(page).replace(/^```[^\n]*\n[\s\S]*?^```/gm, '');
      return [...prose.matchAll(/(?:\]\(|href=["'])(\/docs(?:\/[^\s"')#?]*)?)/g)]
        .map((match) => match[1].replace(/^\/docs\/?/, '').replace(/\/$/, ''))
        .filter((target) => !slugs.has(target))
        .map((target) => `${page}: /docs/${target}`);
    });
    expect(failures).toEqual([]);
  });
});
