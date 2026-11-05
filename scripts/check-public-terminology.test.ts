import { expect, test } from 'bun:test';
import { mkdtempSync, mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join } from 'node:path';
import { spawnSync } from 'node:child_process';

function check(extra: Record<string, string>) {
  const root = mkdtempSync(join(tmpdir(), 'pptx-terminology-'));
  try {
    // Minimal authored surfaces, not a repository checkout.
    const files: Record<string, string> = {
      'README.md': 'Presentation editor',
      'CONTRIBUTING.md': '',
      'REVIEW.md': '',
      'SECURITY.md': '',
      'packages/i18n/README.md': '',
      'packages/i18n/en.json': '{}',
      'packages/i18n/locales/.gitkeep': '',
      'docs/site/content/index.mdx': '',
      'examples/vite/src/App.tsx': '',
      'examples/vite/index.html': '',
      'examples/vite/public/icon.svg': '',
      'examples/chat/src/App.tsx': '',
      'examples/chat/index.html': '',
      '.github/ISSUE_TEMPLATE/bug.yml': '',
      '.github/pull_request_template.md': '',
      'scripts/check-public-terminology.mjs': readFileSync(
        join(import.meta.dir, 'check-public-terminology.mjs'),
        'utf8'
      ),
      ...extra,
    };
    for (const [path, text] of Object.entries(files)) {
      const target = join(root, path);
      mkdirSync(dirname(target), { recursive: true });
      writeFileSync(target, text);
    }
    return spawnSync('node', [join(root, 'scripts/check-public-terminology.mjs')], {
      encoding: 'utf8',
    });
  } finally {
    rmSync(root, { recursive: true, force: true });
  }
}

const vendorName = ['Power', 'Point'].join('');

test('checks translated catalogs, including language tags with script and region', () => {
  const result = check({
    'packages/i18n/locales/zh-Hant-TW.json': JSON.stringify({ label: vendorName }),
  });
  expect(result.status).toBe(1);
  expect(result.stderr).toContain('packages/i18n/locales/zh-Hant-TW.json:1');
});

test('keeps source attribution and legal notices outside the authored-product gate', () => {
  const result = check({
    'packages/i18n/locales/fr.json': '{"label":"Présentation"}',
    LICENSE: vendorName,
    'reference/NOTICE': vendorName,
    'packages/core/README.md': '',
    'packages/core/src/source.ts': `// Source: ${vendorName}`,
  });
  expect(result.status).toBe(0);
  expect(result.stdout).toContain('Public terminology check passed');
});

test('refuses the trademarked control name in authored text', () => {
  const result = check({
    'packages/i18n/en.json': JSON.stringify({ label: `${['Active', 'X'].join('')} controls` }),
  });
  expect(result.status).toBe(1);
  expect(result.stderr).toContain('packages/i18n/en.json:1');
});
