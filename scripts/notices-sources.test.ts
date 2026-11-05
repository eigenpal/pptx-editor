// Every source note that names a third-party origin is listed in THIRD_PARTY_NOTICES.md (audit
// task 1.1.9). A source note is a marker in the leading comment of a file under `packages/*/src`;
// a note about EigenPal code (DOCX Editor, the site repository) needs no notice.
import { expect, test } from 'bun:test';
import { existsSync, readFileSync, readdirSync } from 'node:fs';
import { join, relative } from 'node:path';
import { PACKAGES } from './lib-packages.mjs';
import { VENDORED } from './vendored-notices.mjs';

const root = join(import.meta.dir, '..');
const MARKER =
  /\b(?:copied|ported|vendored|adapted|taken) from\b|\bpath data\b|\(c\)|\bcopyright\b|\blicensed under\b|\bApache-2\.0\b|\bMIT\b|\bBSD\b/gi;
const EIGENPAL = /docx-editor|DOCX Editor|EigenPal|docx-editor\.dev/;

function sourceFiles(dir: string): string[] {
  const out: string[] = [];
  for (const entry of readdirSync(dir, { withFileTypes: true })) {
    if (entry.name === '__tests__' || entry.name === 'node_modules') continue;
    const path = join(dir, entry.name);
    if (entry.isDirectory()) out.push(...sourceFiles(path));
    else if (/\.(?:ts|tsx|mts|js|mjs|css)$/.test(entry.name)) out.push(path);
  }
  return out;
}

/** The leading comment block of a file, comment markers removed, joined into one line. */
function leadingComment(text: string): string {
  const lines: string[] = [];
  for (const line of text.split('\n')) {
    const trimmed = line.trim();
    if (!/^(?:\/\/|\/\*|\*)/.test(trimmed)) break;
    lines.push(trimmed.replace(/^(?:\/\/+|\/\*+|\*+\/?)\s?/, '').replace(/\s*\*\/$/, ''));
  }
  return lines.join(' ');
}

/** Third-party source notes in `comment`: each marker whose sentence does not name EigenPal code. */
function thirdPartyNotes(comment: string): string[] {
  const notes = new Set<string>();
  for (const match of comment.matchAll(MARKER)) {
    const start = Math.max(0, comment.lastIndexOf('. ', match.index) + 1);
    const end = comment.indexOf('. ', match.index);
    const sentence = comment.slice(start, end === -1 ? undefined : end).trim();
    if (!EIGENPAL.test(sentence)) notes.add(sentence);
  }
  return [...notes];
}

test('the source-note scanner tells EigenPal notes from third-party notes', () => {
  expect(thirdPartyNotes('Copied from docx-editor/packages/core/src/x.ts (EigenPal).')).toEqual([]);
  expect(thirdPartyNotes('Adapted from libfoo (MIT). Second sentence.')).toHaveLength(1);
  expect(leadingComment('// a\n/* b\n * c */\nconst x = 1; // copied from y')).toBe('a b c');
});

for (const pkg of PACKAGES) {
  const listed = VENDORED[pkg.name]?.join('\n') ?? '';
  test(`${pkg.name}: every third-party source note is in the notices`, () => {
    const missing: string[] = [];
    for (const file of sourceFiles(join(root, pkg.root, 'src'))) {
      const notes = thirdPartyNotes(leadingComment(readFileSync(file, 'utf8')));
      if (notes.length === 0) continue;
      const path = relative(join(root, pkg.root), file);
      const named = path.split('/').some((_, i, parts) => {
        const prefix = parts.slice(0, i + 1).join('/');
        return listed.includes(`\`${prefix}\``) || listed.includes(`\`${prefix}/\``);
      });
      if (!named) missing.push(`${path}: ${notes[0]}`);
    }
    expect(missing).toEqual([]);
  });

  test(`${pkg.name}: every source path in the notices exists`, () => {
    for (const [, path] of listed.matchAll(/`(src\/[^`]+)`/g))
      expect(existsSync(join(root, pkg.root, path))).toBe(true);
  });
}
