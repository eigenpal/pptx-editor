/** Check authored product surfaces only; see CONTRIBUTING.md for scope and exclusions. */
import { readdirSync, readFileSync } from 'node:fs';
import { dirname, join, relative } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = dirname(dirname(fileURLToPath(import.meta.url)));
function walk(dir) {
  return readdirSync(join(root, dir), { withFileTypes: true }).flatMap((entry) => {
    const path = join(dir, entry.name);
    return entry.isDirectory() ? walk(path) : [path];
  });
}
const files = [
  'README.md',
  'CONTRIBUTING.md',
  'SECURITY.md',
  ...readdirSync(join(root, 'packages')).map((name) => `packages/${name}/README.md`),
  ...walk('docs/site').filter((path) => /\.(mdx?|json)$/.test(path)),
  'packages/i18n/en.json',
  ...walk('packages/i18n/locales').filter((path) => /\.json$/.test(path)),
  ...walk('examples/vite/src').filter((path) => /\.(tsx?|html)$/.test(path)),
  'examples/vite/index.html',
  'examples/vite/public/icon.svg',
  ...walk('examples/chat/src').filter((path) => /\.(tsx?|html)$/.test(path)),
  'examples/chat/index.html',
  ...walk('.github/ISSUE_TEMPLATE'),
  '.github/pull_request_template.md',
];
// Build the forbidden names without adding them to authored product text: a product name and a
// trademarked control name (ECMA-376 calls the element `p:control`, an embedded control).
const vendorProduct = new RegExp(
  [['power', 'point'].join(''), ['active', 'x'].join('')].join('|'),
  'i'
);
const violations = [];
for (const file of files) {
  const lines = readFileSync(join(root, file), 'utf8').split('\n');
  lines.forEach((line, index) => {
    if (vendorProduct.test(line))
      violations.push(`${relative(root, join(root, file))}:${index + 1}`);
  });
}
if (violations.length) {
  console.error(`Public terminology check failed:\n${violations.join('\n')}`);
  process.exitCode = 1;
} else {
  console.info(`Public terminology check passed (${files.length} authored files).`);
}
