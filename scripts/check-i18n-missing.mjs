// Every i18n key that shipping code names must exist in packages/i18n/en.json.
//
// `TranslationKey` checks the keys TypeScript can see. This check covers the rest: keys that
// travel as plain strings (engine refusal reasons such as `disabled.viewing`, registry labels
// in data) and key prefixes built at runtime (`shapes.${preset}`). A missing key would show
// as its raw id in the chrome.
//
// A literal counts as a key when it starts with a top-level section of en.json (`toolbar.`,
// `errors.`, …). A literal prefix followed by an interpolation or a concatenation must name a
// section of en.json. Comments and tests are skipped.
import { readdirSync, readFileSync, statSync } from 'node:fs';
import { join, relative } from 'node:path';

const en = JSON.parse(readFileSync('packages/i18n/en.json', 'utf8'));
const ROOTS = ['packages', 'examples/vite/src'];
const SKIP_DIRS = new Set(['node_modules', 'dist', '.git']);
const SOURCE = /\.(ts|tsx|mjs)$/;
const TEST = /(^|\/)(__tests__|test|tests)\/|\.(test|spec)\./;
const COMMENT = /\/\*[\s\S]*?\*\/|^[ \t]*\/\/.*$/gm;

/** The node at a dotted path, or undefined. */
function at(path) {
  let node = en;
  for (const part of path.split('.')) {
    if (node === null || typeof node !== 'object' || !Object.hasOwn(node, part)) return undefined;
    node = node[part];
  }
  return node;
}

function files(dir, out = []) {
  for (const name of readdirSync(dir)) {
    if (SKIP_DIRS.has(name)) continue;
    const path = join(dir, name);
    if (statSync(path).isDirectory()) files(path, out);
    else if (SOURCE.test(name) && !TEST.test(relative('.', path))) out.push(path);
  }
  return out;
}

const sections = Object.keys(en).join('|');
const KEY = new RegExp(`(['"\`])((?:${sections})\\.[A-Za-z0-9_.]*[A-Za-z0-9_])\\1`, 'g');
const PREFIX = new RegExp(
  `[\`'"]((?:${sections})(?:\\.[A-Za-z0-9_]+)*)\\.(?:\\$\\{|['"]\\s*\\+)`,
  'g'
);

const missing = [];
for (const file of ROOTS.flatMap((root) => files(root))) {
  const text = readFileSync(file, 'utf8').replace(COMMENT, '');
  for (const m of text.matchAll(KEY))
    if (typeof at(m[2]) !== 'string') missing.push(`${relative('.', file)}: ${m[2]}`);
  for (const m of text.matchAll(PREFIX)) {
    const node = at(m[1]);
    if (node === undefined || typeof node !== 'object')
      missing.push(`${relative('.', file)}: ${m[1]}.*`);
  }
}

if (missing.length) {
  console.error(`i18n: ${missing.length} keys used in code are missing from en.json:`);
  for (const line of [...new Set(missing)]) console.error(`  ${line}`);
  process.exit(1);
}
console.log('i18n: every key used in code is in en.json');
