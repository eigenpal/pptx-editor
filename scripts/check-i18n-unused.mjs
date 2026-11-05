// Lists en.json keys that no source file references. Unused keys fail the gate, so the
// catalog never drifts from the chrome. `--list` prints without failing.
import { readFileSync, readdirSync, statSync } from 'node:fs';
import { join } from 'node:path';

const en = JSON.parse(readFileSync('packages/i18n/en.json', 'utf8'));
const keys = [];
(function flatten(o, p = '') {
  for (const [k, v] of Object.entries(o))
    typeof v === 'string' ? keys.push(`${p}${k}`) : flatten(v, `${p}${k}.`);
})(en);
let source = '';
(function walk(dir) {
  for (const n of readdirSync(dir)) {
    const p = join(dir, n);
    if (n === 'node_modules' || n === 'dist') continue;
    if (statSync(p).isDirectory()) walk(p);
    else if (/\.(ts|tsx)$/.test(n)) source += readFileSync(p, 'utf8');
  }
})('packages');
(function walk(dir) {
  for (const n of readdirSync(dir)) {
    const p = join(dir, n);
    if (n === 'node_modules' || n === 'dist') continue;
    if (statSync(p).isDirectory()) walk(p);
    else if (/\.(ts|tsx)$/.test(n)) source += readFileSync(p, 'utf8');
  }
})('examples/vite/src');
// A key counts as used when its full path appears, or when a dynamic prefix (`'panel.tabs.' +`)
// covers it.
const unused = keys.filter((k) => {
  if (source.includes(`'${k}'`) || source.includes(`"${k}"`) || source.includes(`\`${k}\``))
    return false;
  const parts = k.split('.');
  for (let i = parts.length - 1; i > 0; i -= 1) {
    const prefix = `${parts.slice(0, i).join('.')}.`;
    if (source.includes(`'${prefix}`) || source.includes(`\`${prefix}`)) return false;
  }
  return true;
});
if (unused.length) console.log(`Unused i18n keys (${unused.length}):\n  ${unused.join('\n  ')}`);
else console.log('i18n: no unused keys');
process.exit(unused.length && !process.argv.includes('--list') ? 1 : 0);
