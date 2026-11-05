// The security grep from AGENTS.md "Security". Zero hits outside tests is the gate.
import { readdirSync, readFileSync, statSync } from 'node:fs';
import { join } from 'node:path';

const PATTERN =
  /innerHTML|outerHTML|insertAdjacentHTML|dangerouslySetInnerHTML|v-html|document\.write|window\.open\(|\.href\s*=|font-family:.*\$\{|new Function\(|\beval\(/;
const hits = [];
function walk(dir) {
  for (const name of readdirSync(dir)) {
    const p = join(dir, name);
    if (name === 'node_modules' || name === 'dist' || name === '__tests__') continue;
    if (statSync(p).isDirectory()) walk(p);
    else if (
      /\.(ts|tsx|vue)$/.test(name) &&
      !/\.(test|spec)\./.test(name) &&
      !name.endsWith('.generated.ts')
    ) {
      readFileSync(p, 'utf8')
        .split('\n')
        .forEach((line, i) => {
          if (PATTERN.test(line)) hits.push(`${p}:${i + 1}: ${line.trim()}`);
        });
    }
  }
}
walk('packages');
walk('examples/vite/src');
walk('examples/vue/src');
walk('examples/chat/src');
if (hits.length) {
  console.error(`Security grep: ${hits.length} hit(s)\n${hits.join('\n')}`);
  process.exit(1);
}
console.log('Security grep: 0 hits');
