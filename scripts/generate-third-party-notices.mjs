// THIRD_PARTY_NOTICES.md per published package, from the bundler metafile: every package that
// tsup inlined into dist must ship its license text (docx-editor/scripts/generate-third-party-notices.mjs).
// `--check` compares against the current dist.
import { readFileSync, writeFileSync, existsSync, readdirSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { createRequire } from 'node:module';
import { PACKAGES } from './lib-packages.mjs';
import { VENDORED } from './vendored-notices.mjs';
import { validateHarfBuzzNotices } from './harfbuzz-notices.mjs';

const check = process.argv.includes('--check');
/** The installed directory of `dep` as `root` resolves it (Bun installs it per workspace). */
function packageDir(dep, root) {
  for (const candidate of [resolve(root, 'node_modules', dep), resolve('node_modules', dep)])
    if (existsSync(join(candidate, 'package.json'))) return candidate;
  let dir = dirname(createRequire(resolve(root, 'package.json')).resolve(dep));
  while (dir !== dirname(dir)) {
    const manifest = join(dir, 'package.json');
    if (existsSync(manifest) && JSON.parse(readFileSync(manifest, 'utf8')).name === dep) return dir;
    dir = dirname(dir);
  }
  return resolve('node_modules', dep);
}

let failed = false;
for (const pkg of PACKAGES) {
  const dist = resolve(pkg.root, 'dist');
  if (!existsSync(dist)) {
    console.error(`${pkg.name}: no dist. Build first.`);
    failed = true;
    continue;
  }
  if (pkg.name === '@pptx-editor.dev/core') validateHarfBuzzNotices();
  const inlined = new Set();
  const metafiles = readdirSync(dist).filter(
    (n) => n.startsWith('metafile-') && n.endsWith('.json')
  );
  // Without a metafile the inlined packages are unknown: fail, never claim "bundles no third-party code".
  if (metafiles.length === 0) {
    console.error(`${pkg.name}: no metafile in dist. Set \`metafile: true\` in tsup.config.ts.`);
    failed = true;
    continue;
  }
  for (const f of metafiles) {
    const meta = JSON.parse(readFileSync(join(dist, f), 'utf8'));
    for (const input of Object.keys(meta.inputs)) {
      const m = /node_modules\/(?:\.bun\/[^/]+\/node_modules\/)?((?:@[^/]+\/)?[^/]+)\//.exec(input);
      if (m) inlined.add(m[1]);
    }
  }
  let text = `# Third-party notices for ${pkg.name}\n\n`;
  const vendored = VENDORED[pkg.name] ?? [];
  if (inlined.size === 0 && vendored.length === 0)
    text += 'This package bundles no third-party code.\n';
  for (const entry of vendored) text += `${entry}\n`;
  for (const dep of [...inlined].sort()) {
    const base = packageDir(dep, pkg.root);
    const lic = ['LICENSE', 'LICENSE.md', 'LICENSE.txt', 'COPYING']
      .map((n) => join(base, n))
      .find(existsSync);
    if (!lic) {
      console.error(`${pkg.name}: inlined ${dep} has no license file`);
      failed = true;
      continue;
    }
    text += `## ${dep}\n\n\`\`\`\n${readFileSync(lic, 'utf8').trim()}\n\`\`\`\n\n`;
  }
  const out = resolve(pkg.root, 'THIRD_PARTY_NOTICES.md');
  if (check) {
    if (!existsSync(out) || readFileSync(out, 'utf8') !== text) {
      console.error(`${pkg.name}: notices missing or stale`);
      failed = true;
    }
  } else writeFileSync(out, text);
}
process.exit(failed ? 1 : 0);
