#!/usr/bin/env node
// Published exports and declarations must be self-contained
// (docx-editor/scripts/check-package-artifacts.mjs).
import { execFileSync } from 'node:child_process';
import { existsSync, readFileSync, readdirSync, statSync } from 'node:fs';
import path from 'node:path';

const root = path.resolve(import.meta.dirname, '..');
// Discover publishable packages so new packages cannot bypass artifact checks.
const packageDirs = readdirSync(path.join(root, 'packages')).filter((directory) => {
  const manifest = path.join(root, 'packages', directory, 'package.json');
  return existsSync(manifest) && JSON.parse(readFileSync(manifest, 'utf8')).private !== true;
});
const errors = [];

// What a published artifact may name when it reaches for the engine: exactly the
// subpaths core's own export map declares. Deriving the set instead of listing it
// keeps this honest when core adds or retires a lane — a hand-written denylist ages
// into false positives on subpaths that have since become public.
const corePackageJson = JSON.parse(
  readFileSync(path.join(root, 'packages/core/package.json'), 'utf8')
);
const publicCoreSubpaths = new Set(
  Object.keys(corePackageJson.exports ?? {})
    .filter((subpath) => subpath.startsWith('./'))
    .map((subpath) => subpath.slice(2))
);

function filesBelow(directory) {
  if (!existsSync(directory)) return [];
  return readdirSync(directory).flatMap((entry) => {
    const absolute = path.join(directory, entry);
    return statSync(absolute).isDirectory() ? filesBelow(absolute) : [absolute];
  });
}

function exportTargets(value) {
  if (typeof value === 'string') return [value];
  if (!value || typeof value !== 'object') return [];
  return Object.values(value).flatMap(exportTargets);
}

for (const packageDir of packageDirs) {
  const packageRoot = path.join(root, 'packages', packageDir);
  const packageJson = JSON.parse(readFileSync(path.join(packageRoot, 'package.json'), 'utf8'));

  for (const [subpath, declaration] of Object.entries(packageJson.exports ?? {})) {
    for (const target of exportTargets(declaration)) {
      if (!target.startsWith('./')) continue;
      if (!existsSync(path.resolve(packageRoot, target))) {
        errors.push(`${packageJson.name} ${subpath} points to missing ${target}`);
      }
    }
  }

  const artifacts = ['dist', 'types'].flatMap((directory) =>
    filesBelow(path.join(packageRoot, directory))
  );
  for (const artifact of artifacts) {
    if (!/\.(?:[cm]?js|d\.[cm]?ts)$/.test(artifact)) continue;
    const content = readFileSync(artifact, 'utf8');
    const isDeclaration = /\.d\.[cm]?ts$/.test(artifact);
    if (
      isDeclaration &&
      (/(?:\.\.\/)+core\/src\//.test(content) || /\/packages\/[^/]+\/src\//.test(content))
    ) {
      errors.push(`${path.relative(root, artifact)} exposes a workspace-only source path`);
    }
    if (isDeclaration) {
      for (const [, request] of content.matchAll(
        /\b(?:from\s*|import\s*\(\s*)['"](\.{1,2}\/[^'"]+)['"]/g
      )) {
        const target = path.resolve(path.dirname(artifact), request);
        const candidates = [
          target,
          target.replace(/\.([cm]?)js$/, '.d.$1ts'),
          `${target}.d.ts`,
          path.join(target, 'index.d.ts'),
        ];
        if (!candidates.some(existsSync)) {
          errors.push(`${path.relative(root, artifact)} references missing declaration ${request}`);
        }
      }
    }

    // A subpath core does not export resolves to nothing on a consumer's disk, so the
    // failure lands at their `npm install`, not ours.
    if (packageDir !== 'core') {
      for (const [, subpath] of content.matchAll(/@pptx-editor\.dev\/core\/([\w./-]+)/g)) {
        if (!publicCoreSubpaths.has(subpath)) {
          errors.push(
            `${path.relative(root, artifact)} imports @pptx-editor.dev/core/${subpath}, which core does not export`
          );
        }
      }
    }

    if (!isDeclaration && artifact.endsWith('.js')) {
      for (const match of content.matchAll(/\brequire\(\s*['"](\.{1,2}\/[^'"]+)['"]\s*\)/g)) {
        const request = match[1];
        const target = path.resolve(path.dirname(artifact), request);
        const candidates = [target, `${target}.js`, path.join(target, 'index.js')];
        if (!candidates.some(existsSync)) {
          errors.push(
            `${path.relative(root, artifact)} requires missing private artifact ${request}`
          );
        }
      }
    }
  }
}

// What npm would publish. npm skips a missing `files` entry without an error, so a lost license
// or notice file would publish silently: read the tarball list and require them.
for (const packageDir of packageDirs) {
  const packageRoot = path.join(root, 'packages', packageDir);
  const packageJson = JSON.parse(readFileSync(path.join(packageRoot, 'package.json'), 'utf8'));
  const name = packageJson.name;
  const [packed] = JSON.parse(
    execFileSync('npm', ['pack', '--dry-run', '--json', '--ignore-scripts'], {
      cwd: packageRoot,
      encoding: 'utf8',
      stdio: ['ignore', 'pipe', 'pipe'],
      // npm is `npm.cmd` on Windows, which only a shell resolves.
      shell: process.platform === 'win32',
    })
  );
  const files = new Set(packed.files.map((file) => file.path));
  const proprietary = String(packageJson.license).startsWith('LicenseRef-');
  const required = ['README.md', 'THIRD_PARTY_NOTICES.md', proprietary ? 'LICENSE.md' : 'LICENSE'];
  // The component notices of the bundled HarfBuzz binary.
  if (existsSync(path.join(packageRoot, 'licenses'))) {
    for (const file of readdirSync(path.join(packageRoot, 'licenses')))
      required.push(`licenses/${file}`);
  }
  for (const file of required) {
    if (!files.has(file)) errors.push(`${name} would publish without ${file}`);
  }
  for (const file of files) {
    if (/\.map$|(?:^|\/)metafile-[^/]*\.json$/.test(file))
      errors.push(`${name} would publish ${file}, which holds source or build paths`);
  }

  const scripts = [...files].filter((file) => /^dist\/.*\.[cm]?js$/.test(file));
  const declarations = [...files].filter((file) => /^dist\/.*\.d\.[cm]?ts$/.test(file));
  // A home directory or CI workspace path is a build machine path, wherever it was checked out.
  const machinePath = /(?:^|[\s'"`(])\/(?:Users|home|runner|private\/var|tmp)\/[\w.-]+\//;
  for (const file of [...scripts, ...declarations]) {
    const content = readFileSync(path.join(packageRoot, file), 'utf8');
    if (content.includes(root) || machinePath.test(content))
      errors.push(`${name} ${file} contains a build machine path`);
  }
  for (const file of [...scripts, ...declarations]) {
    const content = readFileSync(path.join(packageRoot, file), 'utf8');
    // Every file of a proprietary package opens with its license header (license-header.txt).
    if (proprietary && !/^\/\*!\s*Copyright \(c\) \d{4} EigenPal, Inc\./.test(content))
      errors.push(`${name} ${file} does not start with the license header`);
  }
  // The Ecma notice travels with the bundled preset geometry, inside the code itself.
  if (name === '@pptx-editor.dev/core') {
    const geometry = scripts.filter((file) =>
      readFileSync(path.join(packageRoot, file), 'utf8').includes('PRESET_SHAPES')
    );
    for (const file of geometry) {
      const content = readFileSync(path.join(packageRoot, file), 'utf8');
      // The whole notice inside one legal comment: the match may not cross a `*/`.
      if (!/\/\*!(?:(?!\*\/)[^])*?© Ecma International(?:(?!\*\/)[^])*?\*\//.test(content))
        errors.push(`${name} ${file} bundles the preset geometry without the Ecma notice`);
    }
    if (geometry.length === 0) errors.push(`${name} dist has no preset geometry table`);
  }
}

if (errors.length > 0) {
  console.error(errors.join('\n'));
  process.exit(1);
}

console.log(
  'Published package exports and declarations are self-contained, and every tarball carries its license and notices.'
);
