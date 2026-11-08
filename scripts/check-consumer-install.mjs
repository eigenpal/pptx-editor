#!/usr/bin/env node
//
// Installs the packages the way a consumer does: from tarballs, into an empty project, with npm
// resolving what it finds inside them. Then it exports a PDF under Node, builds a React app with
// Vite and with Bun's bundler, and opens the Vite build in Chromium to check that a slide renders
// (docx-editor/scripts/check-consumer-install.mjs, sized for PPTX Editor). DOCX Editor commit
// adf0e178.
//
// This is the only check that reads a published manifest instead of a workspace one, so it
// is the only place where a `workspace:` range, a missing `exports` subpath, a `files` list
// that drops a needed file, or a WASM asset that a bundler cannot find (for example,
// when `harfbuzzjs` stays external) fails before a user sees it.
//
// Needs built packages (`bun run build:packages`) and network access to the npm registry.
// Usage: node scripts/check-consumer-install.mjs [--keep]
import {
  cpSync,
  existsSync,
  mkdtempSync,
  mkdirSync,
  readFileSync,
  readdirSync,
  rmSync,
  writeFileSync,
} from 'node:fs';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { spawn, spawnSync } from 'node:child_process';
import { PACKAGES } from './lib-packages.mjs';

const ROOT = path.resolve(import.meta.dirname, '..');
const keep = process.argv.includes('--keep');
const tempRoot = mkdtempSync(path.join(tmpdir(), 'pptx-editor-consumer-'));
const stageDir = path.join(tempRoot, 'staged');
const packDir = path.join(tempRoot, 'packs');
const appDir = path.join(tempRoot, 'app');

function run(command, args, cwd) {
  const result = spawnSync(command, args, { cwd, encoding: 'utf8', stdio: 'pipe' });
  if (result.status !== 0) {
    process.stderr.write(result.stdout ?? '');
    process.stderr.write(result.stderr ?? '');
    throw new Error(`${command} ${args.join(' ')} failed with exit code ${result.status}`);
  }
  return result.stdout;
}

const rootManifest = JSON.parse(readFileSync(path.join(ROOT, 'package.json'), 'utf8'));
// Stage each package as it would be published. `changeset publish` runs `npm publish`, which
// copies a `workspace:` range into the tarball unchanged, so such a range fails here instead
// of being rewritten.
const tarballs = new Map();
mkdirSync(packDir, { recursive: true });
for (const pkg of PACKAGES) {
  const source = path.join(ROOT, pkg.root);
  if (!existsSync(path.join(source, 'dist'))) throw new Error(`${pkg.name}: build first`);
  const target = path.join(stageDir, pkg.slug);
  cpSync(source, target, {
    recursive: true,
    filter: (file) => !file.includes(`${path.sep}node_modules`),
  });
  const manifest = JSON.parse(readFileSync(path.join(target, 'package.json'), 'utf8'));
  for (const field of ['dependencies', 'peerDependencies', 'optionalDependencies']) {
    for (const [name, range] of Object.entries(manifest[field] ?? {}))
      if (String(range).startsWith('workspace:'))
        throw new Error(`${pkg.name}: ${field} lists ${name}@${range}; npm publishes it as it is`);
  }
  delete manifest.devDependencies;
  writeFileSync(path.join(target, 'package.json'), `${JSON.stringify(manifest, null, 2)}\n`);
  const file = run('npm', ['pack', '--silent', '--pack-destination', packDir], target).trim();
  tarballs.set(pkg.name, path.join(packDir, file.split('\n').pop()));
}

// A minimal consumer app, written the way README.md tells users to write it.
const dep = (name) => rootManifest.devDependencies[name];
mkdirSync(path.join(appDir, 'src'), { recursive: true });
mkdirSync(path.join(appDir, 'public'), { recursive: true });
writeFileSync(
  path.join(appDir, 'package.json'),
  JSON.stringify(
    {
      name: 'pptx-editor-consumer',
      private: true,
      type: 'module',
      dependencies: {
        ...Object.fromEntries([...tarballs].map(([name, file]) => [name, `file:${file}`])),
        react: dep('react'),
        'react-dom': dep('react-dom'),
      },
      devDependencies: { vite: dep('vite'), '@vitejs/plugin-react': dep('@vitejs/plugin-react') },
    },
    null,
    2
  )
);
writeFileSync(
  path.join(appDir, 'vite.config.js'),
  "import { defineConfig } from 'vite';\nimport react from '@vitejs/plugin-react';\n\nexport default defineConfig({ plugins: [react()] });\n"
);
writeFileSync(
  path.join(appDir, 'index.html'),
  '<!doctype html><html><head><meta charset="UTF-8" /></head><body><div id="app" style="height:100vh"></div><script type="module" src="/src/main.jsx"></script></body></html>\n'
);
writeFileSync(
  path.join(appDir, 'src', 'main.jsx'),
  `import { createRoot } from 'react-dom/client';
import { PptxEditor } from '@pptx-editor.dev/react';
import '@pptx-editor.dev/core/styles/editor.css';

const bytes = new Uint8Array(await (await fetch('/sample.pptx')).arrayBuffer());
createRoot(document.getElementById('app')).render(<PptxEditor document={bytes} mode="edit" />);
`
);
cpSync(
  path.join(ROOT, 'examples/vite/public/sample.pptx'),
  path.join(appDir, 'public/sample.pptx')
);
cpSync(path.join(ROOT, 'examples/vite/public/showcase.pptx'), path.join(appDir, 'showcase.pptx'));
// PDF export under Node, as the pptx-to-pdf README tells users to write it. Every unit test
// runs under Bun, whose `fetch` reads `file:` URLs; Node's does not, so only this run proves
// that the packaged substitute fonts load and the PDF carries its text.
writeFileSync(
  path.join(appDir, 'export-pdf.mjs'),
  `import { readFile } from 'node:fs/promises';
import { exportPdf } from '@pptx-editor.dev/pptx-to-pdf';

if (process.versions.bun) throw new Error('run this under Node, not Bun');
const result = await exportPdf(await readFile('showcase.pptx'), { fidelityPolicy: 'best-effort' });
const pdf = Buffer.from(result.bytes).toString('latin1');
console.log(
  JSON.stringify({
    pages: result.pageCount,
    embeddedFonts: (pdf.match(/\\/FontFile2? /g) ?? []).length,
    substituteFaces: result.fontResolution.faces.filter((f) => f.origin === 'substitute').length,
    substituteFailures: result.fontResolution.substituteFailures,
    fontDiagnostics: result.diagnostics
      .filter((d) => d.code === 'font.missing' || d.code === 'font.substitute-failed')
      .map((d) => d.message),
  })
);
`
);

const failures = [];
try {
  run('npm', ['install', '--no-audit', '--no-fund', '--loglevel=error'], appDir);

  // Node: the PDF must embed the packaged substitute fonts and draw the deck's text.
  checkNodePdf(JSON.parse(run('node', ['export-pdf.mjs'], appDir)), failures);

  // Vite: the build must emit HarfBuzz's WASM, and the app must render a slide.
  run('npx', ['vite', 'build', '--logLevel', 'warn'], appDir);
  const assets = readdirSync(path.join(appDir, 'dist', 'assets'));
  if (!assets.some((name) => name.endsWith('.wasm')))
    failures.push('vite: the build emitted no .wasm asset (harfbuzzjs)');
  await checkRender(failures);

  // Bun: the browser bundle must build. Bun does not emit `new URL(..., import.meta.url)`
  // assets, so a Bun app serves `@pptx-editor.dev/core/harfbuzz.wasm` itself and calls
  // `setHarfBuzzWasmUrl` (README.md); the export must exist for that.
  const bun = spawnSync(
    'bun',
    ['build', 'src/main.jsx', '--outdir', 'dist-bun', '--target', 'browser', '--splitting'],
    { cwd: appDir, encoding: 'utf8' }
  );
  if (bun.status !== 0) failures.push(`bun build failed: ${bun.stderr.trim().split('\n')[0]}`);
  const wasm = path.join(appDir, 'node_modules/@pptx-editor.dev/core/dist/harfbuzz.wasm');
  if (!existsSync(wasm)) failures.push('the core package does not ship dist/harfbuzz.wasm');
} finally {
  if (keep) console.log(`Kept ${tempRoot}`);
  else rmSync(tempRoot, { recursive: true, force: true });
}

function checkNodePdf(report, out) {
  if (report.pages < 1) out.push('node pdf: the export has no pages');
  if (report.embeddedFonts < 1) out.push('node pdf: the PDF embeds no font, so it has no text');
  if (report.substituteFaces < 1) out.push('node pdf: no packaged substitute font loaded');
  for (const f of report.substituteFailures)
    out.push(
      `node pdf: a substitute font failed to load: ${f.family} (${f.file}): ${f.diagnostic}`
    );
  for (const message of report.fontDiagnostics) out.push(`node pdf: ${message}`);
}

async function checkRender(out) {
  // Vite's own bin in its own process group. Through `npx`, killing the child on Linux left
  // the Vite server running, and its open stdout pipe kept this script alive until CI's timeout.
  const vite = path.join(appDir, 'node_modules/vite/bin/vite.js');
  const preview = spawn(process.execPath, [vite, 'preview', '--port', '4319', '--strictPort'], {
    cwd: appDir,
    stdio: 'pipe',
    detached: true,
  });
  try {
    await new Promise((resolve, reject) => {
      const timer = setTimeout(() => reject(new Error('vite preview did not start')), 30_000);
      preview.stdout.on('data', (chunk) => {
        if (String(chunk).includes('4319')) {
          clearTimeout(timer);
          resolve();
        }
      });
    });
    const { chromium } = await import('@playwright/test');
    const browser = await chromium.launch();
    try {
      const page = await browser.newPage();
      const errors = [];
      page.on('pageerror', (error) => errors.push(error.message));
      page.on('console', (message) => {
        if (message.type() === 'error') errors.push(message.text());
      });
      await page.goto('http://localhost:4319/');
      try {
        await page.waitForSelector('.pptx-editor svg text', { timeout: 30_000 });
      } catch {
        out.push('vite: no slide text rendered within 30 s');
      }
      for (const error of errors) out.push(`vite: console error: ${error}`);
    } finally {
      await browser.close();
    }
  } finally {
    // The negative pid signals the whole group: the server and anything it started.
    try {
      process.kill(-preview.pid, 'SIGTERM');
    } catch {
      // The group has already exited.
    }
  }
}

if (failures.length > 0) {
  console.error(failures.join('\n'));
  process.exit(1);
}
console.log(
  'Consumer install: the Node PDF export embeds its fonts, the Vite app builds and renders, and the Bun bundle builds.'
);
