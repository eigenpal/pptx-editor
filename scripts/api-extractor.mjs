// Public API snapshots, as in docx-editor/scripts/api-extractor.mjs: one report per published entry
// point in docs/api/<slug>/<entry>.api.md. `--local` rewrites the snapshots; without it, any drift
// fails. Packages must be built first (`bun run build:packages`). DOCX Editor commit adf0e178.
import { Extractor, ExtractorConfig } from '@microsoft/api-extractor';
import { existsSync, mkdirSync } from 'node:fs';
import { resolve } from 'node:path';
import { PACKAGES } from './lib-packages.mjs';

const local = process.argv.includes('--local');
const only = process.argv.includes('--package')
  ? process.argv[process.argv.indexOf('--package') + 1]
  : null;
let failed = false;
for (const pkg of PACKAGES) {
  if (only && pkg.name !== only) continue;
  const root = resolve(pkg.root);
  for (const entry of pkg.entries) {
    const dts = resolve(root, 'dist', `${entry}.d.ts`);
    if (!existsSync(dts)) {
      console.error(`${pkg.name}: ${dts} is missing. Run bun run build:packages first.`);
      failed = true;
      continue;
    }
    const reportFolder = resolve('docs/api', pkg.slug);
    mkdirSync(reportFolder, { recursive: true });
    const config = ExtractorConfig.prepare({
      configObject: {
        projectFolder: root,
        mainEntryPointFilePath: dts,
        compiler: {
          overrideTsconfig: { compilerOptions: { skipLibCheck: true, lib: ['ES2022', 'DOM'] } },
        },
        apiReport: {
          enabled: true,
          reportFileName: `${entry}.api.md`,
          reportFolder,
          reportTempFolder: resolve('.cache/api-temp', pkg.slug),
        },
        docModel: { enabled: false },
        dtsRollup: { enabled: false },
        tsdocMetadata: { enabled: false },
        messages: {
          extractorMessageReporting: {
            default: { logLevel: 'none' },
            'ae-forgotten-export': { logLevel: 'none' },
          },
          compilerMessageReporting: { default: { logLevel: 'warning' } },
          tsdocMessageReporting: { default: { logLevel: 'none' } },
        },
      },
      configObjectFullPath: resolve(root, 'api-extractor.json'),
      packageJsonFullPath: resolve(root, 'package.json'),
    });
    const result = Extractor.invoke(config, { localBuild: local, showVerboseMessages: false });
    if (!result.succeeded) {
      console.error(`${pkg.name}/${entry}: API report drift. Run bun run api:extract and commit.`);
      failed = true;
    }
  }
}
process.exit(failed ? 1 : 0);
