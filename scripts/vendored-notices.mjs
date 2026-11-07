// Notices for third-party data copied into source files. npm packages are found from the
// bundler metafile; these are not, so they are listed here by hand. `notices-sources.test.ts`
// fails when a source note in `packages/*/src` names a third-party origin that is not listed.
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { HARFBUZZ_NOTICE } from './harfbuzz-notices.mjs';

// The Ecma copyright notice must travel with the generated geometry data (reference/ecma-376/NOTICE).
const ecmaFile = readFileSync(resolve(import.meta.dirname, '../reference/ecma-376/NOTICE'), 'utf8');
const ECMA_NOTICE = ecmaFile.slice(ecmaFile.indexOf('COPYRIGHT NOTICE')).trim();
// Data copied into source files (not npm packages, so the metafile cannot see it).
export const VENDORED = {
  '@pptx-editor.dev/core': [
    "## Material Symbols\n\n`src/editor/icon-paths.ts` holds SVG path data from Material Symbols by Google, licensed under the Apache License, Version 2.0 (https://github.com/google/material-design-icons). The license text is in this package's `LICENSE` file.\n",
    '## ECMA-376 preset geometries\n\nThe preset geometry and text warp tables, which this package inlines from `@eigenpal/ooxml-core` (`drawingml-preset-geometry.generated.ts` and `drawingml-preset-text-warp.generated.ts`) at build time, are generated from `presetShapeDefinitions.xml` and `presetTextWarpDefinitions.xml` in ECMA-376 Part 1, 5th edition (2016), published by Ecma International (https://ecma-international.org/publications-and-standards/standards/ecma-376/). They are used under the Ecma copyright notice below.\n\n```\n' +
      ECMA_NOTICE +
      '\n```\n',
    HARFBUZZ_NOTICE,
    '## MicroType Express\n\n`src/store/fonts/` decodes embedded fonts in the MicroType Express (MTX) format described in the W3C Member Submission "MicroType Express (MTX) Font Format" (2008), https://www.w3.org/Submission/MTX/. Monotype Imaging grants a royalty-free license to implement the format, and a patent license, in its MTX licensing statement sent on 5 June 2014 (https://www.w3.org/standards/licensing/). The statement covers the format as of 28 December 2011. The decoder is an independent implementation; it contains no Monotype code.\n',
  ],
};
