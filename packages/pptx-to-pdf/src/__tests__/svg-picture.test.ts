/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pptx-to-pdf/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// SVG pictures in the PDF: an icon the editor draws as vector art draws as PDF paths, and any
// other SVG picture draws its PNG fallback with a `picture.svg-fallback` diagnostic.

import { describe, expect, test } from 'bun:test';
import { PNG } from 'pngjs';
import { getDocument, OPS } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { buildPptx } from '../../../../scripts/fixtures/builder.ts';
import { exportPdf } from '../index.ts';

const IN = 914_400;
const BEST = { fidelityPolicy: 'best-effort' } as const;
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const SVG_EXT =
  '<a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}">' +
  '<asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="rIdSvg"/>' +
  '</a:ext></a:extLst>';
const NS = 'xmlns="http://www.w3.org/2000/svg"';

function png(): Uint8Array {
  const image = new PNG({ width: 2, height: 2 });
  image.data.fill(255);
  return new Uint8Array(PNG.sync.write(image));
}

const deck = (svg: string, blipChildren = '', fill = '<a:stretch><a:fillRect/></a:stretch>') =>
  buildPptx({
    slides: [
      {
        layout: 'blank',
        shapes:
          `<p:pic><p:nvPicPr><p:cNvPr id="4" name="Picture"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>` +
          `<p:blipFill><a:blip r:embed="rIdImg">${blipChildren}${SVG_EXT}</a:blip>${fill}</p:blipFill>` +
          `<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${IN}" cy="${IN}"/></a:xfrm>` +
          `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`,
        rels: [
          ['rIdImg', `${R}/image`, '../media/img.png'],
          ['rIdSvg', `${R}/image`, '../media/img.svg'],
        ],
      },
    ],
    media: [
      { name: 'img.png', bytes: png(), contentType: 'image/png' },
      { name: 'img.svg', bytes: new TextEncoder().encode(svg), contentType: 'image/svg+xml' },
    ],
  });

async function operators(pdf: Uint8Array) {
  const doc = await getDocument({ data: pdf.slice(), useSystemFonts: false }).promise;
  return (await doc.getPage(1)).getOperatorList();
}

const ICON = `<svg ${NS} viewBox="0 0 96 96"><rect width="96" height="96" fill="#DEEBF7"/><circle cx="48" cy="48" r="36" fill="#2E75B6"/></svg>`;

describe('SVG pictures in the PDF', () => {
  test('an icon draws as vector paths, not its PNG fallback', async () => {
    const result = await exportPdf(deck(ICON), BEST);
    expect(result.diagnostics.find((d) => d.code === 'picture.svg-fallback')).toBeUndefined();
    const { fnArray, argsArray } = await operators(result.bytes);
    expect(fnArray).not.toContain(OPS.paintImageXObject);
    const colors = fnArray.flatMap((op, i) =>
      op === OPS.setFillRGBColor ? [argsArray[i] as unknown] : []
    );
    // pdf.js reports fill colors as hex.
    expect(colors.map(String)).toEqual(expect.arrayContaining(['#deebf7', '#2e75b6']));
  });

  test('an icon with opacity draws as a transparency group', async () => {
    const svg = `<svg ${NS} viewBox="0 0 10 10"><g opacity="0.5"><rect width="10" height="10"/><rect width="5" height="5" fill="red"/></g></svg>`;
    const result = await exportPdf(deck(svg), BEST);
    const text = new TextDecoder('latin1').decode(result.bytes);
    expect(text).toMatch(/\/S \/Transparency/);
    expect(text).toMatch(/\/ca 0\.5/);
  });

  const fallbacks: Record<string, [string, string?, string?]> = {
    'outside the subset (a gradient)': [
      `<svg ${NS} viewBox="0 0 4 4"><linearGradient id="g"/><rect width="4" height="4" fill="url(#g)"/></svg>`,
    ],
    'with an external reference': [
      `<svg ${NS} viewBox="0 0 4 4"><image href="http://example.com/a.png" width="4" height="4"/></svg>`,
    ],
    'with a script': [`<svg ${NS} viewBox="0 0 4 4"><script>alert(1)</script></svg>`],
    'nested past the XML depth limit': [
      `<svg ${NS} viewBox="0 0 4 4">${'<g>'.repeat(400)}${'</g>'.repeat(400)}</svg>`,
    ],
    'with color effects': [ICON, '<a:grayscl/>'],
    tiled: [ICON, '', '<a:tile tx="0" ty="0" sx="100000" sy="100000" flip="none" algn="tl"/>'],
  };
  for (const [name, [svg, blip, fill]] of Object.entries(fallbacks))
    test(`an SVG picture ${name} draws its PNG fallback and says so`, async () => {
      const result = await exportPdf(deck(svg, blip, fill), BEST);
      const fallback = result.diagnostics.find((d) => d.code === 'picture.svg-fallback');
      expect(fallback?.severity).toBe('information');
      expect(new TextDecoder('latin1').decode(result.bytes)).toContain('/Subtype /Image');
    });
});
