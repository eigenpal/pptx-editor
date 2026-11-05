import { afterEach, beforeEach, describe, expect, test } from 'bun:test';
import { copyFileSync, mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { missingReport } from './missing-references.ts';

const FIXTURE = join(import.meta.dir, '../../fixtures/colors.pptx');

let dir: string;
beforeEach(() => {
  dir = mkdtempSync(join(tmpdir(), 'pptx-missing-'));
  copyFileSync(FIXTURE, join(dir, 'colors.pptx'));
});
afterEach(() => rmSync(dir, { recursive: true, force: true }));

describe('fidelity:missing', () => {
  test('lists a missing PDF with the next step; --check exits 1', () => {
    const report = missingReport([dir], { check: true, names: false, gated: dir });
    expect(report.lines).toEqual([
      `${dir}: 1 without a reference PDF, 0 stale`,
      '  missing  colors.pptx',
      `  Next: add the reference PDF of each deck next to it, and its entry in ${dir}/references.json`,
    ]);
    expect(report.exitCode).toBe(1);
    expect(missingReport([dir], { check: false, names: false, gated: dir }).exitCode).toBe(0);
  });

  test('a PDF with no recorded identity is stale', () => {
    writeFileSync(join(dir, 'colors.pdf'), '%PDF-1.4');
    const report = missingReport([dir], { check: true, names: false, gated: dir });
    expect(report.lines[1]).toBe('  stale    colors.pptx: no identity recorded for the PDF');
    expect(report.exitCode).toBe(1);
  });

  test('a folder that does not exist prints nothing', () => {
    const report = missingReport([join(dir, 'none')], { check: true, names: false });
    expect(report).toEqual({ lines: [], exitCode: 0 });
  });

  test('the committed fixtures have current references', () => {
    const fixtures = join(import.meta.dir, '../../fixtures');
    const report = missingReport([fixtures], { check: true, names: false, gated: fixtures });
    expect(report.exitCode).toBe(0);
  });
});
