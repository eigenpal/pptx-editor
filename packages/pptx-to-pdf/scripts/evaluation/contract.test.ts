/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pptx-to-pdf/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// The contract that eigenpal/pptx-eval calls: every file and path it names exists, the protocol
// versions agree, and the Bun commands, run as the contract writes them, print and write what
// its schemas say. `test_contract.py` checks the Python commands the same way.

import { afterAll, describe, expect, test } from 'bun:test';
import { existsSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { openPresentation } from '@pptx-editor.dev/core/editor';
import { PROTOCOL } from './settings.ts';
import { THRESHOLDS_JSON, thresholdsJson } from './thresholds-json.ts';

const ROOT = join(import.meta.dir, '../../../..');
const FIXTURES = join(ROOT, 'fixtures');
const contract = JSON.parse(readFileSync(join(import.meta.dir, 'contract.json'), 'utf8')) as {
  protocol: number;
  requiredFiles: string[];
  identity: Record<string, { paths: string[] }>;
  commands: Record<
    string,
    { argv: string[]; optionalArgv?: string[][]; output?: string; stdout?: Record<string, string> }
  >;
  schemas: Record<string, Schema>;
};

interface Schema {
  type?: string | string[];
  const?: unknown;
  enum?: unknown[];
  required?: string[];
  properties?: Record<string, Schema>;
  items?: Schema;
  minItems?: number;
  maxItems?: number;
  minLength?: number;
  maxLength?: number;
  $ref?: string;
}

/** The JSON Schema subset the contract uses: type, const, enum, required, properties, items, lengths. */
function validate(schema: Schema | string, value: unknown, at = '$'): string[] {
  if (typeof schema === 'string' || schema.$ref) {
    const ref = typeof schema === 'string' ? schema : schema.$ref!;
    return validate(contract.schemas[ref.replace('#/schemas/', '')]!, value, at);
  }
  const kind =
    value === null
      ? 'null'
      : Array.isArray(value)
        ? 'array'
        : Number.isInteger(value)
          ? 'integer'
          : typeof value;
  const types = schema.type === undefined ? [] : [schema.type].flat();
  if (types.length && !types.some((t) => t === kind || (t === 'number' && kind === 'integer')))
    return [`${at}: expected ${types.join(' or ')}, got ${kind}`];
  if (value === null) return [];
  if ('const' in schema && JSON.stringify(schema.const) !== JSON.stringify(value))
    return [`${at}: expected ${JSON.stringify(schema.const)}`];
  if (schema.enum && !schema.enum.includes(value as never)) return [`${at}: not in enum`];
  const errors: string[] = [];
  if (typeof value === 'string') {
    if (schema.minLength !== undefined && value.length < schema.minLength)
      errors.push(`${at}: short`);
    if (schema.maxLength !== undefined && value.length > schema.maxLength)
      errors.push(`${at}: long`);
  } else if (Array.isArray(value)) {
    if (schema.minItems !== undefined && value.length < schema.minItems) errors.push(`${at}: few`);
    if (schema.maxItems !== undefined && value.length > schema.maxItems) errors.push(`${at}: many`);
    if (schema.items)
      for (const [i, item] of value.entries())
        errors.push(...validate(schema.items, item, `${at}[${i}]`));
  } else if (kind === 'object') {
    const record = value as Record<string, unknown>;
    for (const key of schema.required ?? [])
      if (!(key in record)) errors.push(`${at}.${key}: required`);
    for (const [key, sub] of Object.entries(schema.properties ?? {}))
      if (key in record) errors.push(...validate(sub, record[key], `${at}.${key}`));
  }
  return errors;
}

const folder = mkdtempSync(join(tmpdir(), 'pptx-eval-contract-'));
afterAll(() => rmSync(folder, { recursive: true, force: true }));

/** Run a contract command as the evaluator does: from the checkout root, one process each. */
function run(name: string, values: Record<string, string>, optional = false) {
  const command = contract.commands[name]!;
  const argv = [...command.argv, ...(optional ? (command.optionalArgv ?? []).flat() : [])].map(
    (arg) => arg.replace(/^\{(\w+)\}$/, (_, key: string) => values[key] ?? `{${key}}`)
  );
  expect(argv.some((arg) => arg.startsWith('{'))).toBe(false);
  const child = Bun.spawnSync([process.execPath, ...argv.slice(1)], {
    cwd: ROOT,
    stdout: 'pipe',
    stderr: 'pipe',
  });
  const lines = child.stdout.toString().trim().split('\n');
  expect(lines.length).toBe(1);
  return { exitCode: child.exitCode, report: JSON.parse(lines[0]!) as Record<string, unknown> };
}

describe('the evaluation contract', () => {
  test('thresholds.json is the JSON of scripts/fidelity/thresholds.ts', () => {
    expect(
      JSON.parse(readFileSync(THRESHOLDS_JSON, 'utf8')),
      'thresholds.json is out of date: run `bun packages/pptx-to-pdf/scripts/evaluation/thresholds-json.ts --write`'
    ).toEqual(JSON.parse(thresholdsJson()));
  });

  test('names files that exist and one protocol version', () => {
    expect(contract.protocol).toBe(PROTOCOL);
    const worker = readFileSync(join(import.meta.dir, 'worker.py'), 'utf8');
    expect(worker).toContain(`PROTOCOL = ${PROTOCOL}\n`);
    for (const file of contract.requiredFiles) expect(existsSync(join(ROOT, file))).toBe(true);
    for (const { paths } of Object.values(contract.identity))
      for (const path of paths) expect(existsSync(join(ROOT, path))).toBe(true);
    for (const { argv } of Object.values(contract.commands)) {
      const script = argv.find((arg) => arg.startsWith('packages/'));
      expect(contract.requiredFiles).toContain(script!);
    }
    // Every schema reference resolves.
    const refs = JSON.stringify(contract).match(/#\/schemas\/[\w-]+/g) ?? [];
    for (const ref of refs) expect(contract.schemas[ref.replace('#/schemas/', '')]).toBeDefined();
  });

  test('export prints the export report and writes a PDF', () => {
    const output = join(folder, 'text-runs.pdf');
    const { exitCode, report } = run('export', {
      input: join(FIXTURES, 'text-runs.pptx'),
      output,
    });
    expect(exitCode).toBe(0);
    expect(validate(contract.commands.export!.stdout!.exported!, report)).toEqual([]);
    expect(readFileSync(output).subarray(0, 5).toString()).toBe('%PDF-');
    expect(report.fontPlan).toBeNull();
  });

  test('a refused deck prints the error report', () => {
    const input = join(folder, 'broken.pptx');
    writeFileSync(input, 'PK\u0003\u0004 not a zip');
    const { exitCode, report } = run('export', { input, output: join(folder, 'broken.pdf') });
    expect(exitCode).toBe(1);
    expect(validate(contract.commands.export!.stdout!.error!, report)).toEqual([]);
    expect(report.code).toBe('documentOpenFailed');
    expect(existsSync(join(folder, 'broken.pdf'))).toBe(false);
  });

  test('the layout summary writes the summary schema', () => {
    const output = join(folder, 'tables.json');
    const { exitCode, report } = run('layoutSummary', {
      input: join(FIXTURES, 'tables.pptx'),
      output,
    });
    expect(exitCode).toBe(0);
    expect(validate(contract.commands.layoutSummary!.stdout!['laid-out']!, report)).toEqual([]);
    const summary = JSON.parse(readFileSync(output, 'utf8')) as unknown;
    expect(validate(contract.commands.layoutSummary!.output!, summary)).toEqual([]);
  });

  test('a font plan puts its faces first and reports its identity', async () => {
    // A packaged substitute face stands in for a reference font here.
    const opened = await openPresentation(
      new Uint8Array(readFileSync(join(FIXTURES, 'text-runs.pptx')))
    );
    if (!opened.ok) throw new Error('fixture');
    const face = opened.session.faces.find((f) => f.family === 'Carlito' && !f.bold && !f.italic)!;
    writeFileSync(join(folder, 'face.ttf'), face.bytes);
    const plan = {
      protocol: 1,
      faces: [{ path: 'face.ttf', family: 'Calibri', bold: false, italic: false }],
    };
    expect(validate('#/schemas/font-plan', plan)).toEqual([]);
    // The schema bounds a family name as the export does (64 characters).
    const long = { ...plan, faces: [{ ...plan.faces[0]!, family: 'F'.repeat(65) }] };
    expect(validate('#/schemas/font-plan', long)).toEqual(['$.faces[0].family: long']);
    writeFileSync(join(folder, 'plan.json'), JSON.stringify(plan));
    const values = {
      input: join(FIXTURES, 'text-runs.pptx'),
      output: join(folder, 'planned.pdf'),
      fontPlan: join(folder, 'plan.json'),
    };
    const { exitCode, report } = run('export', values, true);
    expect(exitCode).toBe(0);
    expect(validate(contract.commands.export!.stdout!.exported!, report)).toEqual([]);
    expect(report.fontPlan).toMatchObject({ faces: 1 });
    const resolution = report.fontResolution as { faces: { family: string; origin: string }[] };
    expect(resolution.faces).toContainEqual({ family: 'Calibri', origin: 'plan' });
    // A collection or a non-font file is refused with a message, never measured.
    writeFileSync(join(folder, 'face.ttf'), 'ttcf');
    const refused = run('export', values, true);
    expect(refused.exitCode).toBe(1);
    expect(String(refused.report.error)).toContain('not a single TrueType or OpenType face');
  });
});
