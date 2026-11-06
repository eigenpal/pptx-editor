/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Every refusal reason the core can return through an automation host.
//
// It follows the relative imports of the two automation hosts (`core/src/automation/index.ts`
// and `core/src/editor/automation-host.ts`) through `core/src`, and collects:
// - each `new CommandRefusal('<reason>', ...)`: an edit refusal (`disabled.*`, `errors.*`), which
//   the host passes on as its detail;
// - each `new AutomationFailure('<code>', <message>, '<detail>')` and `automationError(...)` with a
//   literal detail;
// - each `badArgument('<name>')`, an argument refusal named by its field.
//
// Run `node packages/editor-api/scripts/refusal-reasons.mjs` to print the list as JSON.
// `scripts/__tests__/refusal-reasons.test.ts` checks that each one maps to a diagnostic.

import { readFileSync } from 'node:fs';
import { dirname, relative, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const CORE = resolve(dirname(fileURLToPath(import.meta.url)), '../../core/src');
const ENTRIES = ['automation/index.ts', 'editor/automation-host.ts'];

/** The arguments of the call whose `(` is at `open`, split at top-level commas. */
function argumentsAt(source, open) {
  const args = [];
  let depth = 0;
  let start = open + 1;
  let quote = null;
  for (let i = open; i < source.length; i += 1) {
    const c = source[i];
    if (quote) {
      if (c === '\\') i += 1;
      else if (c === quote) quote = null;
      continue;
    }
    if (c === "'" || c === '"' || c === '`') quote = c;
    else if (c === '(' || c === '[' || c === '{') depth += 1;
    else if (c === ')' || c === ']' || c === '}') {
      depth -= 1;
      if (depth === 0) {
        args.push(source.slice(start, i).trim());
        return args;
      }
    } else if (c === ',' && depth === 1) {
      args.push(source.slice(start, i).trim());
      start = i + 1;
    }
  }
  return args;
}

/** The text of a plain string literal, or undefined for anything else. */
function literal(arg) {
  const m = /^(['"])((?:[^'"\\]|\\.)*)\1$/.exec(arg ?? '');
  return m ? m[2] : undefined;
}

/** The files the hosts reach through relative imports. */
function reachable() {
  const seen = new Set();
  const stack = ENTRIES.map((e) => resolve(CORE, e));
  while (stack.length > 0) {
    const file = stack.pop();
    if (seen.has(file)) continue;
    seen.add(file);
    const source = readFileSync(file, 'utf8');
    for (const m of source.matchAll(/(?:from|import)\s*\(?\s*'(\.{1,2}\/[^']+\.ts)'/g)) {
      const target = resolve(dirname(file), m[1]);
      if (target.startsWith(CORE) && !target.includes('__tests__')) stack.push(target);
    }
  }
  return [...seen].sort();
}

/**
 * Each reason with its kind and the files that give it. `dynamic` lists the call sites whose
 * reason is not a literal, so a reader can check them by hand.
 */
export function refusalReasons() {
  const reasons = new Map();
  const dynamic = [];
  const add = (kind, value, file, code) => {
    const key = `${kind}:${value}`;
    const entry = reasons.get(key) ?? { kind, reason: value, codes: new Set(), files: new Set() };
    entry.files.add(relative(CORE, file));
    if (code) entry.codes.add(code);
    reasons.set(key, entry);
  };
  for (const file of reachable()) {
    const source = readFileSync(file, 'utf8');
    const calls = [
      [/new CommandRefusal\(/g, 'refusal', 0, -1],
      [/new AutomationFailure\(/g, 'failure', 2, 0],
      [/\bautomationError\(/g, 'failure', 2, 0],
      [/\bbadArgument\(/g, 'argument', 0, -1],
    ];
    for (const [pattern, kind, at, codeAt] of calls)
      for (const m of source.matchAll(pattern)) {
        if (/function\s+$/.test(source.slice(Math.max(0, m.index - 12), m.index))) continue;
        const args = argumentsAt(source, m.index + m[0].length - 1);
        if (kind === 'failure' && args.length <= at) continue;
        const code = codeAt >= 0 ? literal(args[codeAt]) : undefined;
        // A refusal reason chosen by a condition: every reason literal of the expression.
        const values =
          kind === 'refusal' && literal(args[at]) === undefined
            ? [...(args[at] ?? '').matchAll(/'((?:disabled|errors)\.[A-Za-z]+)'/g)].map((x) => x[1])
            : [literal(args[at])].filter((v) => v !== undefined);
        if (values.length === 0)
          dynamic.push(`${relative(CORE, file)}:${source.slice(0, m.index).split('\n').length}`);
        for (const value of values)
          add(kind, value, file, kind === 'argument' ? 'invalid-argument' : code);
      }
  }
  return {
    reasons: [...reasons.values()]
      .map((r) => ({ ...r, codes: [...r.codes].sort(), files: [...r.files].sort() }))
      .sort((a, b) => (a.kind + a.reason).localeCompare(b.kind + b.reason)),
    dynamic,
  };
}

if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url))
  process.stdout.write(`${JSON.stringify(refusalReasons(), null, 2)}\n`);
