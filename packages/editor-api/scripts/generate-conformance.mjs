#!/usr/bin/env node
// Copied from DOCX Editor packages/editor-api/scripts/generate-conformance.mjs at commit 5993bd223945cabdaf3900557d9fb59bfd8c2d8c.
// Changes: the presentation namespace, PPTX Editor names, and paths. DOCX Editor compares the
// reference with a hand-written declarations file (compat/docxeditor/declarations.ts) through
// extract-docxeditor-shape.mjs; PPTX Editor has no such file, so that module is not copied and
// no shape JSON is written. The authored side of each assertion is the real public type, read
// from src/index.ts by the compiler (`Pick<PptxEditor.Shape, 'fill'>`,
// `PptxEditor.ShapeCollection['addTextBox']`). The reference side comes from
// reference/presentation.full-inventory.json, which keeps the enum names that the subset
// reference collapses to string literals. One assertion covers all the overloads of a member,
// in order; a property assertion also covers `readonly` and `?`; a writable property also gets
// a write check. `manifest.conformanceExceptions` names the selected members that cannot hold,
// with the reason. `--check` fails when the committed file differs from a fresh generation.
/**
 * The conformance generator.
 *
 * Reads the manifest selection and the pinned reference inventory, and emits
 * `compat/generated/conformance.generated.ts`: strict TypeScript compile assertions, one for
 * each selected member, against the public types of `src/index.ts`. The output depends only on
 * committed compat data, so a change to the public API never changes the file. The compiler,
 * not this script, finds the drift: `compat/tsconfig.json` type-checks the generated file in
 * `scripts/__tests__/conformance.test.ts`.
 *
 * This is a pure, offline transformation. It never fetches anything over the network.
 */

import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import ts from 'typescript';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
export const ASSERTIONS_PATH = path.join(root, 'compat', 'generated', 'conformance.generated.ts');

/** The namespace alias the generated file imports the package as. */
const NS = 'PptxEditor';
/** Type names that the lib declares. Every other type reference names a package export. */
const GLOBAL_TYPES = new Set(['Array', 'Date', 'Promise', 'ReadonlyArray', 'Record']);

/** `PowerPoint.Shape#fill` → `{ owner: 'Shape', member: 'fill' }`; `PowerPoint.run` → `{ owner: 'run' }`. */
function localName(uid) {
  const [qualified, member] = uid.split('#');
  return { owner: qualified.replace(/^(PowerPoint|OfficeExtension)\./, ''), member };
}

function selectedUids(manifest) {
  const uids = [];
  for (const [name, selection] of Object.entries(manifest.symbols ?? {})) {
    const namespace = selection.namespace ?? 'PowerPoint';
    if (selection.isFunction) uids.push(`${namespace}.${name}`);
    else for (const member of selection.members ?? []) uids.push(`${namespace}.${name}#${member}`);
  }
  return uids;
}

/**
 * Re-qualifies each type reference in a normalized inventory type (`ShapeFill`,
 * `ShapeType`) as `PptxEditor.ShapeFill`, so the reference names the package's own export. A
 * name the package does not export then fails as a real "has no exported member" compiler
 * error. Type parameters and lib types stay bare. String literal contents never change, since
 * the rewrite works on the syntax tree, not on the text.
 */
export function qualifyTypeText(text, typeParameterNames = []) {
  const bare = new Set([...GLOBAL_TYPES, ...typeParameterNames]);
  const source = ts.createSourceFile('t.ts', `type X = ${text};`, ts.ScriptTarget.Latest, true);
  if (source.parseDiagnostics.length) throw new Error(`Cannot parse reference type: ${text}`);
  const result = ts.transform(source.statements[0].type, [
    (context) => {
      const visit = (node) => {
        const child = ts.visitEachChild(node, visit, context);
        if (
          ts.isTypeReferenceNode(child) &&
          ts.isIdentifier(child.typeName) &&
          !bare.has(child.typeName.text)
        ) {
          return ts.factory.updateTypeReferenceNode(
            child,
            ts.factory.createQualifiedName(ts.factory.createIdentifier(NS), child.typeName.text),
            child.typeArguments
          );
        }
        return child;
      };
      return visit;
    },
  ]);
  try {
    const printer = ts.createPrinter({ removeComments: true, newLine: ts.NewLineKind.LineFeed });
    return printer
      .printNode(ts.EmitHint.Unspecified, result.transformed[0], source)
      .replace(/\s*\n\s*/g, ' ');
  } finally {
    result.dispose();
  }
}

function typeParameterText(typeParameters, names) {
  if (!typeParameters?.length) return '';
  const parts = typeParameters.map(
    (parameter) =>
      parameter.name +
      (parameter.constraint ? ` extends ${qualifyTypeText(parameter.constraint, names)}` : '') +
      (parameter.default ? ` = ${qualifyTypeText(parameter.default, names)}` : '')
  );
  return `<${parts.join(', ')}>`;
}

/** Parameter names from the subset reference, when it lists the same overload; else `argN`. */
function parameterNames(subsetMember, index, count) {
  const names = subsetMember?.overloads?.[index]?.params?.map((param) => param.name) ?? [];
  return names.length === count ? names : Array.from({ length: count }, (_, i) => `arg${i}`);
}

/** One call signature: `<T>(name: Type, other?: Type): Return`. */
function signatureText(overload, names) {
  const typeNames = (overload.typeParameters ?? []).map((parameter) => parameter.name);
  const parameters = overload.parameters.map(
    (parameter, i) =>
      `${parameter.rest ? '...' : ''}${names[i]}${parameter.optional ? '?' : ''}: ${qualifyTypeText(parameter.type, typeNames)}`
  );
  return `${typeParameterText(overload.typeParameters, typeNames)}(${parameters.join(', ')}): ${qualifyTypeText(overload.returns, typeNames)}`;
}

function sanitizeIdentifier(text) {
  return text.replace(/[^A-Za-z0-9_]/g, '_');
}

function quoted(name) {
  return `'${name.replace(/\\/g, '\\\\').replace(/'/g, "\\'")}'`;
}

function emitAssertion(label, refText, authText) {
  // Two steps, not one `AssertExact<A, B>` combinator: see the `Expect` comment in
  // compat/type-assert.ts.
  return [
    `type Ref_${label} = ${refText};`,
    `type Auth_${label} = ${authText};`,
    `type _check_${label} = IsExact<Ref_${label}, Auth_${label}>;`,
    `type _assert_${label} = Expect<_check_${label}>;`,
  ];
}

/**
 * A method or function: all overloads, in reference order, as one call-signature type.
 * `IsExact` compares the overload lists, so a missing, extra, reordered, narrowed, or widened
 * overload fails.
 */
function emitCallable(uid, row, subsetMember) {
  const { owner, member } = localName(uid);
  const label = sanitizeIdentifier(member ? `${owner}_${member}` : owner);
  const signatures = row.overloads.map((overload, index) =>
    signatureText(overload, parameterNames(subsetMember, index, overload.parameters.length))
  );
  const refText = `{ ${signatures.join('; ')} }`;
  const authText = member ? `${NS}.${owner}[${quoted(member)}]` : `typeof ${NS}.${owner}`;
  return emitAssertion(label, refText, authText);
}

/**
 * A property: its read type, `readonly`, and `?` in one object type, against `Pick` of the
 * real class, which keeps the modifiers. A writable property also gets a write check: the
 * reference write type must assign to the real setter. TypeScript has no type operator for a
 * setter type, so the write check is one direction only: a setter that also accepts more than
 * the reference does not fail here. `compat:report` compares setter types exactly.
 */
function emitProperty(uid, row) {
  const { owner, member } = localName(uid);
  const label = sanitizeIdentifier(`${owner}_${member}`);
  const key = /^[A-Za-z_$][A-Za-z0-9_$]*$/.test(member) ? member : quoted(member);
  const refText = `{ ${row.readonly ? 'readonly ' : ''}${key}${row.optional ? '?' : ''}: ${qualifyTypeText(row.type)} }`;
  const lines = emitAssertion(label, refText, `Pick<${NS}.${owner}, ${quoted(member)}>`);
  if (!row.readonly && row.writeType) {
    lines.push(
      `export function _write_${label}(target: ${NS}.${owner}, value: ${qualifyTypeText(row.writeType)}): void {`,
      `  target[${quoted(member)}] = value;`,
      '}'
    );
  }
  return lines;
}

/** A reason as `//` lines of at most 100 characters. */
function wrapComment(text) {
  const lines = [];
  let line = '//';
  for (const word of text.split(/\s+/)) {
    if (line.length + word.length + 1 > 100 && line !== '//') {
      lines.push(line);
      line = '//';
    }
    line += ` ${word}`;
  }
  lines.push(line);
  return lines;
}

/** Throws on a manifest exception that does not name a selected member, or that has no reason. */
function validateExceptions(manifest, selected) {
  const exceptions = new Map();
  for (const entry of manifest.conformanceExceptions ?? []) {
    if (typeof entry?.uid !== 'string' || !selected.has(entry.uid))
      throw new Error(`conformanceExceptions: "${entry?.uid}" is not a selected member`);
    if (typeof entry.reason !== 'string' || !entry.reason.trim())
      throw new Error(`conformanceExceptions: "${entry.uid}" has no reason`);
    if (exceptions.has(entry.uid))
      throw new Error(`conformanceExceptions: "${entry.uid}" is listed twice`);
    exceptions.set(entry.uid, entry.reason.trim());
  }
  return exceptions;
}

/**
 * @param {object} params
 * @param {object} params.inventory Parsed `reference/presentation.full-inventory.json`.
 * @param {object} params.manifest Parsed `manifest.json`.
 * @param {object} [params.reference] Parsed `reference/presentation.reference.json`, for
 *   parameter names only.
 * @param {string} [params.packageImport] The module the generated file imports the package
 *   from.
 * @param {string} [params.typeAssertImport] The module the generated file imports
 *   `IsExact` and `Expect` from.
 */
export function generateConformance({
  inventory,
  manifest,
  reference = { symbols: {} },
  packageImport = '../../src/index.ts',
  typeAssertImport = '../type-assert.ts',
}) {
  const rows = new Map((inventory.endpoints ?? []).map((row) => [row.uid, row]));
  const uids = selectedUids(manifest);
  const exceptions = validateExceptions(manifest, new Set(uids));
  const body = [];
  const skipped = [];
  let assertions = 0;
  let writes = 0;
  let previousOwner = null;
  for (const uid of uids) {
    const row = rows.get(uid);
    if (!row) throw new Error(`The reference inventory has no row for ${uid}`);
    const { owner, member } = localName(uid);
    if (member && owner !== previousOwner) body.push(`// ${uid.split('#')[0]}`, '');
    previousOwner = owner;
    body.push(`// ${uid}`);
    const reason = exceptions.get(uid);
    if (reason) {
      skipped.push({ uid, reason });
      body.push(
        `// Not asserted: manifest.json conformanceExceptions "${uid}".`,
        ...wrapComment(reason)
      );
      body.push('');
      continue;
    }
    const subsetMember = member
      ? reference.symbols?.[owner]?.members?.[member]
      : reference.symbols?.[owner];
    let lines;
    if (row.kind === 'property') lines = emitProperty(uid, row);
    else if (row.kind === 'method' || row.kind === 'function')
      lines = emitCallable(uid, row, subsetMember);
    else throw new Error(`Unsupported reference kind "${row.kind}" for ${uid}`);
    assertions += 1;
    if (lines.some((line) => line.startsWith('export function _write_'))) writes += 1;
    body.push(...lines, '');
  }

  const source = [
    '// GENERATED FILE. Do not edit by hand.',
    '//',
    '// packages/editor-api/scripts/generate-conformance.mjs writes this file from',
    '// compat/manifest.json (the selection and the conformance exceptions) and',
    '// compat/reference/presentation.full-inventory.json (the reference types). Run',
    '// `bun run compat:generate` to write it again.',
    '//',
    '// Each `_assert_*` alias compares one selected Office.js member with the real public type',
    '// of the package. IsExact checks both directions, so a narrower or wider type fails, not',
    '// only an unrelated one. A reference type names the package export of the same name, so a',
    '// missing export fails as a real compiler error. Each `_write_*` function checks that the',
    '// reference write type assigns to the real setter.',
    '',
    `import type { Expect, IsExact } from '${typeAssertImport}';`,
    `import type * as ${NS} from '${packageImport}';`,
    '',
    ...body,
  ];
  while (source.at(-1) === '') source.pop();
  return { assertionsSource: `${source.join('\n')}\n`, assertions, writes, skipped };
}

function readJson(file) {
  return JSON.parse(fs.readFileSync(file, 'utf8'));
}

/** The generation from the committed compat data. */
export function generateFromCompat() {
  const compat = path.join(root, 'compat');
  return generateConformance({
    inventory: readJson(path.join(compat, 'reference', 'presentation.full-inventory.json')),
    manifest: readJson(path.join(compat, 'manifest.json')),
    reference: readJson(path.join(compat, 'reference', 'presentation.reference.json')),
  });
}

function main() {
  const args = process.argv.slice(2);
  if (args.length > 1 || (args.length === 1 && args[0] !== '--check'))
    throw new Error('Usage: compat:generate [--check]');
  const result = generateFromCompat();
  const relative = path.relative(process.cwd(), ASSERTIONS_PATH);
  if (args[0] === '--check') {
    const committed = fs.existsSync(ASSERTIONS_PATH)
      ? fs.readFileSync(ASSERTIONS_PATH, 'utf8')
      : null;
    if (committed !== result.assertionsSource) {
      console.error(`${relative} is out of date. Run \`bun run compat:generate\` and commit it.`);
      process.exitCode = 1;
      return;
    }
    console.log(`${relative} is current.`);
  } else {
    fs.mkdirSync(path.dirname(ASSERTIONS_PATH), { recursive: true });
    fs.writeFileSync(ASSERTIONS_PATH, result.assertionsSource);
    console.log(`Wrote ${relative}`);
  }
  console.log(
    `${result.assertions} member assertions (${result.writes} with a write check), ${result.skipped.length} conformance exceptions.`
  );
  for (const { uid } of result.skipped) console.log(`  not asserted: ${uid}`);
}

if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href) {
  try {
    main();
  } catch (error) {
    console.error(error);
    process.exitCode = 1;
  }
}
