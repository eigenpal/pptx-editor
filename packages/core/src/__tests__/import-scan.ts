// The one import scanner that `core-lane-imports.test.ts` and `browser-bundle-graph.test.ts`
// share.
//
// DOCX Editor's originals scan source text with regular expressions. Each regex fix there
// followed a statement that the scan had dropped without a sign: a long specifier list, a
// clause comment, a from-less `export type`. This scanner parses the file with the TypeScript
// compiler instead, so no import form depends on its layout. It sees every form that loads a
// module or names one in a type:
//
//   import x from 'm'           import type { T } from 'm'     import 'm'
//   export { x } from 'm'       export type { T } from 'm'     export * as ns from 'm'
//   import x = require('m')     require('m')                   import('m'), import(`m`)
//   type T = import('m').T
//
// Comments and strings never count, at any length. A dynamic import whose specifier is not a
// literal (`import(name)`, `import(`${base}/m`)`) cannot be followed, so the scanner reports it
// with `specifier: null`, and the guards fail on it instead of skipping it.

import ts from 'typescript';

export interface ScannedImport {
  /** The module specifier, or `null` when it is computed at run time. */
  readonly specifier: string | null;
  /**
   * `true` only when the import is erased at compile time: `import type`, `export type`,
   * `import type x = require('m')`, and `import('m').T` in a type position. A mixed
   * `import { value, type T }` is a VALUE import. Under `verbatimModuleSyntax`, so is
   * `import { type T }`: it compiles to `import {} from 'm'`, which still loads the module.
   */
  readonly typeOnly: boolean;
  /** One-based line of the statement, for failure messages. */
  readonly line: number;
  /** The statement as written, on one line, for failure messages. */
  readonly statement: string;
}

function literalText(node: ts.Node | undefined): string | null {
  if (!node) return null;
  if (ts.isStringLiteral(node) || ts.isNoSubstitutionTemplateLiteral(node)) return node.text;
  return null;
}

function scriptKind(fileName: string): ts.ScriptKind {
  return fileName.endsWith('.tsx') ? ts.ScriptKind.TSX : ts.ScriptKind.TS;
}

/** Every import of one source text. `fileName` only picks TS or TSX parsing. */
export function scanImports(fileName: string, source: string): ScannedImport[] {
  const file = ts.createSourceFile(
    fileName,
    source,
    ts.ScriptTarget.Latest,
    true,
    scriptKind(fileName)
  );
  const found: ScannedImport[] = [];
  const record = (node: ts.Node, specifier: string | null, typeOnly: boolean) => {
    const start = node.getStart(file);
    found.push({
      specifier,
      typeOnly,
      line: file.getLineAndCharacterOfPosition(start).line + 1,
      statement: node.getText(file).replace(/\s+/g, ' ').slice(0, 160),
    });
  };

  const visit = (node: ts.Node): void => {
    if (ts.isImportDeclaration(node)) {
      record(node, literalText(node.moduleSpecifier), node.importClause?.isTypeOnly === true);
    } else if (ts.isExportDeclaration(node) && node.moduleSpecifier) {
      record(node, literalText(node.moduleSpecifier), node.isTypeOnly);
    } else if (
      ts.isImportEqualsDeclaration(node) &&
      ts.isExternalModuleReference(node.moduleReference)
    ) {
      record(node, literalText(node.moduleReference.expression), node.isTypeOnly);
    } else if (ts.isImportTypeNode(node)) {
      const argument = node.argument;
      record(node, ts.isLiteralTypeNode(argument) ? literalText(argument.literal) : null, true);
    } else if (ts.isCallExpression(node)) {
      const isDynamicImport = node.expression.kind === ts.SyntaxKind.ImportKeyword;
      const isRequire = ts.isIdentifier(node.expression) && node.expression.text === 'require';
      if (isDynamicImport || isRequire) record(node, literalText(node.arguments[0]), false);
    }
    ts.forEachChild(node, visit);
  };
  visit(file);
  return found;
}

/**
 * Fail a guard with a message a contributor can act on, or pass when `lines` is empty.
 *
 * A thrown `Error` prints its message whole, where `expect(list).toEqual([])` prints only a
 * diff of the list. Each guard builds lines that name the file, the line, the import, the
 * rule it breaks, and the fix.
 */
export function failGuard(lines: readonly string[]): void {
  if (lines.length > 0) throw new Error(`\n${lines.join('\n')}\n`);
}
