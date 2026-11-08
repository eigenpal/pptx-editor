// The normalization of `check-docx-copies.ts`: a file body in the form in which a copy and its
// DOCX Editor source compare equal when only the documented differences apply. No file system
// and no git here, so the tests run without a DOCX Editor checkout.

import { posix } from 'node:path';
import { format } from 'prettier';
import ts from 'typescript';

const SCRIPT_KINDS: Readonly<Record<string, ts.ScriptKind>> = {
  '.ts': ts.ScriptKind.TS,
  '.mts': ts.ScriptKind.TS,
  '.tsx': ts.ScriptKind.TSX,
  '.js': ts.ScriptKind.JS,
  '.mjs': ts.ScriptKind.JS,
};

/**
 * Comments that change how a tool reads the code: triple-slash references, `@ts-` and lint
 * directives, a leading `@jsx` pragma, and bundler annotations. The normalization keeps them.
 */
const DIRECTIVE =
  /^\/\/\/|^\/\/\s*@ts-|^\/[/*]\*?\s*(?:eslint|oxlint)-|^\/\*\*?\s*@jsx|^\/\*\s*[#@]__\w+__\s*\*\/$|^\/\*\s*(?:webpack\w*:|@vite-ignore)/;

/** The parse errors of a source file, which the printer would otherwise drop with their text. */
function parseErrors(source: ts.SourceFile): string[] {
  const diagnostics =
    (source as unknown as { parseDiagnostics?: readonly ts.Diagnostic[] }).parseDiagnostics ?? [];
  return diagnostics.map((d) => ts.flattenDiagnosticMessageText(d.messageText, ' '));
}

function parsed(text: string, kind: ts.ScriptKind): ts.SourceFile {
  const source = ts.createSourceFile('file', text, ts.ScriptTarget.Latest, true, kind);
  const errors = parseErrors(source);
  if (errors.length > 0) throw new Error(`the code does not parse: ${errors[0]}`);
  return source;
}

/** `text` with each comment that is not a directive replaced by a space. JSX text stays. */
function withoutComments(text: string, kind: ts.ScriptKind): string {
  const source = parsed(text, kind);
  const ranges = new Map<number, number>();
  const jsxText: [number, number][] = [];
  const visit = (node: ts.Node): void => {
    if (ts.isJsxText(node)) {
      jsxText.push([node.pos, node.end]);
      return;
    }
    for (const range of ts.getLeadingCommentRanges(text, node.getFullStart()) ?? [])
      ranges.set(range.pos, range.end);
    for (const range of ts.getTrailingCommentRanges(text, node.getEnd()) ?? [])
      ranges.set(range.pos, range.end);
    for (const child of node.getChildren(source)) visit(child);
  };
  visit(source);
  let out = '';
  let at = 0;
  for (const [pos, end] of [...ranges].sort((a, b) => a[0] - b[0])) {
    if (pos < at || jsxText.some(([from, to]) => pos >= from && pos < to)) continue;
    const comment = text.slice(pos, end);
    out += text.slice(at, pos) + (DIRECTIVE.test(comment) ? comment : ' ');
    at = end;
  }
  return out + text.slice(at);
}

/**
 * The module that a local import specifier names, as a path from the repository root without
 * its extension or a trailing `/index`: `./b.ts` in `packages/core/src/a/x.ts` is
 * `packages/core/src/a/b`, and the workspace package `@scope/core/layout` is
 * `packages/core/src/layout`. Null for any other specifier.
 */
export function localModule(specifier: string, from: string): string | null {
  let target: string;
  if (/^\.{1,2}\//.test(specifier)) {
    target = posix.normalize(posix.join(posix.dirname(from), specifier));
  } else {
    const m = /^@(?:pptx-editor\.dev|docx-editor\.dev|eigenpal)\/([^/]+)(?:\/(.+))?$/.exec(
      specifier
    );
    if (m === null) return null;
    target = `packages/${m[1]}/src${m[2] ? `/${m[2]}` : ''}`;
  }
  return target.replace(/\.(?:d\.)?(?:tsx?|mts|mjs|js)$/, '').replace(/\/index$/, '');
}

/** The AST rewrites of steps 2 to 4 of `normalizedBody`, for the file at `path`. */
function rewrite(path: string): ts.TransformerFactory<ts.SourceFile> {
  const masked = (specifier: string): string => {
    const module = localModule(specifier, path);
    return module === null ? specifier : `<module>/${module}`;
  };
  return (context) => (root) => {
    const f = context.factory;
    const literal = (node: ts.Expression | undefined): node is ts.StringLiteral =>
      node !== undefined && ts.isStringLiteral(node);
    const visit = (node: ts.Node): ts.Node => {
      if (ts.isImportDeclaration(node) && literal(node.moduleSpecifier)) {
        const next = ts.visitEachChild(node, visit, context);
        return f.updateImportDeclaration(
          next,
          next.modifiers,
          next.importClause,
          f.createStringLiteral(masked(node.moduleSpecifier.text)),
          next.attributes
        );
      }
      if (ts.isExportDeclaration(node) && literal(node.moduleSpecifier)) {
        const next = ts.visitEachChild(node, visit, context);
        return f.updateExportDeclaration(
          next,
          next.modifiers,
          next.isTypeOnly,
          next.exportClause,
          f.createStringLiteral(masked(node.moduleSpecifier.text)),
          next.attributes
        );
      }
      if (
        ts.isCallExpression(node) &&
        (node.expression.kind === ts.SyntaxKind.ImportKeyword ||
          (ts.isIdentifier(node.expression) && node.expression.text === 'require')) &&
        literal(node.arguments[0])
      ) {
        const [first, ...rest] = node.arguments as unknown as [
          ts.StringLiteral,
          ...ts.Expression[],
        ];
        return f.updateCallExpression(node, node.expression, node.typeArguments, [
          f.createStringLiteral(masked(first.text)),
          ...rest.map((argument) => ts.visitNode(argument, visit) as ts.Expression),
        ]);
      }
      if (
        ts.isElementAccessExpression(node) &&
        node.questionDotToken === undefined &&
        ts.isStringLiteral(node.argumentExpression) &&
        /^[A-Za-z_$][\w$]*$/.test(node.argumentExpression.text)
      ) {
        return f.createPropertyAccessExpression(
          ts.visitNode(node.expression, visit) as ts.Expression,
          node.argumentExpression.text
        );
      }
      if (ts.isStringLiteral(node) && /@(?:docx|pptx)-editor\.dev\//.test(node.text)) {
        return f.createStringLiteral(
          node.text.replace(/@(?:docx|pptx)-editor\.dev\//g, '@editor.dev/')
        );
      }
      const next = ts.visitEachChild(node, visit, context);
      if (ts.canHaveModifiers(next)) {
        const modifiers = ts.getModifiers(next);
        if (modifiers?.some((m) => m.kind === ts.SyntaxKind.OverrideKeyword)) {
          return f.replaceModifiers(
            next,
            modifiers.filter((m) => m.kind !== ts.SyntaxKind.OverrideKeyword)
          );
        }
      }
      return next;
    };
    return ts.visitNode(root, visit) as ts.SourceFile;
  };
}

/** The quote that is still open at the end of a shell line, given the one open at its start. */
function shellQuoteAfter(line: string, open: string | null): string | null {
  for (let at = 0; at < line.length; at += 1) {
    const ch = line[at]!;
    if (open === "'") {
      if (ch === "'") open = null;
    } else if (open === '"') {
      if (ch === '\\') at += 1;
      else if (ch === '"') open = null;
    } else if (ch === '\\') at += 1;
    else if (ch === '#' && (at === 0 || /\s/.test(line[at - 1]!))) break;
    else if (ch === '"' || ch === "'") open = ch;
  }
  return open;
}

/**
 * A `#` comment file without its full-line comments and blank lines. The lines of a heredoc
 * or a multi-line quoted string (shell), a triple-quoted string (Python), and a block scalar
 * (YAML) are content, so they stay as they are, and so does the text after a `#` on a code line.
 */
function hashCode(text: string, ext: string): string {
  const out: string[] = [];
  let heredoc: { readonly end: string; readonly tabs: boolean } | null = null;
  let quote: string | null = null;
  let triple: string | null = null;
  let scalarIndent: number | null = null;
  for (const raw of text.split('\n')) {
    const line = raw.trimEnd();
    if (heredoc !== null) {
      out.push(line);
      if ((heredoc.tabs ? line.replace(/^\t+/, '') : line) === heredoc.end) heredoc = null;
      continue;
    }
    if (quote !== null) {
      out.push(raw);
      quote = shellQuoteAfter(raw, quote);
      continue;
    }
    if (triple !== null) {
      out.push(line);
      if (line.split(triple).length % 2 === 0) triple = null;
      continue;
    }
    if (scalarIndent !== null) {
      if (line === '' || line.length - line.trimStart().length > scalarIndent) {
        out.push(line);
        continue;
      }
      scalarIndent = null;
    }
    if (line.trim() === '' || /^\s*#/.test(line)) continue;
    out.push(line);
    if (ext === '.sh') {
      const start = /(?<!<)<<(?!<)(-?)\s*(['"]?)([A-Za-z_]\w*)\2/.exec(line);
      if (start) heredoc = { end: start[3]!, tabs: start[1] === '-' };
      else quote = shellQuoteAfter(raw, null);
    } else if (ext === '.py') {
      for (const delimiter of ['"""', "'''"])
        if (line.split(delimiter).length % 2 === 0) {
          triple = delimiter;
          break;
        }
    } else if (/:\s*[|>][-+0-9]*$/.test(line)) {
      scalarIndent = line.length - line.trimStart().length;
    }
  }
  return out.join('\n');
}

/**
 * The body of a copy or of its source at `path` (from the root of its repository), normalized
 * for the comparison. It throws when TypeScript or JavaScript does not parse.
 *
 * 1. Comments are removed, and so the provenance header and license notices go too. Directive
 *    comments (`///`, `// @ts-`, `eslint-`, `oxlint-`, a leading `@jsx` pragma, and bundler
 *    annotations such as `/*#__PURE__*\/`) stay, and so does JSX text. In `#` comment files,
 *    only full-line comments outside heredocs, quoted strings, triple-quoted strings, and YAML
 *    block scalars go, with blank lines and trailing spaces. A leading `#!` line goes too.
 * 2. In TypeScript and JavaScript, each local import specifier, a relative path or a workspace
 *    package, becomes `<module>/` and the module path from the repository root (`localModule`).
 * 3. `@docx-editor.dev/` and `@pptx-editor.dev/` become `@editor.dev/` in string literals.
 * 4. The `override` modifier is removed, and a constant property name in brackets (`a['b']`)
 *    becomes a dot access (`a.b`): this repository's compiler options require both.
 * 5. The TypeScript printer and then Prettier, with fixed options, print the code, so that
 *    layout is never drift. CSS is printed by Prettier.
 */
export async function normalizedBody(text: string, ext: string, path: string): Promise<string> {
  const body = text.replace(/^#!.*\n/, '');
  if (ext === '.css') {
    return format(body.replace(/\/\*[\s\S]*?\*\//g, ''), { parser: 'css', printWidth: 100 });
  }
  const kind = SCRIPT_KINDS[ext];
  if (kind === undefined) return `${hashCode(body, ext).trimEnd()}\n`;
  const source = parsed(withoutComments(body, kind), kind);
  const result = ts.transform(source, [rewrite(path)]);
  let code: string;
  try {
    code = ts.createPrinter().printFile(result.transformed[0]!);
  } finally {
    result.dispose();
  }
  return format(code, {
    parser: ext === '.js' || ext === '.mjs' ? 'babel' : 'typescript',
    printWidth: 100,
    singleQuote: true,
    trailingComma: 'es5',
  });
}

/** The first line where two normalized bodies differ, as a short excerpt. */
export function firstDifference(expected: string, actual: string): string {
  const a = expected.split('\n');
  const b = actual.split('\n');
  let at = 0;
  while (at < a.length && at < b.length && a[at] === b[at]) at += 1;
  const show = (line: string | undefined): string =>
    line === undefined ? '(end of file)' : line.trim().slice(0, 96);
  return `normalized line ${at + 1}:\n      DOCX: ${show(a[at])}\n      here: ${show(b[at])}`;
}
