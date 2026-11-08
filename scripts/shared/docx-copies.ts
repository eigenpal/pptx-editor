// The manifest of the files that this repository takes from DOCX Editor (`docx-copies.json`),
// and the checks of `check-docx-copies.ts`. The functions here read no file and run no git, so
// the tests run without a checkout.
//
// Each entry names a file here, its kind, its DOCX Editor sources, and the DOCX Editor commit:
// - `copied`: the code equals its one source at the commit, after the normalization in
//   `docx-normalize.ts`.
// - `adapted`: the code changes its sources on purpose, but keeps at least `ADAPTED_FLOOR` of
//   the lines of the shorter body (`lineOverlap`).
// - `reference`: the file follows a DOCX Editor design, and its code is its own.
// The sources of each entry must exist at its commit. A file can have more than one entry.

/** The kinds of manifest entries. */
export const KINDS = ['copied', 'adapted', 'reference'] as const;
export type CopyKind = (typeof KINDS)[number];

/** One manifest entry. Paths are from the root of each repository. */
export interface CopyEntry {
  readonly file: string;
  readonly kind: CopyKind;
  readonly source: readonly string[];
  readonly commit: string | null;
}

/**
 * A mention of DOCX Editor: "DOCX Editor", `docx-editor`, or `docx_editor`, in any case, also
 * across a line break inside a comment ("DOCX\n// Editor"). The package scope
 * `@docx-editor.dev/` is not one.
 */
export const MENTION = /\bdocx(?:[-_]|\s+(?:(?:\/\/+|\*|#|--|<!--)\s*)*)editor(?!\.dev)/i;

/**
 * Tracked files that mention DOCX Editor without a manifest entry, as [pattern, reason]. A
 * pattern that ends in `/` matches a folder, `*.ext` matches an extension, and any other
 * pattern matches one path.
 */
export const MENTIONS_ALLOWED: readonly (readonly [string, string])[] = [
  ['*.md', 'Markdown guides, READMEs, and changelogs describe the products; they hold no code'],
  ['docs/site/content/', 'public documentation pages, which name DOCX Editor for shared concepts'],
  ['scripts/shared/check-docx-copies.ts', 'the drift check, which names the checkout it reads'],
  ['scripts/shared/docx-checkout.ts', 'the checkout search of the drift check'],
  ['scripts/shared/docx-copies.ts', 'the manifest checks, which hold the mention pattern'],
  ['scripts/shared/docx-copies.test.ts', 'tests that hold headers as data'],
  ['scripts/shared/docx-normalize.ts', 'the normalization of the drift check'],
  ['scripts/shared/docx-module-tests.ts', 'runs DOCX Editor tests against the shared modules'],
  ['scripts/notices-sources.test.ts', 'tests that hold a DOCX Editor header as data'],
  ['scripts/shared/docx-copies.json', 'the manifest, which lists DOCX Editor paths'],
  [
    'examples/shared/shell.css',
    'its source is in the docx-editor.dev site repository, which a DOCX Editor checkout lacks',
  ],
];

/** The reason when `path` is in `MENTIONS_ALLOWED`. */
export function allowedReason(path: string): string | undefined {
  const match = MENTIONS_ALLOWED.find(([pattern]) =>
    pattern.startsWith('*.')
      ? path.endsWith(pattern.slice(1))
      : pattern.endsWith('/')
        ? path.startsWith(pattern)
        : path === pattern
  );
  return match?.[1];
}

const COMMIT = /^[0-9a-f]{7,40}$/;

/** The manifest entries of `text`, and each problem with its shape. */
export function parseManifest(text: string): { entries: CopyEntry[]; problems: string[] } {
  const problems: string[] = [];
  let data: unknown;
  try {
    data = JSON.parse(text);
  } catch (error) {
    return { entries: [], problems: [`the manifest is not JSON: ${(error as Error).message}`] };
  }
  if (!Array.isArray(data)) return { entries: [], problems: ['the manifest is not an array'] };
  const entries: CopyEntry[] = [];
  const seen = new Set<string>();
  data.forEach((item: unknown, at) => {
    const where = `entry ${at}`;
    if (typeof item !== 'object' || item === null) {
      problems.push(`${where} is not an object`);
      return;
    }
    const { file, kind, source, commit, ...rest } = item as Record<string, unknown>;
    const name = typeof file === 'string' ? file : where;
    const wrong = (message: string) => problems.push(`${name}: ${message}`);
    if (typeof file !== 'string' || file === '') return wrong('`file` must be a path');
    if (Object.keys(rest).length > 0) wrong(`unknown fields ${Object.keys(rest).join(', ')}`);
    if (!KINDS.includes(kind as CopyKind)) return wrong(`\`kind\` must be ${KINDS.join(', ')}`);
    if (!Array.isArray(source) || source.some((path) => typeof path !== 'string' || path === ''))
      return wrong('`source` must be an array of paths');
    if (commit !== null && (typeof commit !== 'string' || !COMMIT.test(commit)))
      return wrong('`commit` must be a commit hash or null');
    if (source.length === 0) wrong('an entry must name a source');
    if (commit === null) wrong('an entry must name the commit of its sources');
    if (kind === 'copied' && source.length > 1) wrong('a copied entry names one source');
    const key = `${file}\0${kind}\0${commit}\0${[...source].sort().join('\0')}`;
    if (seen.has(key)) wrong(`a second ${kind} entry with the same sources at ${commit}`);
    seen.add(key);
    entries.push({
      file,
      kind: kind as CopyKind,
      source: source as string[],
      commit: commit as string | null,
    });
  });
  return { entries, problems };
}

/** Whether `bytes` look like a binary file: a NUL byte in the first 8000 bytes. */
export const isBinary = (bytes: Uint8Array): boolean => bytes.subarray(0, 8000).includes(0);

/** What `sourceProblems` reads from a DOCX Editor checkout. */
export interface DocxLookup {
  /** The full hash of a commit name, or null when the checkout has no such commit. */
  readonly commit: (name: string) => string | null;
  /** For each `<commit>:<path>`, whether the checkout holds that file or folder. */
  readonly exists: (objects: readonly string[]) => readonly boolean[];
}

/**
 * The entries whose commit or sources are not in DOCX Editor, with a message. An entry without
 * a commit is read at `HEAD`.
 */
export function sourceProblems(
  entries: readonly CopyEntry[],
  lookup: DocxLookup
): { entry: CopyEntry; message: string }[] {
  const problems: { entry: CopyEntry; message: string }[] = [];
  const commits = new Map<string, string | null>();
  const checks: { entry: CopyEntry; source: string; object: string }[] = [];
  for (const entry of entries) {
    const name = entry.commit ?? 'HEAD';
    if (!commits.has(name)) commits.set(name, lookup.commit(name));
    const commit = commits.get(name);
    if (commit === null || commit === undefined) {
      problems.push({ entry, message: `no commit ${name} in DOCX Editor` });
      continue;
    }
    for (const source of entry.source)
      checks.push({ entry, source, object: `${commit}:${source}` });
  }
  const found = checks.length === 0 ? [] : lookup.exists(checks.map(({ object }) => object));
  checks.forEach(({ entry, source }, at) => {
    if (found[at] !== true)
      problems.push({
        entry,
        message: `${source} is not in DOCX Editor at ${entry.commit ?? 'HEAD'}`,
      });
  });
  return problems;
}

/** A copy verb: a text with one near a mention claims a copy. */
const COPY_VERB =
  /\b(?:copied|copies|copy of|adapted|adapts|ported|ports|based on|derived from|taken from)\b/i;

/** A copy claim about the file itself, in any case and across line breaks. */
const COPY_CLAIM = /\b(?:copied|a copy)\s+(?:(?:verbatim|unchanged)\s+)?(?:from|of)\s+/i;

/** A commit hash: 7 to 40 lowercase hex digits with a digit and a letter. */
const COMMIT_IN_TEXT = /(?<![\w#-])(?=[0-9a-f]*[0-9])(?=[0-9a-f]*[a-f])[0-9a-f]{7,40}(?![\w-])/;

/** A file name with a source extension. */
const FILE_NAME = String.raw`[\w-]+(?:\.[\w-]+)*\.(?:ts|tsx|mts|cts|js|jsx|mjs|cjs|css|vue|py|sh|json|mdx|ya?ml|toml)\b`;

/** A path with a folder, such as `scripts/lib/packages.mjs`, or a file name. */
const PATH_TOKEN = new RegExp(
  String.raw`(?<![@\w./-])(?:(?:[\w-][\w.-]*/)+(?:${FILE_NAME}|[\w.-]*\w)|${FILE_NAME})`,
  'g'
);

/** A path right after a mention, as in "DOCX Editor's `x.ts`" or "DOCX Editor (`a/b`)". */
const PATH_AFTER = new RegExp(
  String.raw`^(?:'s)?\s*[(:]?\s*\x60?(?:[\w-][\w.-]*/)*(?:${FILE_NAME}|[\w.-]*\w/[\w.-]*\w)`
);

/**
 * Whether a path token names a file: it ends in a source extension, has two folders, or starts
 * at a repository folder. "React/Vue" and "and/or" do not.
 */
const fileLike = (token: string): boolean =>
  !LIBRARY_NAMES.has(token.toLowerCase()) &&
  (new RegExp(String.raw`^${FILE_NAME}$|/${FILE_NAME}$`).test(token) ||
    token.split('/').length > 2 ||
    /^(?:packages|scripts|docs|examples|e2e|src|lib|test|\.github)\//.test(token));

/** Library names that look like file names. */
const LIBRARY_NAMES = new Set(['office.js', 'pdf.js', 'node.js', 'vue.js', 'next.js']);

/** How far, in characters, a copy verb may be from a mention to count as near it. */
const NEAR = 40;

/** The extensions whose `#` lines are comments. A file without an extension counts too. */
const HASH_COMMENTS = new Set(['.sh', '.py', '.yml', '.yaml', '.toml', '']);

/**
 * A paragraph of a file: consecutive comment lines without a blank comment line, with their
 * markers removed and joined by spaces; or one line of code. `first` and `last` are line
 * indexes.
 */
export interface Paragraph {
  readonly text: string;
  readonly first: number;
  readonly last: number;
  readonly comment: boolean;
}

/** The extension of `path`, or '' for a file such as `.npmrc` or `pre-commit`. */
function extension(path: string): string {
  const name = path.slice(path.lastIndexOf('/') + 1);
  const dot = name.lastIndexOf('.');
  return dot <= 0 ? '' : name.slice(dot);
}

/**
 * The paragraphs of `text`. Line comments (`//`, and `#` in hash comment files), block
 * comments that start a line (`/*`, `<!--`), and the `"comment"` and `"$comment"` strings of a
 * JSON file are comment lines. Any other line is code.
 */
export function paragraphs(text: string, path: string): Paragraph[] {
  const hash = HASH_COMMENTS.has(extension(path));
  const json = extension(path) === '.json';
  const out: Paragraph[] = [];
  let close: string | null = null;
  let open: { lines: string[]; first: number; last: number } | null = null;
  const flush = () => {
    if (open)
      out.push({ text: open.lines.join(' '), first: open.first, last: open.last, comment: true });
    open = null;
  };
  text.split('\n').forEach((line, at) => {
    const trimmed = line.trim();
    let content: string | null = null;
    // A block comment that ends before code on the same line makes the line code.
    const block = (rest: string, end: string): string | null => {
      const stop = rest.indexOf(end);
      if (stop < 0) close = end;
      if (stop >= 0 && rest.slice(stop + end.length).trim() !== '') return null;
      return stop < 0 ? rest : rest.slice(0, stop);
    };
    if (close !== null) {
      const end = close;
      close = null;
      content = block(trimmed, end)?.replace(/^\*+(?!\/)/, '') ?? null;
    } else if (trimmed.startsWith('//')) content = trimmed.replace(/^\/\/+/, '');
    else if (trimmed.startsWith('/*')) content = block(trimmed.slice(2).replace(/^\*+/, ''), '*/');
    else if (trimmed.startsWith('<!--')) content = block(trimmed.slice(4), '-->');
    else if (hash && trimmed.startsWith('#')) content = trimmed.replace(/^#+!?/, '');
    else if (json) content = /^"\$?comment"\s*:\s*"(.*)",?$/.exec(trimmed)?.[1] ?? null;
    if (content === null) {
      flush();
      if (trimmed !== '') out.push({ text: trimmed, first: at, last: at, comment: false });
      return;
    }
    content = content.trim();
    if (content === '') return flush();
    open ??= { lines: [], first: at, last: at };
    open.lines.push(content);
    open.last = at;
  });
  flush();
  return out;
}

const escaped = (text: string) => text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

/**
 * `term` as a whole word or path in `text`, in any case: not part of a longer path or word. A
 * path may follow `docx-editor/`, and a folder may end in `/`.
 */
export function namesTerm(text: string, term: string): boolean {
  const start = '(?:(?<![\\w./-])|(?<=(?<![\\w.-])docx-editor/))';
  return new RegExp(`${start}${escaped(term)}(?![\\w-]|\\.\\w|/[\\w.])`, 'i').test(text);
}

/** What a paragraph lacks to be the header of `entry`, in the order the entry names them. */
function lacks(entry: CopyEntry, text: string): string[] {
  const missing: string[] = [];
  if (entry.kind !== 'reference' && !new RegExp(`\\b${entry.kind}\\b`, 'i').test(text))
    missing.push(`"${entry.kind}"`);
  if (!MENTION.test(text)) missing.push('DOCX Editor');
  for (const source of entry.source) if (!namesTerm(text, source)) missing.push(source);
  if (entry.commit !== null && !namesTerm(text, entry.commit))
    missing.push(`commit ${entry.commit}`);
  return missing;
}

/**
 * The comment paragraph that is the header of `entry`: the first that names the kind (for
 * `copied` and `adapted`), DOCX Editor, each source, and the commit, each as a whole word.
 */
export function headerOf(entry: CopyEntry, text: string, path: string): Paragraph | undefined {
  return paragraphs(text, path).find((p) => p.comment && lacks(entry, p.text).length === 0);
}

/**
 * How the text of a manifest file disagrees with its entry. One comment paragraph must name
 * the kind, DOCX Editor, each source, and the commit (`headerOf`). A file whose entries are
 * not `copied` must not claim a copy of DOCX Editor in any paragraph, in any case.
 */
export function headerProblems(entry: CopyEntry, text: string, path = entry.file): string[] {
  const problems: string[] = [];
  const found = paragraphs(text, path);
  if (!found.some((p) => p.comment && lacks(entry, p.text).length === 0)) {
    const near = found
      .filter((p) => p.comment)
      .map((p) => ({ p, missing: lacks(entry, p.text) }))
      .sort((a, b) => a.missing.length - b.missing.length)[0];
    problems.push(
      near === undefined
        ? 'the file has no comment paragraph'
        : `no comment paragraph names all of the entry; the closest, at line ${near.p.first + 1}, ` +
            `lacks ${near.missing.join(', ')}`
    );
  }
  if (entry.kind !== 'copied') {
    const claim = found.find((p) => claimsCopy(p.text));
    if (claim !== undefined)
      problems.push(
        `line ${claim.first + 1} claims a copy of DOCX Editor, but the entry is ${entry.kind}`
      );
  }
  return problems;
}

/** Whether `text` says "copied from" or "a copy of" DOCX Editor. */
function claimsCopy(text: string): boolean {
  const source = new RegExp(COPY_CLAIM.source + MENTION.source, 'i');
  return source.test(text);
}

/**
 * The structural signs in `text` that it takes something from DOCX Editor, which need a
 * manifest entry, in a paragraph that mentions DOCX Editor: a `docx-editor/` path, a commit
 * hash, a path that is not in this repository (`isLocal` is false), a path right after the
 * mention, or a copy verb near it. A plain name mention, such as "as DOCX Editor's toolbar",
 * needs none.
 */
export function copySignals(
  text: string,
  path: string,
  isLocal: (token: string) => boolean = () => false
): string[] {
  const signals: string[] = [];
  for (const p of paragraphs(text, path)) {
    const mention = new RegExp(MENTION.source, 'gi');
    let match: RegExpExecArray | null;
    const at = (what: string) => signals.push(`line ${p.first + 1}: ${what}`);
    if (!MENTION.test(p.text)) continue;
    const foreign = [...p.text.matchAll(PATH_TOKEN)]
      .map(([token]) => token)
      .find((token) => fileLike(token) && !isLocal(token));
    if (/docx-editor\//i.test(p.text)) at('a docx-editor/ path');
    else if (COMMIT_IN_TEXT.test(p.text)) at('a commit hash');
    else if (foreign !== undefined) at(`the path ${foreign}, which is not in this repository`);
    else
      while ((match = mention.exec(p.text)) !== null) {
        const end = match.index + match[0].length;
        const near = p.text.slice(Math.max(0, match.index - NEAR), end + NEAR);
        if (PATH_AFTER.test(p.text.slice(end))) at('a path right after the mention');
        else if (COPY_VERB.test(near)) at('a copy verb near the mention');
        else continue;
        break;
      }
  }
  return signals;
}

/**
 * The files with copy signals (`copySignals`) but no manifest entry, and not in
 * `MENTIONS_ALLOWED`. `signals` maps each tracked file to its signals.
 */
export function unlistedMentions(
  signals: ReadonlyMap<string, readonly string[]>,
  entries: readonly CopyEntry[]
): string[] {
  const listed = new Set(entries.map((entry) => entry.file));
  return [...signals]
    .filter(
      ([path, found]) => found.length > 0 && !listed.has(path) && allowedReason(path) === undefined
    )
    .map(([path]) => path);
}

/**
 * The share of normalized lines that two bodies have in common, from 0 to 1: the lines in
 * both, counted with repeats, over the lines of the shorter body. Lines without a letter or a
 * digit, such as `}`, do not count.
 */
export function lineOverlap(a: string, b: string): number {
  const lines = (body: string) =>
    body
      .split('\n')
      .map((line) => line.trim())
      .filter((line) => /[\p{L}\p{N}]/u.test(line));
  const left = lines(a);
  const right = lines(b);
  if (left.length === 0 || right.length === 0) return 0;
  const counts = new Map<string, number>();
  for (const line of left) counts.set(line, (counts.get(line) ?? 0) + 1);
  let common = 0;
  for (const line of right) {
    const n = counts.get(line) ?? 0;
    if (n > 0) {
      common += 1;
      counts.set(line, n - 1);
    }
  }
  return common / Math.min(left.length, right.length);
}

/**
 * Whether a path token names a file or folder of this repository: a tracked path, the end of
 * one, or a folder in one. `tracked` lists the tracked files.
 */
export function localPaths(tracked: readonly string[]): (token: string) => boolean {
  const known = new Set<string>();
  for (const file of tracked) {
    const parts = file.split('/');
    for (let from = 0; from < parts.length; from += 1)
      for (let to = from + 1; to <= parts.length; to += 1)
        known.add(parts.slice(from, to).join('/'));
  }
  return (token) => known.has(token.replace(/^\.\//, '').replace(/\/$/, ''));
}

/**
 * The least line overlap (`lineOverlap`) of an `adapted` file with its sources. In the manifest,
 * the lowest overlap of an adapted file is 0.190, about twice the floor. A wrong source overlaps
 * less: `chart-number-format.ts` for `chart-axes.ts` gives 0.022, and for `chart-scale.ts` 0.000.
 * A file under the floor follows a design and is a `reference`, or names a wrong source.
 */
export const ADAPTED_FLOOR = 0.1;

/** The extensions whose bodies `normalizedBody` compares. A file without one is a hash file. */
const COMPARED_EXTENSIONS = new Set([
  '.ts',
  '.mts',
  '.tsx',
  '.js',
  '.mjs',
  '.css',
  '.sh',
  '.py',
  '.yml',
  '.yaml',
  '',
]);

/** Whether the check can compare the body of `path` with `normalizedBody`. */
export const comparable = (path: string): boolean => COMPARED_EXTENSIONS.has(extension(path));

/**
 * How the code of a manifest file disagrees with the kind of its entry, from the normalized
 * bodies of the file (`ours`) and of each source (`theirs`). An `adapted` or `reference` entry
 * with one source must not equal it: it is a copy. An `adapted` entry must keep at least
 * `ADAPTED_FLOOR` of the lines of its sources.
 */
export function kindProblems(entry: CopyEntry, ours: string, theirs: readonly string[]): string[] {
  if (entry.kind === 'copied') return [];
  if (theirs.length === 1 && theirs[0] === ours)
    return [`the code equals ${entry.source[0]}: mark the entry as copied`];
  if (entry.kind !== 'adapted') return [];
  const overlap = lineOverlap(theirs.join('\n'), ours);
  return overlap < ADAPTED_FLOOR
    ? [
        `the code shares ${overlap.toFixed(3)} of its lines with its sources, under ` +
          `${ADAPTED_FLOOR}: name the right source, or mark the entry as a reference`,
      ]
    : [];
}
