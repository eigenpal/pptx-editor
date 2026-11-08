// The manifest of the files that this repository takes from DOCX Editor (`docx-copies.json`),
// and the checks of `check-docx-copies.ts`. The functions here read no file and run no git, so
// the tests run without a checkout.
//
// Each entry names a file here, its kind, its DOCX Editor sources, and the DOCX Editor commit:
// - `copied`: the code equals its one source at the commit, after the normalization in
//   `docx-normalize.ts`.
// - `adapted`: the code changes its sources on purpose, but keeps at least `ADAPTED_FLOOR` of
//   the lines of the shorter body (`lineOverlap`).
// - `reference`: the file follows a DOCX Editor design, and its code is its own: it shares at
//   most `REFERENCE_CEILING` of its lines with each source.
// The sources of each entry must exist at its commit. A file can have more than one entry. An
// optional `reason` is a note for reviewers. It changes one check only: a Markdown guide with a
// reason can describe the copies in other files (`headerProblems`).
//
// The checks catch honest drift and honest labeling mistakes, not deliberate evasion.

/** The kinds of manifest entries. */
export const KINDS = ['copied', 'adapted', 'reference'] as const;
export type CopyKind = (typeof KINDS)[number];

/** One manifest entry. Paths are from the root of each repository. */
export interface CopyEntry {
  readonly file: string;
  readonly kind: CopyKind;
  readonly source: readonly string[];
  readonly commit: string | null;
  readonly reason?: string;
}

/**
 * A mention of DOCX Editor: "DOCX Editor", `docx-editor`, `docx_editor`, `docx-js-editor`,
 * `DocxEditor`, or `docxeditor`, in any case, also across a line break inside a comment
 * ("DOCX\n// Editor"), with one flat run of space and comment marks so the match stays linear.
 * The package scope `@docx-editor.dev/` and the site are not one.
 */
export const MENTION = /\bdocx(?:[-_]?(?:js[-_])?|\s[\s/*#<!-]*)editor(?!\.dev)/i;

/**
 * Tracked files that may hold copy signals without a manifest entry, as [pattern, reason]. A
 * pattern that ends in `/` matches a folder and everything in it. In any other pattern, `*`
 * matches a part of one path segment and `**` any number of folders, so `*.md` matches the
 * Markdown files of the root folder only.
 */
export const MENTIONS_ALLOWED: readonly (readonly [string, string])[] = [
  ['*.md', 'the root guides (README, AGENTS, CONTRIBUTING) describe the products; no code'],
  ['docs/**/*.md', 'the repository documentation, which names DOCX Editor for shared concepts'],
  ['packages/*/CHANGELOG.md', 'release notes that changesets writes, with this repository commits'],
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
  const match = MENTIONS_ALLOWED.find(([pattern]) => {
    if (pattern.endsWith('/')) return path.startsWith(pattern);
    const parts = pattern
      .split('/')
      .map((part) =>
        part === '**' ? '(?:[^/]+/)*' : `${part.split('*').map(escaped).join('[^/]*')}/`
      );
    return new RegExp(`^${parts.join('').replace(/\/$/, '')}$`).test(path);
  });
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
    const { file, kind, source, commit, reason, ...rest } = item as Record<string, unknown>;
    const name = typeof file === 'string' ? file : where;
    const wrong = (message: string) => problems.push(`${name}: ${message}`);
    if (typeof file !== 'string' || file === '') return wrong('`file` must be a path');
    if (Object.keys(rest).length > 0) wrong(`unknown fields ${Object.keys(rest).join(', ')}`);
    if (!KINDS.includes(kind as CopyKind)) return wrong(`\`kind\` must be ${KINDS.join(', ')}`);
    if (!Array.isArray(source) || source.some((path) => typeof path !== 'string' || path === ''))
      return wrong('`source` must be an array of paths');
    if (commit !== null && (typeof commit !== 'string' || !COMMIT.test(commit)))
      return wrong('`commit` must be a commit hash or null');
    if (reason !== undefined && (typeof reason !== 'string' || reason.trim().length < 10))
      wrong('`reason` must be a sentence');
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
      ...(typeof reason === 'string' ? { reason } : {}),
    });
  });
  return { entries, problems };
}

/**
 * The text of a file: its bytes without NUL bytes, decoded as UTF-8, or as Latin-1 when they
 * are not UTF-8. Git calls a file with a NUL byte binary (`i/-text`), so the check reads the
 * bytes itself, and no file is skipped: a mention in any encoding that keeps ASCII is found.
 */
export function textOf(bytes: Uint8Array): string {
  const kept = bytes.includes(0) ? bytes.filter((byte) => byte !== 0) : bytes;
  try {
    return new TextDecoder('utf-8', { fatal: true }).decode(kept);
  } catch {
    return new TextDecoder('latin1').decode(kept);
  }
}

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

/** A copy verb: a paragraph with one and a mention claims a copy. */
const COPY_VERB =
  /\b(?:copied|copying|adapted|adapts|ported|ports? from|port of|based on|derived from|taken from|forked|lifted|borrowed|cloned|verbatim|line for line|same code|duplicat\w*|carried over|identical to)\b/i;

/** A copy word, which a file without a `copied` entry must not put near a mention. */
const COPY_WORD = /\bcop(?:y|ied|ies|ying)\b/i;

/** A commit hash: 7 to 40 lowercase hex digits with a digit and a letter. */
const COMMIT_IN_TEXT = /(?<![\w#-])(?=[0-9a-f]*[0-9])(?=[0-9a-f]*[a-f])[0-9a-f]{7,40}(?![\w-])/;

/** A file name with a source extension. */
const FILE_NAME = String.raw`[\w-]+(?:\.[\w-]+)*\.(?:ts|tsx|mts|cts|js|jsx|mjs|cjs|css|vue|py|sh|json|mdx|ya?ml|toml)\b`;

/** A path with a folder, such as `scripts/lib/packages.mjs`, or a file name. */
const PATH_TOKEN = new RegExp(
  String.raw`(?<![@*\w./-])(?:(?:[\w-][\w.-]*/)+(?:${FILE_NAME}|[\w.-]*\w)|${FILE_NAME})`,
  'g'
);

/** A path right after a mention, as in "DOCX Editor's `x.ts`" or "DOCX Editor (`a/b`)". */
const PATH_AFTER = new RegExp(
  String.raw`^(?:'s)?\s*[(:]?\s*\x60?(?:[\w-][\w.-]*/)*(?:${FILE_NAME}|[\w.-]*\w/[\w.-]*\w)`
);

/**
 * Whether a path token names a file: it ends in a source extension, or starts at a repository
 * folder. "React/Vue" and "DOM/CSS/fetch" do not.
 */
const fileLike = (token: string): boolean =>
  !LIBRARY_NAMES.has(token.toLowerCase()) &&
  (new RegExp(String.raw`^${FILE_NAME}$|/${FILE_NAME}$`).test(token) ||
    /^(?:packages|scripts|docs|examples|e2e|src|lib|test|api|\.github|\.husky)\//.test(token));

/** Library names that look like file names. */
const LIBRARY_NAMES = new Set(['office.js', 'pdf.js', 'node.js', 'vue.js', 'next.js']);

/** How far, in characters, a copy word may be from a mention to count as near it. */
const NEAR = 40;

/** How many comment paragraphs after a mention paragraph the signals also read. */
const FOLLOWING = 2;

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
 * comments that start a line (`/*`, `<!--`), the `"comment"` and `"$comment"` strings of a
 * JSON file, and the prose lines of a Markdown file are comment lines. Any other line is code.
 */
export function paragraphs(text: string, path: string): Paragraph[] {
  const hash = HASH_COMMENTS.has(extension(path));
  const json = extension(path) === '.json';
  const prose = extension(path) === '.md' || extension(path) === '.mdx';
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
    else if (prose && !trimmed.startsWith('```')) content = trimmed;
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
 * not `copied` must not claim a copy of DOCX Editor in any paragraph, in any case, unless it
 * is a Markdown guide whose entry gives a `reason`: a guide can describe the copies in other
 * files (`options.noted`: an entry of the file gives a reason). No paragraph that mentions
 * DOCX Editor may name a foreign file that no entry of the file lists (with `options.isLocal`),
 * or, in a Markdown guide, that no manifest entry lists (`options.recorded`).
 */
export function headerProblems(
  entry: CopyEntry,
  text: string,
  path = entry.file,
  options: {
    isLocal?: (token: string) => boolean;
    sources?: readonly string[];
    copied?: boolean;
    noted?: boolean;
    recorded?: readonly string[];
  } = {}
): string[] {
  const problems: string[] = [];
  const found = paragraphs(text, path);
  const markdown = /\.mdx?$/.test(path);
  if (options.isLocal !== undefined) {
    // A guide may also name any source that the manifest records for another file.
    const named = [
      ...(options.sources ?? entry.source),
      ...(markdown ? (options.recorded ?? []) : []),
    ];
    for (const p of found.filter((paragraph) => MENTION.test(paragraph.text))) {
      const extra = [...p.text.matchAll(PATH_TOKEN)]
        .map(([token]) => token.replace(/^docx-editor\//, ''))
        .filter((token) => fileLike(token) && !options.isLocal!(token))
        .find((token) => {
          const bare = token.replace(/\/$/, '');
          return !named.some(
            (s) =>
              s === bare ||
              s.endsWith(`/${bare}`) ||
              s.startsWith(`${bare}/`) ||
              bare.startsWith(`${s}/`)
          );
        });
      if (extra !== undefined) {
        problems.push(`line ${p.first + 1} names ${extra}, which no entry of the file lists`);
        break;
      }
    }
  }
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
  const guide = markdown && (options.noted ?? entry.reason !== undefined);
  if (entry.kind !== 'copied' && options.copied !== true && !guide) {
    const claim = found.find((p) => claimsCopy(p.text));
    if (claim !== undefined)
      problems.push(
        `line ${claim.first + 1} puts a copy word near DOCX Editor, but the entry is ${entry.kind}`
      );
  }
  return problems;
}

/** Whether `text` has a copy word ("copy", "copied", "copies", "copying") near a mention. */
export function claimsCopy(text: string): boolean {
  // A code span such as `Copy` names a control, not a claim.
  const prose = text.replace(/`[^`]*`/g, (span) => ' '.repeat(span.length));
  const mention = new RegExp(MENTION.source, 'gi');
  let match: RegExpExecArray | null;
  while ((match = mention.exec(prose)) !== null) {
    const end = match.index + match[0].length;
    const before = prose.slice(Math.max(0, match.index - NEAR), match.index).split(SENTENCE_END);
    const after = prose.slice(end, end + NEAR).split(SENTENCE_END);
    if (COPY_WORD.test(before.at(-1)!) || COPY_WORD.test(after[0]!)) return true;
  }
  return false;
}

/** The end of a sentence: a full stop, a question mark, or an exclamation mark, then a space. */
const SENTENCE_END = /[.!?]\s+/;

/**
 * The structural signs in `text` that it takes something from DOCX Editor, which need a
 * manifest entry. In a paragraph that mentions DOCX Editor: a `docx-editor/` path, a commit
 * hash, a path that is not in this repository (`isLocal` is false), a path right after the
 * mention, or a copy verb anywhere. A commit or a foreign path in the next `FOLLOWING` comment
 * paragraphs counts too. A plain name mention, such as "as DOCX Editor's toolbar", needs none.
 */
export function copySignals(
  text: string,
  path: string,
  isLocal: (token: string) => boolean = () => false
): string[] {
  const signals: string[] = [];
  const found = paragraphs(text, path);
  const foreignIn = (p: Paragraph) =>
    [...p.text.matchAll(PATH_TOKEN)]
      .map(([token]) => token)
      .find((token) => fileLike(token) && !isLocal(token));
  found.forEach((p, index) => {
    const at = (what: string) => signals.push(`line ${p.first + 1}: ${what}`);
    if (!MENTION.test(p.text)) return;
    const mention = new RegExp(MENTION.source, 'gi');
    let match: RegExpExecArray | null;
    let after = false;
    while ((match = mention.exec(p.text)) !== null)
      after ||= PATH_AFTER.test(p.text.slice(match.index + match[0].length));
    const foreign = foreignIn(p);
    const next = found
      .slice(index + 1)
      .filter((q) => q.comment)
      .slice(0, FOLLOWING);
    const later = next.find((q) => COMMIT_IN_TEXT.test(q.text) || foreignIn(q) !== undefined);
    if (/docx-editor\//i.test(p.text)) at('a docx-editor/ path');
    else if (COMMIT_IN_TEXT.test(p.text)) at('a commit hash');
    else if (foreign !== undefined) at(`the path ${foreign}, which is not in this repository`);
    else if (after) at('a path right after the mention');
    else if (COPY_VERB.test(p.text)) at('a copy verb in the paragraph of the mention');
    else if (claimsCopy(p.text)) at('a copy word near the mention');
    else if (later !== undefined) at(`a commit or a foreign path at line ${later.first + 1}`);
  });
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
 * The least line overlap (`lineOverlap`) of an `adapted` file with its sources together, and with
 * each source on its own (`kindProblems`). When the floor was set, the lowest overlap of an adapted
 * file was 0.190. The check prints the current lowest.
 */
export const ADAPTED_FLOOR = 0.1;

/**
 * The most line overlap of a `reference` file with any one source: above it, the file keeps
 * the code of that source and is `adapted`. When the ceiling was set, the highest overlap of a
 * reference was 0.440, so no reference was within 0.05 of it. The check prints the current
 * highest (`--report` prints each overlap).
 */
export const REFERENCE_CEILING = 0.5;

/**
 * How the code of a manifest file disagrees with the files of a folder source, as
 * [path, normalized body] pairs that the check can compare: no file may equal the code, and a
 * `reference` must share at most `REFERENCE_CEILING` with each.
 */
export function folderProblems(
  entry: CopyEntry,
  folder: string,
  ours: string,
  files: readonly (readonly [string, string])[]
): string[] {
  if (entry.kind === 'copied') return [];
  const problems: string[] = [];
  for (const [path, body] of files) {
    if (body === ours) {
      problems.push(`the code equals ${path} in ${folder}: mark the entry as copied`);
      continue;
    }
    const overlap = lineOverlap(body, ours);
    if (entry.kind === 'reference' && overlap > REFERENCE_CEILING)
      problems.push(
        `the code shares ${overlap.toFixed(3)} of its lines with ${path} in ${folder}, over ` +
          `${REFERENCE_CEILING}: mark the entry as adapted`
      );
  }
  return problems;
}

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

/** Whether `normalizedBody` can normalize `path`. */
export const normalizable = (path: string): boolean => COMPARED_EXTENSIONS.has(extension(path));

/** The files that the check compares by their raw lines (`rawBody`): guides and data. */
const RAW_EXTENSIONS = new Set(['.md', '.mdx', '.json']);

/** Whether the check can compare the body of `path`: normalized, or by its raw lines. */
export const comparable = (path: string): boolean =>
  normalizable(path) || RAW_EXTENSIONS.has(extension(path));

/** The body of a file that `normalizedBody` cannot normalize: its trimmed lines. */
export const rawBody = (text: string): string =>
  text
    .split('\n')
    .map((line) => line.trim())
    .join('\n');

/** The language of a comparable file: script, style sheet, a file with hash comments, or data. */
function language(path: string): string {
  const ext = extension(path);
  if (ext === '.css') return 'css';
  if (ext === '.md' || ext === '.mdx') return 'markdown';
  if (ext === '.json') return 'json';
  return ['.ts', '.mts', '.tsx', '.js', '.mjs'].includes(ext) ? 'script' : 'hash';
}

/**
 * Whether the check compares the body of `source` with the body of `file`: both comparable,
 * in the same language. A style sheet never shares a line with a script.
 */
export const comparablePair = (file: string, source: string): boolean =>
  comparable(file) && comparable(source) && language(file) === language(source);

/**
 * How the code of a manifest file disagrees with the kind of its entry, from the normalized
 * body of the file (`ours`) and of each source (`theirs`, in the order of `entry.source`; null
 * for a source that the check cannot compare, such as a folder):
 * - An `adapted` or `reference` file must not equal any of its sources: it is a copy.
 * - An `adapted` file must keep `ADAPTED_FLOOR` of the lines of its sources together.
 * - Each source of an `adapted` entry must reach `ADAPTED_FLOOR` on its own, and the check
 *   must be able to compare it. A reference may share none: it follows a design.
 * - A `reference` must share at most `REFERENCE_CEILING` with each source.
 */
export function kindProblems(
  entry: CopyEntry,
  ours: string,
  theirs: readonly (string | null)[]
): string[] {
  if (entry.kind === 'copied') return [];
  const equal = theirs.findIndex((body) => body === ours);
  if (equal >= 0) return [`the code equals ${entry.source[equal]}: mark the entry as copied`];
  const problems: string[] = [];
  const overlaps = theirs.map((body) => (body === null ? null : lineOverlap(body, ours)));
  const bodies = theirs.filter((body): body is string => body !== null);
  if (entry.kind === 'adapted' && bodies.length > 0) {
    const overlap = lineOverlap(bodies.join('\n'), ours);
    if (overlap < ADAPTED_FLOOR)
      problems.push(
        `the code shares ${overlap.toFixed(3)} of its lines with its sources, under ` +
          `${ADAPTED_FLOOR}: name the right source, or mark the entry as a reference`
      );
  }
  if (entry.kind === 'adapted')
    overlaps.forEach((overlap, at) => {
      if (overlap !== null && overlap < ADAPTED_FLOOR)
        problems.push(
          `the code shares ${overlap.toFixed(3)} of its lines with ${entry.source[at]}, under ` +
            `${ADAPTED_FLOOR}: fix the source, or move it to a reference entry`
        );
      else if (overlap === null && bodies.length > 0)
        problems.push(
          `the check cannot compare ${entry.source[at]} with the code: name a file in the same ` +
            'language, or move it to a reference entry'
        );
    });
  if (entry.kind === 'reference')
    overlaps.forEach((overlap, at) => {
      if (overlap !== null && overlap > REFERENCE_CEILING)
        problems.push(
          `the code shares ${overlap.toFixed(3)} of its lines with ${entry.source[at]}, over ` +
            `${REFERENCE_CEILING}: mark the entry as adapted`
        );
    });
  return problems;
}
