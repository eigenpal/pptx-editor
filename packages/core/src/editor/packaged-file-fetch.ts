// A `fetch` for packaged font files under Node, after DOCX Editor's
// `packages/core/src/export/packaged-file-fetch.ts` (EigenPal, `createPackagedFileFetch`).
//
// `@docx-editor.dev/fonts` names each face as `new URL('../assets/X.ttf', import.meta.url)`.
// A browser bundler rewrites that to an HTTP URL, and Bun's `fetch` reads a `file:` URL, but
// Node's `fetch` rejects every `file:` URL. Without this reader, a Node host lays out and
// exports text with no font bytes at all.
//
// The reader stays inside one trusted directory, refuses symbolic links that leave it, and
// never allocates more than `maxBytes`. It reaches `fs`, `path`, and `url` through
// `process.getBuiltinModule` (as `layout/node-module-shim.ts` does), so no bundler sees a
// Node specifier and the browser build never touches this path.

interface FileHandle {
  stat(): Promise<{ size: number }>;
  read(
    buffer: Uint8Array,
    offset: number,
    length: number,
    position: number
  ): Promise<{ bytesRead: number }>;
  close(): Promise<void>;
}

interface NodeBuiltins {
  readonly open: (path: string, flags: number | string) => Promise<FileHandle>;
  readonly realpath: (path: string) => Promise<string>;
  readonly noFollow: number | undefined;
  readonly readOnly: number;
  readonly fileURLToPath: (url: string) => string;
  readonly path: {
    resolve(...paths: string[]): string;
    relative(from: string, to: string): string;
    isAbsolute(path: string): boolean;
  };
}

type BuiltinLoader = (id: string) => unknown;

function builtinLoader(): BuiltinLoader | undefined {
  const runtime = (globalThis as { process?: { getBuiltinModule?: BuiltinLoader } }).process;
  return typeof runtime?.getBuiltinModule === 'function'
    ? (id) => runtime.getBuiltinModule!(id)
    : undefined;
}

function nodeBuiltins(load: BuiltinLoader): NodeBuiltins {
  const fsp = load('node:fs/promises') as Pick<NodeBuiltins, 'open' | 'realpath'>;
  const fs = load('node:fs') as { constants: { O_RDONLY: number; O_NOFOLLOW?: number } };
  const url = load('node:url') as Pick<NodeBuiltins, 'fileURLToPath'>;
  return {
    open: fsp.open,
    realpath: fsp.realpath,
    readOnly: fs.constants.O_RDONLY,
    noFollow: fs.constants.O_NOFOLLOW,
    fileURLToPath: url.fileURLToPath,
    path: load('node:path') as NodeBuiltins['path'],
  };
}

/** Whether `candidate` is strictly inside `root`. */
function isInside(root: string, candidate: string, path: NodeBuiltins['path']): boolean {
  const relation = path.relative(path.resolve(root), path.resolve(candidate));
  return relation !== '' && !relation.startsWith('..') && !path.isAbsolute(relation);
}

class PackagedFileError extends Error {
  constructor(readonly status: 404 | 413) {
    super(status === 404 ? 'packaged file not found' : 'packaged file exceeds maxBytes');
  }
}

async function readConfined(
  node: NodeBuiltins,
  root: string,
  fileUrl: string,
  maxBytes: number
): Promise<Uint8Array> {
  let requested: string;
  try {
    // Throws for an encoded separator (`%2F`) or a null byte: never a packaged face.
    requested = node.path.resolve(node.fileURLToPath(fileUrl));
  } catch {
    throw new PackagedFileError(404);
  }
  if (!isInside(root, requested, node.path)) throw new PackagedFileError(404);
  const flags = node.noFollow === undefined ? 'r' : node.readOnly | node.noFollow;
  let handle: FileHandle;
  try {
    handle = await node.open(requested, flags);
  } catch (error) {
    const code = (error as { code?: unknown }).code;
    if (code === 'ENOENT' || code === 'ENOTDIR' || code === 'ELOOP')
      throw new PackagedFileError(404);
    throw error;
  }
  try {
    // A symbolic link in a parent directory can still point outside: compare real paths.
    const actual = await node.realpath(requested);
    if (!isInside(await node.realpath(root), actual, node.path)) throw new PackagedFileError(404);
    const size = (await handle.stat()).size;
    if (!Number.isSafeInteger(size) || size < 0 || size > maxBytes)
      throw new PackagedFileError(413);
    const bytes = new Uint8Array(size);
    let offset = 0;
    while (offset < size) {
      const { bytesRead } = await handle.read(bytes, offset, size - offset, offset);
      if (bytesRead === 0) break;
      offset += bytesRead;
    }
    return offset === size ? bytes : bytes.slice(0, offset);
  } finally {
    await handle.close().catch(() => {});
  }
}

/**
 * A `fetch` that reads `file:` URLs inside `trustedRoot` from disk and sends every other URL
 * to the global `fetch`. Undefined off Node (no `process.getBuiltinModule`) or when
 * `trustedRoot` is not a `file:` directory: the global `fetch` already serves those hosts.
 */
export function createPackagedFileFetch(options: {
  readonly trustedRoot: URL;
  readonly maxBytes: number;
}): typeof fetch | undefined {
  if (options.trustedRoot.protocol !== 'file:') return undefined;
  const load = builtinLoader();
  if (!load) return undefined;
  let node: NodeBuiltins;
  try {
    node = nodeBuiltins(load);
  } catch {
    return undefined;
  }
  const root = node.path.resolve(node.fileURLToPath(new URL('./', options.trustedRoot).href));
  // A filesystem root is not a confinement.
  if (node.path.resolve(root, '..') === root) return undefined;
  const { maxBytes } = options;
  return (async (input: string | URL | Request, init?: RequestInit): Promise<Response> => {
    const href = typeof input === 'string' ? input : 'href' in input ? input.href : input.url;
    const url = new URL(href);
    if (url.protocol !== 'file:') return fetch(input, init);
    if (init?.signal?.aborted) throw init.signal.reason;
    if (url.username !== '' || url.password !== '' || (url.host !== '' && url.host !== 'localhost'))
      return new Response(null, { status: 404 });
    try {
      const bytes = await readConfined(node, root, url.href, maxBytes);
      if (init?.signal?.aborted) throw init.signal.reason;
      return new Response(bytes as unknown as BodyInit, { status: 200 });
    } catch (error) {
      if (error instanceof PackagedFileError) return new Response(null, { status: error.status });
      if (init?.signal?.aborted) throw error;
      // A file system error message holds the absolute path, and these messages reach the
      // font report and PDF diagnostics. Keep only the error code.
      const code = (error as { code?: unknown }).code;
      throw new Error(
        `the packaged font file could not be read (${typeof code === 'string' ? code : 'unknown error'})`
      );
    }
  }) as typeof fetch;
}
