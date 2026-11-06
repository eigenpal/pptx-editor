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

interface FileStats {
  readonly size: number;
  readonly dev: number;
  readonly ino: number;
  isFile(): boolean;
}

interface FileHandle {
  stat(): Promise<FileStats>;
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
  readonly stat: (path: string) => Promise<FileStats>;
  /** `O_RDONLY | O_NOFOLLOW | O_NONBLOCK` where the platform has them, else `'r'`. */
  readonly flags: number | string;
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
  const fsp = load('node:fs/promises') as Pick<NodeBuiltins, 'open' | 'realpath' | 'stat'>;
  const { constants: c } = load('node:fs') as {
    constants: { O_RDONLY: number; O_NOFOLLOW?: number; O_NONBLOCK?: number };
  };
  const url = load('node:url') as Pick<NodeBuiltins, 'fileURLToPath'>;
  return {
    open: fsp.open,
    realpath: fsp.realpath,
    stat: fsp.stat,
    // No following a final symbolic link, and no blocking on a FIFO before the type check.
    // Windows has neither flag; the real path and identity checks below still hold there.
    flags: c.O_NOFOLLOW === undefined ? 'r' : c.O_RDONLY | c.O_NOFOLLOW | (c.O_NONBLOCK ?? 0),
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
  let handle: FileHandle;
  try {
    handle = await node.open(requested, node.flags);
  } catch (error) {
    const code = (error as { code?: unknown }).code;
    // ELOOP: a final symbolic link (`O_NOFOLLOW`). EISDIR: a folder on some platforms.
    if (code === 'ENOENT' || code === 'ENOTDIR' || code === 'ELOOP' || code === 'EISDIR')
      throw new PackagedFileError(404);
    throw error;
  }
  try {
    // Only a regular file: never a folder, a FIFO, or a device.
    const opened = await handle.stat();
    if (!opened.isFile()) throw new PackagedFileError(404);
    // A symbolic link in a parent folder can still point outside: compare real paths. Then
    // check that the open file IS the file at that real path, so a link swapped between the
    // open and the check cannot pass.
    const actual = await node.realpath(requested);
    if (!isInside(await node.realpath(root), actual, node.path)) throw new PackagedFileError(404);
    const there = await node.stat(actual);
    if (there.dev !== opened.dev || there.ino !== opened.ino) throw new PackagedFileError(404);
    const size = opened.size;
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
  const rootUrl = new URL('./', options.trustedRoot);
  let root: string;
  try {
    // Throws for a root this platform cannot name, such as a drive-less path on Windows.
    root = node.path.resolve(node.fileURLToPath(rootUrl.href));
  } catch {
    return undefined;
  }
  // A filesystem root is not a confinement.
  if (node.path.resolve(root, '..') === root) return undefined;
  // A package installed on a Windows network share has a host in its URLs: accept that one.
  const rootHost = rootUrl.host === 'localhost' ? '' : rootUrl.host;
  const { maxBytes } = options;
  return (async (input: string | URL | Request, init?: RequestInit): Promise<Response> => {
    const href = typeof input === 'string' ? input : 'href' in input ? input.href : input.url;
    const url = new URL(href);
    if (url.protocol !== 'file:') return fetch(input, init);
    // The init's signal wins, as in `fetch`; a `Request` carries its own.
    const signal =
      init?.signal ?? (typeof input === 'object' && 'signal' in input ? input.signal : undefined);
    if (signal?.aborted) throw signal.reason;
    const host = url.host === 'localhost' ? '' : url.host;
    if (url.username !== '' || url.password !== '' || host !== rootHost)
      return new Response(null, { status: 404 });
    try {
      const bytes = await readConfined(node, root, url.href, maxBytes);
      if (signal?.aborted) throw signal.reason;
      return new Response(bytes as unknown as BodyInit, { status: 200 });
    } catch (error) {
      if (error instanceof PackagedFileError) return new Response(null, { status: error.status });
      if (signal?.aborted) throw error;
      // A file system error message holds the absolute path, and these messages reach the
      // font report and PDF diagnostics. Keep only the error code.
      const code = (error as { code?: unknown }).code;
      throw new Error(
        `the packaged font file could not be read (${typeof code === 'string' ? code : 'unknown error'})`
      );
    }
  }) as typeof fetch;
}
