// Globals every supported runtime provides (browsers, Node.js ≥ 20, Bun), declared for the
// DOM-free lanes whose tsconfig excludes the DOM library. Nothing DOM-specific belongs here.

declare class TextEncoder {
  encode(input?: string): Uint8Array;
}
declare class TextDecoder {
  constructor(label?: string, options?: { fatal?: boolean; ignoreBOM?: boolean });
  decode(input?: ArrayBufferView | ArrayBuffer): string;
}
declare function atob(data: string): string;
declare function btoa(data: string): string;
// The WHATWG URL class, the subset `layout/harfbuzz-wasm-binary.ts` uses.
declare class URL {
  constructor(url: string | URL, base?: string | URL);
  readonly href: string;
  readonly protocol: string;
  readonly host: string;
  readonly pathname: string;
  readonly search: string;
  readonly hash: string;
  readonly username: string;
  readonly password: string;
}
// The console, the subset the neutral lanes use for host-facing warnings.
declare const console: { warn(...data: unknown[]): void };
// Cancellation, the subset the font composition in `layout/font-resolver.ts` uses. Browsers,
// Node.js, and Bun all provide `AbortSignal` and `DOMException`.
declare class AbortSignal {
  readonly aborted: boolean;
  readonly reason: unknown;
  addEventListener(type: 'abort', listener: () => void, options?: { once?: boolean }): void;
  removeEventListener(type: 'abort', listener: () => void): void;
}
declare class DOMException extends Error {
  constructor(message?: string, name?: string);
}
// The platform CSPRNG, the subset the automation lane uses for section GUIDs and host-scoped
// handle tokens. Browsers, Node.js, and Bun all provide `globalThis.crypto.getRandomValues`.
// `var`, so `globalThis.crypto` names it too.
declare var crypto: { getRandomValues<T extends Uint8Array>(array: T): T };
