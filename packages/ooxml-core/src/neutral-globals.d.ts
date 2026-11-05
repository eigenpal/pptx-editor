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
