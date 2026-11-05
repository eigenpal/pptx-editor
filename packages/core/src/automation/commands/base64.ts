// Base64 for binary parts carried in ops (`addPart`) and clipboard payloads. One helper, so the
// chunk size that keeps `String.fromCharCode` under the argument limit is set in one place.

const CHUNK = 0x8000;

/** Encode bytes as base64. */
export function toBase64(bytes: Uint8Array): string {
  let s = '';
  // `apply` over one bounded chunk, never a spread of the whole part.
  for (let i = 0; i < bytes.length; i += CHUNK)
    s += String.fromCharCode.apply(null, bytes.subarray(i, i + CHUNK) as unknown as number[]);
  return btoa(s);
}

/** Decode base64 to bytes, or null when the text is not valid base64. */
export function fromBase64(b64: string): Uint8Array | null {
  try {
    const bin = atob(b64);
    const out = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i += 1) out[i] = bin.charCodeAt(i);
    return out;
  } catch {
    return null;
  }
}
