// Adapted from DOCX Editor `packages/core/src/contracts/editor.ts` (`EditorFontError`) at commit
// 5993bd223. Keep the exported API equal to the source. It lives in
// the `editor` lane because the PPTX `contracts` lane holds types only.

import type { EditorFontErrorCode, FontFaceRequest } from '../contracts/index.ts';

/** Typed adapter error reported both through `onFontError` and accessible alert UI. @public */
export class EditorFontError extends Error {
  override readonly name: string = 'EditorFontError';
  readonly code: EditorFontErrorCode;
  readonly request?: FontFaceRequest;
  readonly diagnostic?: string;

  constructor(
    code: EditorFontErrorCode,
    message: string,
    details: {
      readonly request?: FontFaceRequest;
      readonly diagnostic?: string;
      readonly cause?: unknown;
    } = {}
  ) {
    super(message, { cause: details.cause });
    this.code = code;
    this.request = details.request;
    this.diagnostic = details.diagnostic;
  }
}
