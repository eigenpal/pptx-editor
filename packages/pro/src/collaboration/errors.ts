/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// After DOCX Editor's `packages/pro/src/collaboration/errors.ts`. DOCX Editor commit adf0e178.

import type { CollaborationFailureCode } from '@pptx-editor.dev/core/editor';

/**
 * Typed collaboration schema or trust-boundary failure, thrown for invalid factory input.
 * Use `code` for application logic; `detail` is diagnostic information.
 *
 * @public
 */
export class CollaborationSchemaError extends Error {
  constructor(
    /** Stable failure code. Branch on it. */
    readonly code: CollaborationFailureCode,
    /** Bounded diagnostic text for logs. Not a parsing contract. */
    readonly detail?: string
  ) {
    const diagnostic = detail ? `${code}: ${detail}` : code;
    super(
      code === 'schema-version-mismatch'
        ? `${diagnostic}\nCollaboration upgrade required. Save local changes and use a compatible app version.`
        : diagnostic
    );
    this.name = 'CollaborationSchemaError';
  }
}
