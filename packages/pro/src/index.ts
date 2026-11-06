/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
/**
 * `@pptx-editor.dev/pro`: review and realtime collaboration, as `EditorModule`s.
 *
 * Registering a module is the whole enablement story, as in DOCX Editor's pro package.
 *
 * @example
 * ```ts
 * import { createPptxEditor } from '@pptx-editor.dev/core';
 * import { collaborationModule, reviewModule } from '@pptx-editor.dev/pro';
 *
 * const editor = createPptxEditor({
 *   document: bytes,
 *   modules: [reviewModule(), collaborationModule({ session })],
 * });
 * ```
 *
 * @packageDocumentation
 */

export { reviewModule, type ReviewModuleOptions } from './review/review-module.ts';
export { legacyComments } from './review/legacy-comments.ts';
export {
  collaborationModule,
  type CollaborationModuleOptions,
} from './collaboration/collaboration-module.ts';
export {
  createYjsCollaborationSession,
  type YjsCollaborationSession,
  type YjsCollaborationSessionOptions,
} from './collaboration/yjs/session.ts';
export { CollaborationSchemaError } from './collaboration/errors.ts';
export { type AwarenessLike } from './collaboration/yjs/presence.ts';
export { type ProLicenseOptions } from './license.ts';
