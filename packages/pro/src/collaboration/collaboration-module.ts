/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// The collaboration module: wraps a ready session, as DOCX Editor's `collaborationModule`
// does. The host creates the room and its transport before mount; core attaches the session
// to each opened document through `CollaborationDocumentPort`.

import type { EditorCollaborationSession, EditorModule } from '@pptx-editor.dev/core/editor';
import { rememberLicenseKey, type ProLicenseOptions } from '../license.ts';

/** Options of `collaborationModule`. @public */
export interface CollaborationModuleOptions extends ProLicenseOptions {
  /** A ready session, for example from `createYjsCollaborationSession`. */
  readonly session: EditorCollaborationSession;
}

/**
 * The collaboration module. Register it with
 * `createPptxEditor({ modules: [collaborationModule({ session })] })`.
 *
 * @public
 */
export function collaborationModule(options: CollaborationModuleOptions): EditorModule {
  rememberLicenseKey(options.licenseKey);
  return { id: 'collaboration', collaboration: { session: options.session } };
}
