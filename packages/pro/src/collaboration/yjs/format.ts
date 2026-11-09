/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// The room's format version, read without joining, after DOCX Editor's
// `packages/pro/src/collaboration/document-format.ts`. DOCX Editor commit adf0e178.
//
// A host reads it to tell a room that holds a presentation from an empty Y.Doc: a tab that
// joins an existing room waits for the room before it creates its session, so it never seeds
// a second copy.

import type * as Y from 'yjs';
import { CollaborationSchemaError } from '../errors.ts';
import { META_KEY, SCHEMA_VERSION } from './schema.ts';

/** The version string of a room schema number. Opaque to hosts: compare it, never parse it. */
const versionOf = (schema: number): string => `pptx-collaboration-${schema}`;

/**
 * The room format version that this release writes and opens, as an opaque string. A session
 * refuses a room with another version (`schema-version-mismatch`).
 *
 * @public
 */
export const COLLABORATION_FORMAT_VERSION: string = versionOf(SCHEMA_VERSION);

/**
 * Read the format version that a room records. Reads only the room's metadata: it does not
 * join the room, change it, or check its content. Compare the result with
 * {@link COLLABORATION_FORMAT_VERSION}.
 *
 * @throws CollaborationSchemaError with `not-initialized` when the Y.Doc holds no room yet,
 * or `schema-version-mismatch` when the room's version information is malformed.
 * @public
 */
export function readCollaborationFormatVersion(doc: Y.Doc): string {
  if (!doc.share.has(META_KEY)) throw new CollaborationSchemaError('not-initialized');
  let schema: unknown;
  try {
    schema = doc.getMap(META_KEY).get('schema');
  } catch {
    throw new CollaborationSchemaError('schema-version-mismatch', 'room metadata is not a map');
  }
  if (schema === undefined) throw new CollaborationSchemaError('not-initialized');
  if (typeof schema !== 'number' || !Number.isSafeInteger(schema))
    throw new CollaborationSchemaError('schema-version-mismatch', 'room schema is not a number');
  return versionOf(schema);
}
