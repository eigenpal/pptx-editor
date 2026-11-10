/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Adapted from DOCX Editor document-format.ts and document-compatibility.ts.
// Source commit: 13f08556889ef7c192023f92aa865cc9cc8e3fe4.
import type * as Y from 'yjs';
import { CollaborationSchemaError } from '../errors.ts';
import { META_KEY, SCHEMA_VERSION } from './schema.ts';

/** Internal format fields. Change only through the collaboration compatibility policy. */
export const COLLABORATION_VERSIONS = Object.freeze({
  protocolVersion: 1,
  sharedSchemaVersion: SCHEMA_VERSION,
  repairVersion: 1,
  canonicalModelVersion: 1,
});
const fields = [
  'protocolVersion',
  'sharedSchemaVersion',
  'repairVersion',
  'canonicalModelVersion',
] as const;
const versionOf = (values: number[]): string => `pptx-collaboration:${values.join('.')}`;

/**
 * Opaque collaboration format, independent of the package version.
 * Send it unchanged during connection admission. Compare it before synchronization.
 * Do not parse, order, override, or use it as a credential.
 * @public
 */
export const COLLABORATION_FORMAT_VERSION: string = versionOf(
  fields.map((field) => COLLABORATION_VERSIONS[field])
);

/**
 * Refuse an incompatible, missing, or malformed connection format before synchronization.
 * Authorize document access separately. This check does not validate content or migrate rooms.
 * @throws CollaborationSchemaError with `schema-version-mismatch` unless the format matches exactly.
 * @public
 */
export function assertCollaborationFormatCompatibility(version: unknown): void {
  if (version !== COLLABORATION_FORMAT_VERSION)
    throw new CollaborationSchemaError(
      'schema-version-mismatch',
      'incompatible collaboration format'
    );
}

/** Write version metadata in the same transaction as the initial document. */
export function writeCollaborationVersions(meta: Y.Map<unknown>): void {
  meta.set('schema', SCHEMA_VERSION);
  for (const field of fields) meta.set(field, COLLABORATION_VERSIONS[field]);
}

/** Inspect the complete tuple, including the legacy schema guard. */
export function hasCompatibleVersions(meta: Y.Map<unknown>): boolean {
  return (
    meta.get('schema') === SCHEMA_VERSION &&
    fields.every((field) => meta.get(field) === COLLABORATION_VERSIONS[field])
  );
}

/**
 * Read a synchronized or restored room's format without joining or changing shared state.
 * Legacy rooms retain their original opaque version strings. Inspection does not validate content.
 * @throws CollaborationSchemaError with `not-initialized` for an empty room,
 * or `schema-version-mismatch` for malformed metadata.
 * @public
 */
export function readCollaborationFormatVersion(doc: Y.Doc): string {
  if (!doc.share.has(META_KEY)) throw new CollaborationSchemaError('not-initialized');
  let meta: Y.Map<unknown>;
  try {
    meta = doc.getMap(META_KEY);
  } catch {
    throw new CollaborationSchemaError('schema-version-mismatch', 'room metadata is not a map');
  }
  const schema = meta.get('schema');
  if (schema === undefined)
    throw new CollaborationSchemaError(
      meta.size === 0 ? 'not-initialized' : 'schema-version-mismatch'
    );
  if (typeof schema !== 'number' || !Number.isSafeInteger(schema) || schema < 1)
    throw new CollaborationSchemaError('schema-version-mismatch', 'invalid room schema');
  if (fields.every((field) => !meta.has(field))) return `pptx-collaboration-${schema}`;
  const values = fields.map((field) => {
    const value = meta.get(field);
    if (typeof value !== 'number' || !Number.isSafeInteger(value) || value < 1)
      throw new CollaborationSchemaError(
        'schema-version-mismatch',
        'invalid collaboration version metadata'
      );
    return value;
  });
  if (values[1] !== schema)
    throw new CollaborationSchemaError('schema-version-mismatch', 'conflicting room schema');
  return versionOf(values);
}
