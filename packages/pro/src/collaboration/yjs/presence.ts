/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Presence over Yjs awareness, after DOCX Editor's `document-awareness.ts` and
// `document-remote-selections.ts`.
//
// An awareness record comes from a PEER, so every field is checked, not believed. A selection
// travels as shared keys (shapes) or Yjs relative positions (text), never as replica-local
// node ids; the receiver maps it onto its own ids through the package mirror. A text endpoint
// is a relative position, so it stays on the same character while peers type around it.

import * as Y from 'yjs';
import type {
  CollaborationIdentity,
  CollaborationLocalSelection,
  CollaborationParticipant,
  CollaborationRemoteSelection,
} from '@pptx-editor.dev/core/editor';
import type { PackageMirror } from './mirror.ts';

/** The awareness field this editor owns. Other fields belong to other clients. */
export const AWARENESS_FIELD = 'pptxEditor';
export const MAX_IDENTITY_LENGTH = 256;
export const MAX_AWARENESS_STATES = 256;
export const MAX_SELECTED_SHAPES = 512;
const MAX_KEY_LENGTH = 256;

const COLOR = /^#(?:[0-9a-f]{3}|[0-9a-f]{6})$/i;

/** The minimal awareness surface the session uses (`y-protocols/awareness`'s `Awareness`). */
export interface AwarenessLike {
  /** This replica's awareness client id. */
  readonly clientID: number;
  /** Every connected client's state, by client id. */
  getStates(): Map<number, Record<string, unknown>>;
  /** Set one field of this replica's state; peers see it on their next update. */
  setLocalStateField(field: string, value: unknown): void;
  /** Listen to changes of any client's state. */
  on(event: 'change', handler: () => void): void;
  /** Stop listening to state changes. */
  off(event: 'change', handler: () => void): void;
}

/** A text point: a relative position in a text node, or the start of an element with no text. */
type EncodedPoint = { readonly rel: unknown } | { readonly key: string };

/** A selection as it travels in awareness. */
export type EncodedSelection =
  | { readonly kind: 'shapes'; readonly part: string; readonly keys: readonly string[] }
  | { readonly kind: 'text'; readonly anchor: EncodedPoint; readonly head: EncodedPoint };

const isText = (v: unknown, max = MAX_IDENTITY_LENGTH): v is string =>
  typeof v === 'string' && v.trim().length > 0 && v.length <= max;

/** Why a host-supplied identity is invalid, or null. */
export function identityError(
  identity: CollaborationIdentity
): 'invalid-identity' | 'invalid-identity-color' | null {
  if (!isText(identity.actorId) || !isText(identity.name)) return 'invalid-identity';
  if (identity.role !== undefined && identity.role !== 'human' && identity.role !== 'agent')
    return 'invalid-identity';
  if (identity.color !== undefined && !COLOR.test(identity.color)) return 'invalid-identity-color';
  return null;
}

/** The identity fields of a peer record, checked. */
function readIdentity(record: Record<string, unknown>): CollaborationIdentity | null {
  const { actorId, name, color, role } = record;
  if (!isText(actorId) || !isText(name)) return null;
  return {
    actorId,
    name,
    ...(typeof color === 'string' && COLOR.test(color) ? { color } : {}),
    ...(role === 'human' || role === 'agent' ? { role } : {}),
  };
}

function recordOf(state: Record<string, unknown> | undefined): Record<string, unknown> | null {
  const value = state?.[AWARENESS_FIELD];
  return typeof value === 'object' && value !== null && !Array.isArray(value)
    ? (value as Record<string, unknown>)
    : null;
}

/** Encode a local selection for the room, or undefined when it names nothing shared yet. */
export function encodeSelection(
  selection: CollaborationLocalSelection,
  mirror: PackageMirror
): EncodedSelection | undefined {
  if (selection.kind === 'shapes') {
    const keys = selection.shapes
      .slice(0, MAX_SELECTED_SHAPES)
      .map((id) => mirror.keyOf(selection.part, id))
      .filter((k): k is string => k !== undefined);
    return keys.length > 0 ? { kind: 'shapes', part: selection.part, keys } : undefined;
  }
  const point = (p: { node: string; offset: number }): EncodedPoint | undefined => {
    const key = mirror.keyOf(selection.part, p.node);
    if (key === undefined) return undefined;
    const text = mirror.textOf(key);
    if (!text) return { key };
    const offset = Math.max(0, Math.min(Math.trunc(p.offset) || 0, text.length));
    return { rel: Y.relativePositionToJSON(Y.createRelativePositionFromTypeIndex(text, offset)) };
  };
  const anchor = point(selection.anchor);
  const head = point(selection.head);
  return anchor && head ? { kind: 'text', anchor, head } : undefined;
}

/** The record this replica publishes: its identity plus an optional selection. */
export function localRecord(
  identity: CollaborationIdentity,
  selection: EncodedSelection | undefined
): Record<string, unknown> {
  return {
    actorId: identity.actorId,
    name: identity.name,
    ...(identity.color ? { color: identity.color } : {}),
    ...(identity.role ? { role: identity.role } : {}),
    ...(selection ? { selection } : {}),
  };
}

/** Everyone the awareness lists, checked and capped. */
export function readParticipants(awareness: AwarenessLike): CollaborationParticipant[] {
  const out: CollaborationParticipant[] = [];
  for (const [clientId, state] of awareness.getStates()) {
    if (out.length >= MAX_AWARENESS_STATES) break;
    const record = recordOf(state);
    const identity = record ? readIdentity(record) : null;
    if (identity) out.push({ ...identity, isLocal: clientId === awareness.clientID });
  }
  return out;
}

function decodePoint(
  value: unknown,
  doc: Y.Doc,
  mirror: PackageMirror
): { part: string; node: string; offset: number } | null {
  if (typeof value !== 'object' || value === null) return null;
  const point = value as Record<string, unknown>;
  if (isText(point.key, MAX_KEY_LENGTH)) {
    const part = mirror.partOf(point.key);
    const node = part ? mirror.idOf(part, point.key) : undefined;
    return part && node ? { part, node, offset: 0 } : null;
  }
  const json = point.rel;
  if (typeof json !== 'object' || json === null) return null;
  let absolute: Y.AbsolutePosition | null;
  try {
    absolute = Y.createAbsolutePositionFromRelativePosition(
      Y.createRelativePositionFromJSON(json),
      doc
    );
  } catch {
    return null;
  }
  if (!absolute || !(absolute.type instanceof Y.Text)) return null;
  const record = mirror.recordOfText(absolute.type);
  const node = record ? mirror.idOf(record.part, record.key) : undefined;
  return record && node ? { part: record.part, node, offset: absolute.index } : null;
}

/** A selection from the room, on this replica's ids, or null when it does not resolve. */
export function decodeSelection(
  value: unknown,
  doc: Y.Doc,
  mirror: PackageMirror
): CollaborationLocalSelection | null {
  if (typeof value !== 'object' || value === null) return null;
  const v = value as Record<string, unknown>;
  if (v.kind === 'shapes' && isText(v.part, 1024) && Array.isArray(v.keys)) {
    const part = v.part;
    const shapes = v.keys
      .slice(0, MAX_SELECTED_SHAPES)
      .filter((k): k is string => isText(k, MAX_KEY_LENGTH))
      .map((k) => mirror.idOf(part, k))
      .filter((id): id is string => id !== undefined);
    return shapes.length > 0 ? { kind: 'shapes', part, shapes } : null;
  }
  if (v.kind === 'text') {
    const anchor = decodePoint(v.anchor, doc, mirror);
    const head = decodePoint(v.head, doc, mirror);
    if (!anchor || !head || anchor.part !== head.part) return null;
    return {
      kind: 'text',
      part: anchor.part,
      anchor: { node: anchor.node, offset: anchor.offset },
      head: { node: head.node, offset: head.offset },
    };
  }
  return null;
}

/** Peers' selections, mapped onto this replica's ids. Unresolvable selections are left out. */
export function readRemoteSelections(
  awareness: AwarenessLike,
  doc: Y.Doc,
  mirror: PackageMirror
): CollaborationRemoteSelection[] {
  const out: CollaborationRemoteSelection[] = [];
  let seen = 0;
  for (const [clientId, state] of awareness.getStates()) {
    if (++seen > MAX_AWARENESS_STATES) break;
    if (clientId === awareness.clientID) continue;
    const record = recordOf(state);
    const identity = record ? readIdentity(record) : null;
    if (!record || !identity) continue;
    const selection = decodeSelection(record.selection, doc, mirror);
    if (!selection) continue;
    out.push({
      actorId: identity.actorId,
      name: identity.name,
      ...(identity.color ? { color: identity.color } : {}),
      selection,
    });
  }
  return out;
}
