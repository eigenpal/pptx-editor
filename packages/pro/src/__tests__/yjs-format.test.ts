import { describe, expect, test } from 'bun:test';
import * as Y from 'yjs';
import { createPptxEditor } from '@pptx-editor.dev/core/editor';
import { buildPptx, para, run, shape } from '../../../../scripts/fixtures/builder.ts';
import {
  COLLABORATION_FORMAT_VERSION,
  CollaborationSchemaError,
  collaborationModule,
  createYjsCollaborationSession,
  readCollaborationFormatVersion,
} from '../index.ts';
import { META_KEY } from '../collaboration/yjs/schema.ts';

const deck = buildPptx({
  slides: [
    {
      layout: 'title',
      shapes: shape({
        id: 2,
        x: 914_400,
        y: 914_400,
        cx: 3_657_600,
        cy: 914_400,
        text: para(run('Hello')),
      }),
    },
  ],
});

function codeOf(doc: Y.Doc): string | undefined {
  try {
    readCollaborationFormatVersion(doc);
    return undefined;
  } catch (error) {
    expect(error).toBeInstanceOf(CollaborationSchemaError);
    return (error as CollaborationSchemaError).code;
  }
}

describe('readCollaborationFormatVersion', () => {
  test('an empty Y.Doc holds no room, and reading it creates nothing', () => {
    const doc = new Y.Doc();
    expect(codeOf(doc)).toBe('not-initialized');
    expect(doc.share.size).toBe(0);
  });

  test('a session attached to an empty doc seeds it with this release format', async () => {
    const doc = new Y.Doc();
    const session = createYjsCollaborationSession({
      doc,
      documentId: 'deck',
      identity: { actorId: 'a', name: 'A' },
    });
    // A session that is not attached yet has not seeded the room.
    expect(codeOf(doc)).toBe('not-initialized');
    const editor = createPptxEditor({
      fixedMeasurer: true,
      substitutes: false,
      modules: [collaborationModule({ session })],
    });
    expect((await editor.load(deck)).ok).toBe(true);
    expect(readCollaborationFormatVersion(doc)).toBe(COLLABORATION_FORMAT_VERSION);

    // A peer that receives the room reads the same version before it joins.
    const peer = new Y.Doc();
    Y.applyUpdate(peer, Y.encodeStateAsUpdate(doc));
    const before = Y.encodeStateAsUpdate(peer);
    expect(readCollaborationFormatVersion(peer)).toBe(COLLABORATION_FORMAT_VERSION);
    expect(Y.encodeStateAsUpdate(peer)).toEqual(before);
    editor.destroy();
    session.destroy();
  });

  test('malformed version information is a format mismatch', () => {
    const doc = new Y.Doc();
    doc.getMap(META_KEY).set('schema', 'seven');
    expect(codeOf(doc)).toBe('schema-version-mismatch');
    const other = new Y.Doc();
    other.getArray(META_KEY).push([1]);
    expect(codeOf(other)).toBe('schema-version-mismatch');
  });
});
