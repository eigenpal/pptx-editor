// Copyright (c) 2026 Eigenpal, Inc. Licensed under Apache-2.0.
// Isolated package worker. Protocol follows DOCX Editor scripts/collaboration/worker.mjs,
// commit 13f08556889ef7c192023f92aa865cc9cc8e3fe4. Uses only installed public APIs.
import * as Y from 'yjs';
import { createPptxEditor } from '@pptx-editor.dev/core/editor';
import { canonicalPackageFingerprint, xmlPartOf } from '@pptx-editor.dev/core/store';
import * as pro from '@pptx-editor.dev/pro';
let doc, editor, session;
let updates = [];
// The initial baseline predates the public format inspection exports.
const roomFormat = (room) =>
  pro.readCollaborationFormatVersion
    ? pro.readCollaborationFormatVersion(room)
    : `pptx-collaboration-${room.getMap('pptx-meta-v1').get('schema')}`;
const ownFormat = () => pro.COLLABORATION_FORMAT_VERSION ?? (doc && roomFormat(doc));
const bytes = (value) => Buffer.from(value, 'base64');
const encoded = (value) => Buffer.from(value).toString('base64');
function close() {
  editor?.destroy();
  session?.destroy();
  doc?.destroy();
  editor = session = doc = undefined;
}
function snapshot() {
  if (session.status() !== 'ready') throw new Error(JSON.stringify(session.statusSnapshot()));
  const paragraphs = [];
  function walk(node) {
    if (node.type !== 'element') return;
    if (node.local === 'p') {
      let text = '';
      function collect(n) {
        if (n.type === 'text') text += n.value;
        else for (const child of n.children ?? []) collect(child);
      }
      collect(node);
      paragraphs.push(text);
    } else for (const child of node.children) walk(child);
  }
  for (const slide of editor.session.info.slides)
    walk(xmlPartOf(editor.session.store.package, slide.partName).root);
  return { fingerprint: canonicalPackageFingerprint(editor.session.store.package), paragraphs };
}
async function execute({ op, ...args }) {
  if (op === 'info') return { format: ownFormat() };
  if (op === 'open' || op === 'probe') {
    close();
    doc = new Y.Doc();
    if (args.state) Y.applyUpdate(doc, bytes(args.state));
    const before = encoded(Y.encodeStateAsUpdate(doc));
    session = pro.createYjsCollaborationSession({
      doc,
      documentId: 'release-baseline',
      identity: { actorId: 'worker', name: 'Worker' },
    });
    editor = createPptxEditor({
      fixedMeasurer: true,
      substitutes: false,
      autoCorrect: false,
      modules: [pro.collaborationModule({ session })],
    });
    const result = await editor.load(bytes(args.document));
    if (!result.ok) throw new Error('Failed to load baseline presentation');
    if (op === 'probe')
      return {
        status: session.status(),
        code: session.statusSnapshot().reason?.code,
        unchanged: encoded(Y.encodeStateAsUpdate(doc)) === before,
      };
    updates = [];
    doc.on('update', (update, origin) => {
      if (origin !== 'link') updates.push(encoded(update));
    });
    return snapshot();
  }
  if (op === 'inspect') {
    const saved = new Y.Doc();
    try {
      Y.applyUpdate(saved, bytes(args.state));
      const before = encoded(Y.encodeStateAsUpdate(saved));
      const format = roomFormat(saved);
      if (encoded(Y.encodeStateAsUpdate(saved)) !== before)
        throw new Error('Inspection mutated room');
      let rejected = false;
      try {
        if (pro.assertCollaborationFormatCompatibility)
          pro.assertCollaborationFormatCompatibility(format);
        else if (format !== ownFormat()) throw new Error('Incompatible format');
      } catch {
        rejected = true;
      }
      return { format, rejected };
    } finally {
      saved.destroy();
    }
  }
  if (op === 'updates') return updates.splice(0);
  if (op === 'state') return encoded(Y.encodeStateAsUpdate(doc));
  if (op === 'apply') {
    Y.applyUpdate(doc, bytes(args.state), 'link');
    return null;
  }
  if (op === 'snapshot') return snapshot();
  if (op === 'save') return encoded(editor.save());
  if (op === 'history') return { undo: session.canUndo(), redo: session.canRedo() };
  if (op === 'edit') {
    const nodeId = editor.session.layout(0).items[0].element.nodeId;
    editor.setSelection({
      kind: 'text',
      nodeId,
      anchor: { paragraph: 0, offset: 1 },
      focus: { paragraph: 0, offset: 1 },
    });
    if (!editor.input('text', args.text).ok) throw new Error('Text edit was refused');
    return snapshot();
  }
  if (op === 'undo' || op === 'redo') {
    editor.exec({ type: op });
    return snapshot();
  }
  throw new Error(`Unknown worker operation: ${op}`);
}
process.on('message', async (message) => {
  try {
    process.send({ result: await execute(message) });
  } catch (error) {
    process.send({ error: error.stack });
  }
});
process.on('disconnect', () => {
  close();
  process.exit(0);
});
