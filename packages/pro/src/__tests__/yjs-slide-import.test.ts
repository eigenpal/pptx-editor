// Slide import with comments through two editors on two linked Y.Docs (`AGENTS.md`
// "Collaboration verification"): each peer imports a slide whose comment names an author the
// deck does not have yet, with the link paused, and both replicas must converge to one deck
// that keeps both authors and both comments.

import { describe, expect, test } from 'bun:test';
import * as Y from 'yjs';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import type { AutomationHost } from '@pptx-editor.dev/core/automation';
import {
  createBrowserAutomationHost,
  createPptxEditor,
  type PptxEditorInstance,
} from '@pptx-editor.dev/core/editor';
import {
  attr,
  canonicalPackageFingerprint,
  contentTypeOf,
  parseXmlPart,
  presentationInfo,
  readPptx,
  relationshipsOf,
  type PptxPackage,
} from '@pptx-editor.dev/core/store';
import { buildPptx, para, run, shape } from '../../../../scripts/fixtures/builder.ts';
import { collaborationModule, createYjsCollaborationSession } from '../index.ts';

const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const CT = 'application/vnd.openxmlformats-officedocument.presentationml';

/** Two docs and a sync link that can pause, to make edits concurrent. */
function link(a: Y.Doc, b: Y.Doc) {
  const queue: [Y.Doc, Uint8Array][] = [];
  let paused = false;
  const send = (to: Y.Doc) => (update: Uint8Array, origin: unknown) => {
    if (origin === 'link') return;
    if (paused) queue.push([to, update]);
    else Y.applyUpdate(to, update, 'link');
  };
  a.on('update', send(b));
  b.on('update', send(a));
  return {
    pause: () => (paused = true),
    flush: () => {
      paused = false;
      for (const [to, update] of queue.splice(0)) Y.applyUpdate(to, update, 'link');
    },
  };
}

const destination = (): Uint8Array =>
  buildPptx({
    slides: [
      {
        layout: 'blank',
        shapes: shape({ id: 2, name: 'Box', x: 0, y: 0, cx: 914_400, cy: 914_400 }),
      },
    ],
  });

/** The destination with a comment author list that names Cat, who has no comments yet. */
function destinationWithAuthors(): Uint8Array {
  const files = unzipSync(destination());
  const edit = (name: string, change: (t: string) => string): void => {
    files[name] = strToU8(change(files[name] ? strFromU8(files[name]) : ''));
  };
  edit(
    'ppt/commentAuthors.xml',
    () =>
      `<p:cmAuthorLst xmlns:p="${P}"><p:cmAuthor id="0" name="Cat" initials="C" lastIdx="0" clrIdx="0"/></p:cmAuthorLst>`
  );
  edit('ppt/_rels/presentation.xml.rels', (t) =>
    t.replace(
      '</Relationships>',
      `<Relationship Id="rIdAuthors" Type="${R}/commentAuthors" Target="commentAuthors.xml"/></Relationships>`
    )
  );
  edit('[Content_Types].xml', (t) =>
    t.replace(
      '</Types>',
      `<Override PartName="/ppt/commentAuthors.xml" ContentType="${CT}.commentAuthors+xml"/></Types>`
    )
  );
  return zipSync(files);
}

/** A one-slide deck whose slide has one comment by `name`. */
function commented(name: string, initials: string, text: string): Uint8Array {
  const files = unzipSync(
    buildPptx({
      slides: [
        {
          layout: 'blank',
          shapes: shape({
            id: 2,
            name,
            x: 0,
            y: 0,
            cx: 914_400,
            cy: 914_400,
            text: para(run(text)),
          }),
        },
      ],
    })
  );
  const edit = (name: string, change: (t: string) => string): void => {
    files[name] = strToU8(change(files[name] ? strFromU8(files[name]) : ''));
  };
  edit(
    'ppt/commentAuthors.xml',
    () =>
      `<p:cmAuthorLst xmlns:p="${P}"><p:cmAuthor id="0" name="${name}" initials="${initials}" lastIdx="1" clrIdx="0"/></p:cmAuthorLst>`
  );
  edit(
    'ppt/comments/comment1.xml',
    () =>
      `<p:cmLst xmlns:p="${P}"><p:cm authorId="0" dt="2026-01-01T00:00:00.000" idx="1"><p:pos x="0" y="0"/><p:text>${text}</p:text></p:cm></p:cmLst>`
  );
  edit('ppt/_rels/presentation.xml.rels', (t) =>
    t.replace(
      '</Relationships>',
      `<Relationship Id="rIdAuthors" Type="${R}/commentAuthors" Target="commentAuthors.xml"/></Relationships>`
    )
  );
  edit('ppt/slides/_rels/slide1.xml.rels', (t) =>
    t.replace(
      '</Relationships>',
      `<Relationship Id="rIdC" Type="${R}/comments" Target="../comments/comment1.xml"/></Relationships>`
    )
  );
  edit('[Content_Types].xml', (t) =>
    t.replace(
      '</Types>',
      `<Override PartName="/ppt/commentAuthors.xml" ContentType="${CT}.commentAuthors+xml"/>` +
        `<Override PartName="/ppt/comments/comment1.xml" ContentType="${CT}.comments+xml"/></Types>`
    )
  );
  return zipSync(files);
}

async function replica(
  doc: Y.Doc,
  actorId: string,
  bytes: Uint8Array
): Promise<PptxEditorInstance> {
  const session = createYjsCollaborationSession({
    doc,
    documentId: 'deck',
    identity: { actorId, name: actorId },
  });
  const editor = createPptxEditor({
    fixedMeasurer: true,
    substitutes: false,
    modules: [collaborationModule({ session })],
  });
  expect((await editor.load(bytes)).ok).toBe(true);
  return editor;
}

function importInto(host: AutomationHost, bytes: Uint8Array): void {
  let s = '';
  for (const b of bytes) s += String.fromCharCode(b);
  const response = host.execute({ operations: [{ op: 'insertSlides', base64: btoa(s) }] });
  if (!response.ok) throw new Error(JSON.stringify(response.results));
}

const fingerprint = (e: PptxEditorInstance): string =>
  canonicalPackageFingerprint(e.session!.store.package);

/** A part's XML tree, whether the package holds it as XML or as bytes. */
function tree(pkg: PptxPackage, name: string | undefined) {
  const part = name ? pkg.parts.get(name) : undefined;
  if (!part) return undefined;
  if (part.kind === 'xml') return part.xml.root;
  const parsed = parseXmlPart(name!, new TextDecoder().decode(part.bytes));
  return parsed.ok ? parsed.part.root : undefined;
}

/** Each author's name, and the author name and text of each comment, in slide order. */
function comments(pkg: PptxPackage) {
  const info = presentationInfo(pkg);
  const authorsPart = relationshipsOf(pkg, info.presentationPart).find((r) =>
    r.type.endsWith('/commentAuthors')
  )?.target;
  const authors = new Map<string, string>();
  const ids: string[] = [];
  for (const a of tree(pkg, authorsPart)?.children ?? [])
    if (a.type === 'element') {
      ids.push(attr(a, 'id') ?? '');
      authors.set(attr(a, 'id') ?? '', attr(a, 'name') ?? '');
    }
  const listed: [string | undefined, string][] = [];
  for (const slide of info.slides)
    for (const r of relationshipsOf(pkg, slide.partName)) {
      if (
        !r.type.endsWith('/comments') ||
        contentTypeOf(pkg, r.target ?? '') !== `${CT}.comments+xml`
      )
        continue;
      for (const cm of tree(pkg, r.target)?.children ?? []) {
        if (cm.type !== 'element') continue;
        const text = cm.children.find((c) => c.type === 'element' && c.local === 'text');
        const value =
          text?.type === 'element'
            ? text.children.map((c) => (c.type === 'text' ? c.value : '')).join('')
            : '';
        listed.push([authors.get(attr(cm, 'authorId') ?? ''), value]);
      }
    }
  return {
    authors: [...authors.values()].sort(),
    uniqueIds: new Set(ids).size === ids.length,
    comments: listed.sort(),
  };
}

/** Two replicas of `bytes`; peer A imports Ann's slide and peer B imports Bob's. */
async function imports(bytes: Uint8Array, concurrent: boolean) {
  const docA = new Y.Doc();
  const docB = new Y.Doc();
  const sync = link(docA, docB);
  const a = await replica(docA, 'a', bytes);
  const b = await replica(docB, 'b', bytes);
  if (concurrent) sync.pause();
  importInto(createBrowserAutomationHost(a), commented('Ann', 'A', 'From Ann'));
  importInto(createBrowserAutomationHost(b), commented('Bob', 'B', 'From Bob'));
  sync.flush();
  expect(fingerprint(a)).toBe(fingerprint(b));
  // Both replicas save and reopen to the same deck.
  const reopened = [a, b].map((e) => {
    const r = readPptx(e.save()!);
    if (!r.ok) throw new Error('reopen');
    return canonicalPackageFingerprint(r.pkg);
  });
  expect(reopened[0]).toBe(reopened[1]!);
  return { pkg: a.session!.store.package, a, b };
}

describe('slide import with comments on two peers', () => {
  for (const [label, deck, extra] of [
    ['a deck with no comment authors', destination, []],
    ['a deck with a comment author list', destinationWithAuthors, ['Cat']],
  ] as const) {
    test(`imports one after the other keep both authors and both comments, on ${label}`, async () => {
      const { pkg } = await imports(deck(), false);
      expect(presentationInfo(pkg).slides).toHaveLength(3);
      expect(comments(pkg)).toEqual({
        authors: ['Ann', 'Bob', ...extra].sort(),
        uniqueIds: true,
        comments: [
          ['Ann', 'From Ann'],
          ['Bob', 'From Bob'],
        ],
      });
    });

    // A limit of the room, not of the import: two peers that add a part of the same name at
    // the same time keep one of the two parts, as two concurrent New Slide commands do. The
    // deck stays whole: every comment names a listed author, and author ids stay unique.
    test(`concurrent imports converge to one whole deck that keeps one import, on ${label}`, async () => {
      const { pkg } = await imports(deck(), true);
      expect(presentationInfo(pkg).slides).toHaveLength(2);
      const read = comments(pkg);
      expect(read.uniqueIds).toBe(true);
      expect(read.comments).toHaveLength(1);
      const [author, text] = read.comments[0]!;
      expect(text).toBe(`From ${author}`);
      expect(read.authors).toContain(author!);
    });
  }

  test('undo of one import removes its slide and comment on both peers; redo brings them back', async () => {
    const { a, b } = await imports(destinationWithAuthors(), false);
    expect(b.exec({ type: 'undo' }).ok).toBe(true);
    expect(fingerprint(a)).toBe(fingerprint(b));
    expect(presentationInfo(a.session!.store.package).slides).toHaveLength(2);
    expect(comments(a.session!.store.package).comments).toEqual([['Ann', 'From Ann']]);
    expect(b.exec({ type: 'redo' }).ok).toBe(true);
    expect(fingerprint(a)).toBe(fingerprint(b));
    expect(comments(a.session!.store.package).comments).toEqual([
      ['Ann', 'From Ann'],
      ['Bob', 'From Bob'],
    ]);
  });
});
