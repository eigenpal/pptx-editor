/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/

// The review items read `editable` and `resolvable` from the thread data, not from one
// `editor.can` per item. These tests hold the two in agreement on decks that mix plain, rich,
// repeated, and legacy comments, in edit and in view mode.

import { describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { createPptxEditor, type PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import {
  parseXmlPart,
  readPptx,
  serializeXmlPart,
  writePptx,
  type PackagePart,
} from '@pptx-editor.dev/core/store';
import { reviewItems } from '../review/comment-view.ts';
import { reviewModule } from '../index.ts';
import { coverageDeck } from './convergence-harness.ts';

const FIXTURES = join(import.meta.dir, '../../../../fixtures');
const DATE = '2026-11-07T10:00:00.000';

async function open(bytes: Uint8Array): Promise<PptxEditorInstance> {
  const e = createPptxEditor({
    fixedMeasurer: true,
    substitutes: false,
    modules: [reviewModule()],
    author: 'Ada',
  });
  expect((await e.load(bytes)).ok).toBe(true);
  return e;
}

/** `bytes` with each part whose name passes `which` rewritten by `edit`. */
function patch(
  bytes: Uint8Array,
  which: (name: string) => boolean,
  edit: (xml: string) => string
): Uint8Array {
  const r = readPptx(bytes);
  if (!r.ok) throw new Error('the deck does not read');
  for (const [name, part] of [...r.pkg.parts]) {
    if (!which(name)) continue;
    const xml =
      part.kind === 'xml' ? serializeXmlPart(part.xml) : new TextDecoder().decode(part.bytes);
    const next = edit(xml);
    const parsed = parseXmlPart(name, next);
    if (!parsed.ok) throw new Error(`the patched ${name} does not parse`);
    const parts = r.pkg.parts as Map<string, PackagePart>;
    parts.set(name, { kind: 'xml', partName: name, xml: parsed.part });
  }
  return writePptx(r.pkg);
}

/** Each item's `editable` and `resolvable` next to what `editor.can` answers. */
function agreement(e: PptxEditorInstance) {
  const items = reviewItems(e, e.snapshot());
  expect(items.length).toBeGreaterThan(0);
  for (const item of items) {
    const edit = e.can({ type: 'editComment', commentId: item.id, text: 'x' }).ok;
    const status = e.can({
      type: item.resolved ? 'reopenComment' : 'resolveComment',
      commentId: item.parentId ?? item.id,
    }).ok;
    expect({ id: item.id, editable: item.editable, resolvable: item.resolvable }).toEqual({
      id: item.id,
      editable: edit,
      resolvable: status,
    });
  }
  return items;
}

/** A threaded deck: plain threads and replies, a resolved one, a rich body, an extension, a repeat. */
async function threadedDeck(): Promise<Uint8Array> {
  const e = await open(coverageDeck());
  for (const text of ['Plain', 'Rich', 'Resolved', 'Repeated'])
    expect(e.exec({ type: 'addComment', text, date: DATE }).ok).toBe(true);
  const [plain, , resolved] = e.snapshot().comments!;
  for (const text of ['Reply one', 'Reply two'])
    expect(e.exec({ type: 'replyToComment', commentId: plain!.id, text, date: DATE }).ok).toBe(
      true
    );
  expect(e.exec({ type: 'resolveComment', commentId: resolved!.id }).ok).toBe(true);
  const saved = e.save()!;
  e.destroy();
  return patch(
    saved,
    (n) => /\/comments\/modernComment/.test(n),
    (xml) => {
      // A field in the body of "Rich": more than plain text.
      let out = xml.replace(
        /(<a:t>Rich<\/a:t><\/a:r>)/,
        '$1<a:fld id="{11111111-2222-4333-8444-555555555555}" type="x"><a:t>@Ada</a:t></a:fld>'
      );
      // An extension list on the second reply.
      out = out.replace(
        /(<a:t>Reply two<\/a:t><\/a:r><\/a:p><\/p188:txBody>)/,
        '$1<p188:extLst><p:ext xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" uri="{X}"/></p188:extLst>'
      );
      // A second thread with the GUID of "Repeated".
      const repeated =
        /<p188:cm [^>]*>(?:(?!<p188:cm ).)*?<a:t>Repeated<\/a:t>.*?<\/p188:cm>/s.exec(out)![0];
      const head = /^<p188:cm [^>]*>/.exec(repeated)![0];
      const guid = / id="([^"]+)"/.exec(head)![1]!;
      const attrOf = (name: string) => new RegExp(` ${name}="[^"]*"`).exec(head)![0];
      // A thread without an id, which the reader skips with its reply, whose GUID repeats.
      const reply =
        `<p188:replyLst><p188:reply${attrOf('authorId')}${attrOf('created')} id="${guid}">` +
        '<p188:txBody><a:bodyPr/><a:p><a:r><a:t>Hidden</a:t></a:r></a:p></p188:txBody>' +
        '</p188:reply></p188:replyLst>';
      const idless = repeated
        .replace(head, head.replace(` id="${guid}"`, ''))
        .replace('<p188:txBody>', `${reply}<p188:txBody>`)
        .replace('Repeated', 'No id');
      // A file id that reads like the second use of the GUID: `<guid>~2`.
      const tilde = repeated.replace(guid, `${guid}~2`).replace('Repeated', 'Tilde');
      const copy = repeated.replace('Repeated', 'Copy');
      return out.replace('</p188:cmLst>', `${idless}${copy}${tilde}</p188:cmLst>`);
    }
  );
}

/** A threaded deck whose one thread `edit` rewrites. */
async function oneThread(edit: (thread: string, guid: string) => string): Promise<Uint8Array> {
  const e = await open(coverageDeck());
  expect(e.exec({ type: 'addComment', text: 'A', date: DATE }).ok).toBe(true);
  const saved = e.save()!;
  e.destroy();
  return patch(
    saved,
    (n) => /\/comments\/modernComment/.test(n),
    (xml) => {
      const thread = /<p188:cm [^>]*>.*?<\/p188:cm>/s.exec(xml)![0];
      const guid = / id="([^"]+)"/.exec(thread)![1]!;
      return xml.replace(thread, edit(thread, guid));
    }
  );
}

const texts = (e: PptxEditorInstance) => e.snapshot().comments!.map((c) => c.text);

describe('comment ids from a file', () => {
  test('a delete of one repeat moves no id onto another comment', async () => {
    // Three threads with one GUID, each created at another time.
    const at = (thread: string, text: string, minute: string) =>
      thread
        .replace('>A<', `>${text}<`)
        .replace(/ created="[^"]*"/, ` created="2026-11-07T10:${minute}:00.000"`);
    const e = await open(
      await oneThread((t) => at(t, 'A', '01') + at(t, 'B', '02') + at(t, 'C', '03'))
    );
    const [a, b, c] = e.snapshot().comments!;
    expect(new Set([a!.id, b!.id, c!.id]).size).toBe(3);
    expect(e.exec({ type: 'deleteComment', commentId: a!.id }).ok).toBe(true);
    expect(e.snapshot().comments!.map((t) => t.id)).toEqual([b!.id, c!.id]);
    expect(e.exec({ type: 'editComment', commentId: b!.id, text: 'B edited' }).ok).toBe(true);
    expect(texts(e)).toEqual(['B edited', 'C']);
    // The last repeat left keeps its id, and the plain GUID never names it.
    expect(e.exec({ type: 'deleteComment', commentId: c!.id }).ok).toBe(true);
    expect(e.exec({ type: 'editComment', commentId: b!.id, text: 'B again' }).ok).toBe(true);
    expect(texts(e)).toEqual(['B again']);
  });

  test('a file id that holds # stays reachable', async () => {
    const e = await open(await oneThread((t, g) => t.replace(g, 'X#Y')));
    const [thread] = e.snapshot().comments!;
    const [item] = reviewItems(e, e.snapshot());
    expect(item!.editable).toBe(
      e.can({ type: 'editComment', commentId: thread!.id, text: 'x' }).ok
    );
    expect(e.exec({ type: 'editComment', commentId: thread!.id, text: 'Edited' }).ok).toBe(true);
    expect(e.exec({ type: 'replyToComment', commentId: thread!.id, text: 'R' }).ok).toBe(true);
    expect(e.exec({ type: 'deleteComment', commentId: thread!.id }).ok).toBe(true);
    expect(e.snapshot().comments).toEqual([]);
  });

  test('only the canonical key names a comment; a reply without an id is not reachable', async () => {
    const e = await open(
      await oneThread((t) => {
        const head = /^<p188:cm [^>]*>/.exec(t)![0];
        const author = / authorId="[^"]*"/.exec(head)![0];
        const reply =
          `<p188:replyLst><p188:reply${author} id=""><p188:txBody><a:bodyPr/><a:p><a:r>` +
          '<a:t>Hidden</a:t></a:r></a:p></p188:txBody></p188:reply></p188:replyLst>';
        return t.replace('<p188:txBody>', `${reply}<p188:txBody>`);
      })
    );
    const [thread] = e.snapshot().comments!;
    expect(thread!.replies).toEqual([]);
    const slide = thread!.id.slice(0, thread!.id.indexOf('#'));
    for (const alias of ['~1', '~01', '~000001', '~00000000~1', '~00000000~01'])
      expect(
        e.exec({ type: 'editComment', commentId: `${thread!.id}${alias}`, text: 'x' }).ok
      ).toBe(false);
    for (const key of ['', '~1', '~00000000~1'])
      expect(e.exec({ type: 'deleteComment', commentId: `${slide}#${key}` }).ok).toBe(false);
    expect(texts(e)).toEqual(['A']);
  });
});

describe('review items', () => {
  test('editable and resolvable agree with editor.can on a threaded deck', async () => {
    const e = await open(await threadedDeck());
    try {
      const items = agreement(e);
      const by = (text: string) => items.find((i) => i.text.startsWith(text))!;
      expect(by('Plain').editable).toBe(true);
      expect(by('Rich').editable).toBe(false);
      expect(by('Reply two').editable).toBe(false);
      expect(by('Resolved').resolvable).toBe(true);
      // A repeated GUID gets its own id, and the commands reach the copy.
      expect(new Set(items.map((i) => i.id)).size).toBe(items.length);
      // Repeats take their mark (author and creation time) and an ordinal among equal marks.
      expect(by('Repeated').id).toMatch(/~[0-9A-F]{8}~1$/);
      expect(by('Copy').id).toBe(by('Repeated').id.replace(/~1$/, '~2'));
      expect(by('Copy').editable).toBe(true);
      const [guid, mark] = by('Repeated').id.split('#')[1]!.split('~');
      expect(by('Tilde').id.split('#')[1]).toBe(`${guid}~2~${mark}~1`);
      expect(items.some((i) => i.text === 'Hidden' || i.text === 'No id')).toBe(false);
      expect(
        e.exec({ type: 'editComment', commentId: by('Copy').id, text: 'Copy edited' }).ok
      ).toBe(true);
      const texts = e.snapshot().comments!.map((t) => t.text);
      expect(texts).toContain('Repeated');
      expect(texts).toContain('Copy edited');
      const tildeEdit = e.exec({ type: 'editComment', commentId: by('Tilde').id, text: 'Tilde 2' });
      expect(tildeEdit.ok).toBe(true);
      expect(e.snapshot().comments!.map((t) => t.text)).toEqual(
        expect.arrayContaining(['Repeated', 'Copy edited', 'Tilde 2'])
      );
      e.setMode('view');
      expect(agreement(e).every((i) => !i.editable && !i.resolvable)).toBe(true);
    } finally {
      e.destroy();
    }
  });

  test('the items ask editor.can nothing, so a long list stays linear', async () => {
    const e = await open(await threadedDeck());
    try {
      let asked = 0;
      const can = e.can.bind(e);
      const spy = Object.assign(Object.create(e) as PptxEditorInstance, {
        can: (...args: Parameters<PptxEditorInstance['can']>) => (asked++, can(...args)),
      });
      expect(reviewItems(spy, e.snapshot()).length).toBeGreaterThan(5);
      expect(asked).toBe(0);
    } finally {
      e.destroy();
    }
  });

  test('editable and resolvable agree with editor.can on a legacy deck', async () => {
    const fixture = new Uint8Array(readFileSync(join(FIXTURES, 'slide-structure.pptx')));
    // An author id with a hyphen, and a comment with no text element.
    const bytes = patch(
      fixture,
      (n) => n === '/ppt/comments/comment1.xml' || n === '/ppt/commentAuthors.xml',
      (xml) =>
        xml
          .replace(/(authorId|<p:cmAuthor id)="0"/g, '$1="a-1"')
          .replace(
            '</p:cmLst>',
            '<p:cm authorId="a-1" dt="2026-01-01T09:10:00.000" idx="3"><p:pos x="20" y="20"/></p:cm></p:cmLst>'
          )
    );
    const e = await open(bytes);
    try {
      const items = agreement(e);
      expect(items.map((i) => [i.text, i.editable])).toEqual([
        ['Check the chart numbers.', true],
        ['Checked.', true],
        ['', false],
      ]);
      expect(items.every((i) => !i.resolvable)).toBe(true);
      const [thread] = e.snapshot().comments!;
      expect(e.exec({ type: 'editComment', commentId: thread!.id, text: 'Edited' }).ok).toBe(true);
      expect(e.snapshot().comments![0]!.text).toBe('Edited');
      e.setMode('view');
      expect(agreement(e).every((i) => !i.editable)).toBe(true);
    } finally {
      e.destroy();
    }
  });
});
