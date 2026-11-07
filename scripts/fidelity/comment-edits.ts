// The comment edits of the `review.comments` row: threaded comments added, replied to, edited,
// resolved, and deleted on a deck without comments; legacy comments edited in their own format;
// slides with comments duplicated; and two collaborators who comment at once. Each recipe saves,
// reopens, reads the threads again, and must open with no repair prompt
// (`SAVE_VERIFIED_COMMENT_EDITS`). `REVIEW_COMMENTS_EDIT` is the row's recipe in `FEATURE_EDITS`.
//
// `bun scripts/fidelity/comment-edits.ts --out <dir>` writes each edited copy for the prompt check.

import { createHash } from 'node:crypto';
import { mkdirSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { parseArgs } from 'node:util';
import { readCommentThreads } from '../../packages/core/src/editor/index.ts';
import type { CommentThread, EditorCommand } from '../../packages/core/src/contracts/index.ts';
import {
  canonicalPackageFingerprint,
  readPptx,
  serializeXmlPart,
} from '../../packages/core/src/store/index.ts';
import {
  added,
  all,
  has,
  kept,
  lacks,
  runFeatureEdit,
  slidePart,
  NEW_SLIDES,
  type EditTools,
  type FeatureEdit,
  type FeatureEditResult,
  type PartRule,
  type SavedDeck,
} from './feature-edit-kit.ts';
import { openPeers } from './collaboration-kit.ts';

const at = (minute: number): string => `2026-11-07T10:${String(minute).padStart(2, '0')}:00.000`;

/** The threads of slide `index` of the open deck. */
function threadsOf(t: EditTools, index = 0): readonly CommentThread[] {
  const pkg = t.editor.session!.store.package;
  return readCommentThreads(pkg, t.editor.session!.info.slides[index]!.partName);
}

/** The threads of slide `index` of a saved deck, as text: author, date, text, and replies. */
function summary(d: SavedDeck, index = 0): string[] {
  return readCommentThreads(d.pkg, slidePart(d.pkg, index)).map((t) =>
    [
      `${t.format}${t.resolved ? ' resolved' : ''} ${t.anchor.kind}`,
      `${t.author} (${t.initials}) ${t.date}: ${t.text}`,
      ...t.replies.map((r) => `  ${r.author} (${r.initials}) ${r.date}: ${r.text}`),
    ].join('\n')
  );
}

/** A failure when the threads of slide `index` are not `want`. */
const threads = (d: SavedDeck, want: readonly string[], index = 0): string | undefined => {
  const got = summary(d, index);
  return got.join('\n') === want.join('\n')
    ? undefined
    : `the threads are\n${got.join('\n')}\nnot\n${want.join('\n')}`;
};

/** The parts that adding the first threaded comment adds or changes. */
const THREADED: readonly PartRule[] = [
  added(/^\/ppt\/comments\/modernComment_[0-9A-F]+_[0-9A-F]+\.xml$/),
  added(/^\/ppt\/authors\.xml$/),
  /^\/\[Content_Types\]\.xml$/,
  /^\/ppt\/_rels\/presentation\.xml\.rels$/,
];

const ADA = { author: 'Ada Lovelace' } as const;
const GRACE = { author: 'Grace Hopper', initials: 'GMH' } as const;

/**
 * The row's recipe: on a deck without comments, a threaded comment on a shape and one on the
 * slide, two replies, an edit, a resolve, and the delete of a reply and of a third thread.
 */
export const REVIEW_COMMENTS_EDIT: FeatureEdit = {
  fixture: 'text-runs',
  parts: THREADED,
  edit: (t) => {
    t.select('Shape 2');
    t.run({ type: 'addComment', text: 'Check the title', date: at(1), ...ADA });
    t.editor.setSelection({ kind: 'none' });
    t.run({ type: 'addComment', text: 'Slide note\nSecond line', date: at(2), ...GRACE });
    t.run({ type: 'addComment', text: 'Remove me', date: at(3), ...ADA });
    const [title, , removed] = threadsOf(t);
    t.run({ type: 'replyToComment', commentId: title!.id, text: 'Done', date: at(4), ...GRACE });
    t.run({ type: 'replyToComment', commentId: title!.id, text: 'Gone', date: at(5), ...ADA });
    t.run({ type: 'editComment', commentId: title!.id, text: 'Check the title & <date>' });
    t.run({ type: 'resolveComment', commentId: title!.id });
    t.run({ type: 'deleteComment', commentId: threadsOf(t)[0]!.replies[1]!.id });
    t.run({ type: 'deleteComment', commentId: removed!.id });
  },
  check: (d) =>
    all(
      threads(d, [
        'modern resolved shape\nAda Lovelace (AL) 2026-11-07T10:01:00.000: Check the title & <date>' +
          '\n  Grace Hopper (GMH) 2026-11-07T10:04:00.000: Done',
        'modern slide\nGrace Hopper (GMH) 2026-11-07T10:02:00.000: Slide note\nSecond line',
      ]),
      has(d.partXml('/ppt/authors.xml'), 'name="Ada Lovelace"', 'name="Grace Hopper"'),
      has(d.slideXml(0), '<p188:commentRel ', '<p14:creationId '),
      has(d.shapeXml(0, 'Shape 2'), '<a16:creationId '),
      kept(d, 0, 'Shape 3')
    ),
};

/** The other comment recipes, by `review.comments: <edit>`. */
export const COMMENT_EDITS: Readonly<Record<string, FeatureEdit>> = {
  // Legacy comments stay legacy: a reply, two edits, a new thread on the chart, and a delete.
  'review.comments: legacy': {
    fixture: 'slide-structure',
    parts: [/^\/ppt\/commentAuthors\.xml$/, /^\/ppt\/comments\/comment1\.xml$/],
    edit: (t) => {
      const [thread] = threadsOf(t);
      t.run({ type: 'replyToComment', commentId: thread!.id, text: 'Thanks', date: at(1), ...ADA });
      t.run({ type: 'editComment', commentId: thread!.id, text: 'Check the chart totals.' });
      t.run({ type: 'editComment', commentId: thread!.replies[0]!.id, text: 'Checked twice.' });
      t.select('Chart 2');
      t.run({ type: 'addComment', text: 'Axis label?', date: at(2), ...ADA });
      t.run({ type: 'addComment', text: 'Remove me', date: at(3), ...ADA });
      t.run({ type: 'deleteComment', commentId: threadsOf(t)[2]!.id });
    },
    check: (d) =>
      all(
        threads(d, [
          'legacy point\nFixture Author (FA) 2026-01-01T09:00:00.000: Check the chart totals.' +
            '\n  Fixture Author (FA) 2026-01-01T09:05:00.000: Checked twice.' +
            '\n  Ada Lovelace (AL) 2026-11-07T10:01:00.000: Thanks',
          'legacy point\nAda Lovelace (AL) 2026-11-07T10:02:00.000: Axis label?',
        ]),
        has(
          d.partXml('/ppt/commentAuthors.xml'),
          '<p:cmAuthor clrIdx="1" id="1" initials="AL" lastIdx="3" name="Ada Lovelace"/>'
        ),
        has(d.partXml('/ppt/comments/comment1.xml'), '<p:pos x="4032" y="1152"/>')
      ),
  },
  // A slide with legacy comments: the copy has its own list, with new indexes.
  'review.comments: duplicate slide': {
    fixture: 'slide-structure',
    parts: NEW_SLIDES.concat(added(/^\/ppt\/(comments|charts|embeddings)\//), [
      /^\/ppt\/commentAuthors\.xml$/,
    ]),
    edit: (t) => t.run({ type: 'duplicateSlide', index: 0 }),
    check: (d) =>
      all(
        threads(
          d,
          [
            'legacy point\nFixture Author (FA) 2026-01-01T09:00:00.000: Check the chart numbers.' +
              '\n  Fixture Author (FA) 2026-01-01T09:05:00.000: Checked.',
          ],
          1
        ),
        has(d.partXml('/ppt/commentAuthors.xml'), 'lastIdx="4"')
      ),
  },
  // A slide with threaded comments: the copy has its own list, with new GUIDs, that names it.
  'review.comments: duplicate threads': {
    fixture: 'text-runs',
    parts: [...THREADED, ...NEW_SLIDES, added(/^\/ppt\/comments\//)],
    edit: (t) => {
      t.select('Shape 3');
      t.run({ type: 'addComment', text: 'Keep this', date: at(1), ...ADA });
      const id = threadsOf(t)[0]!.id;
      t.run({ type: 'replyToComment', commentId: id, text: 'Kept', date: at(2), ...GRACE });
      t.run({ type: 'duplicateSlide', index: 0 });
    },
    check: (d) => {
      const want = [
        'modern shape\nAda Lovelace (AL) 2026-11-07T10:01:00.000: Keep this' +
          '\n  Grace Hopper (GMH) 2026-11-07T10:02:00.000: Kept',
      ];
      const ids = [0, 1].map((i) => readCommentThreads(d.pkg, slidePart(d.pkg, i))[0]?.id);
      return all(
        threads(d, want, 0),
        threads(d, want, 1),
        ids[0]?.split('#')[1] !== ids[1]?.split('#')[1] ? undefined : 'the copy has the same GUID'
      );
    },
  },
  // Two collaborators give the slide its first comments at once, and both reply to one thread.
  'review.comments: collaboration': {
    fixture: 'text-runs',
    collaborate: openPeers,
    parts: [
      ...THREADED,
      added(/^\/ppt\/comments\/modernComment\d+\.xml$/),
      added(/^\/ppt\/authors\d*\.xml$/),
    ],
    edit: (t) => {
      const peer = t.peer!;
      t.link!.pause();
      t.run({ type: 'addComment', text: 'From peer one', date: at(1) });
      peer.run({ type: 'addComment', text: 'From peer two', date: at(2) });
      t.link!.flush();
      const first = threadsOf(t).find((x) => x.text === 'From peer one')!.id;
      t.link!.pause();
      const reply = (text: string, minute: number): EditorCommand => ({
        type: 'replyToComment',
        commentId: first,
        text,
        date: at(minute),
      });
      t.run(reply('One replies', 3));
      peer.run(reply('Two replies', 4));
      t.link!.flush();
      const own = readPptx(t.editor.save()!);
      const other = readPptx(peer.editor.save()!);
      if (!own.ok || !other.ok) throw new Error('a replica did not save');
      if (canonicalPackageFingerprint(own.pkg) !== canonicalPackageFingerprint(other.pkg))
        throw new Error('the replicas differ');
    },
    check: (d) => {
      const got = summary(d).sort();
      const replies = got.find((x) => x.includes('From peer one'));
      const authors = [...d.pkg.parts.keys()].filter((p) => /^\/ppt\/authors\d*\.xml$/.test(p));
      const names = authors.flatMap((p) => d.partXml(p).match(/ name="[^"]*"/g) ?? []);
      return all(
        names.sort().join() === ' name="Peer 1", name="Peer 2"'
          ? undefined
          : `the authors are ${names.join()}`,
        got.length === 2 && got[1]!.includes('Peer 2 (P2) 2026-11-07T10:02:00.000: From peer two')
          ? undefined
          : `the threads are ${got.join(' | ')}`,
        replies?.includes('Peer 1 (P1) 2026-11-07T10:03:00.000: One replies') &&
          replies.includes('Peer 2 (P2) 2026-11-07T10:04:00.000: Two replies')
          ? undefined
          : 'a reply is gone',
        // The save joins the copies that the two peers made.
        lacks(d.partXml('/ppt/_rels/presentation.xml.rels').replace(/authors"/, ''), 'authors"')
      );
    },
  },
};

/**
 * Recipes whose part names, relationship ids, and slide creation id come from the id stripe of a
 * collaboration session. The stripe counts the sessions of the process, so these numbers change
 * with the tests that ran before.
 */
const STRIPED = new Set(['review.comments: collaboration']);

/**
 * The fingerprint of a recipe's copy (`SAVE_VERIFIED_COMMENT_EDITS`): `editFingerprint`, and for a
 * recipe in a collaboration session the same hash with the stripe numbers left out.
 */
export function commentFingerprint(id: string, result: FeatureEditResult): string | undefined {
  if (!STRIPED.has(id) || !result.bytes) return result.fingerprint;
  const read = readPptx(result.bytes);
  if (!read.ok) return undefined;
  const stripe = (text: string): string =>
    text
      .replace(/(authors|modernComment|rId)\d+/g, '$1N')
      .replace(/ cId="\d+"/g, ' cId="N"')
      .replace(/(<p14:creationId [^>]*val=")\d+/g, '$1N')
      .replace(/\{[0-9A-Fa-f-]{36}\}/g, '{GUID}');
  const hash = createHash('sha256');
  for (const [name, part] of [...read.pkg.parts.entries()].sort(([a], [b]) => (a < b ? -1 : 1)))
    hash.update(
      `${stripe(name)}\n${
        part.kind === 'xml'
          ? stripe(serializeXmlPart(part.xml))
          : Buffer.from(part.bytes).toString('base64')
      }\n`
    );
  return `sha256:${hash.digest('hex')}`;
}

if (import.meta.main) {
  const { values } = parseArgs({ options: { out: { type: 'string' } } });
  if (values.out) mkdirSync(values.out, { recursive: true });
  let failed = 0;
  const all = { 'review.comments': REVIEW_COMMENTS_EDIT, ...COMMENT_EDITS };
  for (const [id, recipe] of Object.entries(all)) {
    const result = await runFeatureEdit(id, recipe);
    if (!result.ok) failed++;
    const print = result.ok ? `ok ${commentFingerprint(id, result)}` : `FAIL ${result.failure}`;
    console.log(`${id}: ${print}`);
    const file = id.replace(/[^a-z0-9.-]+/gi, '-');
    if (values.out && result.bytes) writeFileSync(join(values.out, `${file}.pptx`), result.bytes);
  }
  process.exitCode = failed ? 1 : 0;
}
