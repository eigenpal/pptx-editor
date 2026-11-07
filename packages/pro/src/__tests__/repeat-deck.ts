/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// A deck whose slide repeats one threaded comment GUID, as a file can. The React and Vue review
// tests open it.

import { createPptxEditor } from '@pptx-editor.dev/core/editor';
import {
  parseXmlPart,
  readPptx,
  serializeXmlPart,
  writePptx,
  type PackagePart,
} from '@pptx-editor.dev/core/store';
import { buildPptx, shape } from '../../../../scripts/fixtures/builder.ts';
import { reviewModule } from '../index.ts';

const IN = 914_400;

/**
 * A one-slide deck with threads that share a GUID. `twins`: "A" and "B" share an author and a
 * date, and "C" is later; B takes A's id when A is deleted. `pair`: only "A" and "B", with B
 * later; when A is deleted, B's id loses its mark. `same`: three copies of "A", one text.
 */
export async function repeatedThreadDeck(
  kind: 'twins' | 'pair' | 'same' = 'twins'
): Promise<Uint8Array> {
  const editor = createPptxEditor({
    fixedMeasurer: true,
    substitutes: false,
    modules: [reviewModule()],
    author: 'Ada Lovelace',
  });
  const deck = buildPptx({
    slides: [{ shapes: shape({ id: 2, name: 'Title', x: IN, y: IN, cx: IN, cy: IN }) }],
  });
  if (!(await editor.load(deck)).ok) throw new Error('the deck does not open');
  editor.exec({ type: 'addComment', text: 'A', date: '2026-11-07T10:00:00.000' });
  const saved = editor.save()!;
  editor.destroy();
  const r = readPptx(saved);
  if (!r.ok) throw new Error('the saved deck does not read');
  const parts = r.pkg.parts as Map<string, PackagePart>;
  for (const [name, part] of [...parts]) {
    if (!/\/comments\/modernComment/.test(name)) continue;
    const xml =
      part.kind === 'xml' ? serializeXmlPart(part.xml) : new TextDecoder().decode(part.bytes);
    const thread = /<p188:cm [^>]*>.*?<\/p188:cm>/s.exec(xml)![0];
    const copy = (text: string, created?: string): string => {
      const t = thread.replace('<a:t>A</a:t>', `<a:t>${text}</a:t>`);
      return created ? t.replace(/ created="[^"]*"/, ` created="${created}"`) : t;
    };
    const more =
      kind === 'twins'
        ? copy('B') + copy('C', '2026-11-07T11:00:00.000')
        : kind === 'pair'
          ? copy('B', '2026-11-07T10:30:00.000')
          : copy('A') + copy('A');
    const parsed = parseXmlPart(name, xml.replace(thread, thread + more));
    if (!parsed.ok) throw new Error('the patched list does not parse');
    parts.set(name, { kind: 'xml', partName: name, xml: parsed.part });
  }
  return writePptx(r.pkg);
}
