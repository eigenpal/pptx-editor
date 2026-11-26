// The spacing and indent fields of `setParagraph` and `readParagraph` at the protocol level:
// `a:lnSpc`, `a:spcBef`, `a:spcAft` (CT_TextSpacing, ECMA-376 Part 1, 21.1.2.2.5, 21.1.2.2.10,
// 21.1.2.2.9) and `marL` and `indent` of `a:pPr` (CT_TextParagraphProperties, 21.1.2.2.7).
// The object model sends one field for each operation, so these checks of a whole write have
// no other test.

import { describe, expect, test } from 'bun:test';
import { buildPptx, para, run, shape } from '../../../../../scripts/fixtures/builder.ts';
import {
  createServerAutomationHost,
  type AutomationHost,
  type AutomationOperation,
  type AutomationParagraphWrite,
} from '../index.ts';

const IN = 914_400;

function open(text = para(run('Hello'))): AutomationHost {
  const bytes = buildPptx({
    slides: [
      {
        layout: 'blank',
        shapes: shape({ id: 2, x: IN, y: IN, cx: 3 * IN, cy: IN, text }),
      },
    ],
  });
  const opened = createServerAutomationHost(bytes);
  if (!opened.ok) throw new Error(JSON.stringify(opened.error));
  return opened.host;
}

const target: AutomationOperation[] = [
  { op: 'getSlide', key: { index: 0 } },
  { op: 'getShape', container: { result: 0 }, key: { id: '2' } },
];

const write = (host: AutomationHost, paragraph: AutomationParagraphWrite) =>
  host.execute({
    operations: [...target, { op: 'setParagraph', shape: { result: 1 }, paragraph }],
  });

const read = (host: AutomationHost) => {
  const response = host.execute({
    operations: [...target, { op: 'readParagraph', shape: { result: 1 } }],
  });
  if (!response.ok) throw new Error(JSON.stringify(response.results));
  return (response.results[2] as unknown as { value: { paragraph: Record<string, unknown> } }).value
    .paragraph;
};

describe('setParagraph spacing fields', () => {
  test('several fields in one write apply together, the rule before the value', () => {
    const host = open();
    expect(
      write(host, {
        lineSpacingRule: 'exactly',
        lineSpacing: 20,
        lineUnitBefore: 0.5,
        spaceAfter: 3,
        leftIndent: 10,
        firstLineIndent: -5,
      }).ok
    ).toBe(true);
    expect(read(host)).toMatchObject({
      lineSpacing: 20,
      lineSpacingRule: 'exactly',
      spaceBefore: null,
      lineUnitBefore: 0.5,
      spaceAfter: 3,
      lineUnitAfter: null,
      leftIndent: 10,
      firstLineIndent: -5,
    });
  });

  test('a percent string reads as its value, and a rule write keeps it', () => {
    // The percent string of ST_TextSpacingPercentOrPercentString (ECMA-376 Part 1, 20.1.10.77),
    // as in the example of 21.1.2.2.5.
    const host = open(
      para(
        run('Hello'),
        '<a:pPr><a:lnSpc><a:spcPct val="150%"/></a:lnSpc><a:spcBef><a:spcPct val="50%"/></a:spcBef></a:pPr>'
      )
    );
    expect(read(host)).toMatchObject({
      lineSpacing: 18,
      lineSpacingRule: 'multiple',
      lineUnitBefore: 0.5,
    });
    expect(write(host, { lineSpacingRule: 'exactly' }).ok).toBe(true);
    expect(read(host)).toMatchObject({ lineSpacing: 18, lineSpacingRule: 'exactly' });
  });

  test('a rule write keeps a line spacing of at least 0.01 points', () => {
    const host = open(para(run('Hello'), '<a:pPr><a:lnSpc><a:spcPct val="0"/></a:lnSpc></a:pPr>'));
    expect(read(host)).toMatchObject({ lineSpacing: 0, lineSpacingRule: 'multiple' });
    expect(write(host, { lineSpacingRule: 'exactly' }).ok).toBe(true);
    expect(read(host)).toMatchObject({ lineSpacing: 0.01, lineSpacingRule: 'exactly' });
  });

  for (const [label, paragraph, detail] of [
    ['points and lines before', { spaceBefore: 6, lineUnitBefore: 1 }, 'lineUnitBefore'],
    ['points and lines after', { spaceAfter: 6, lineUnitAfter: 1 }, 'lineUnitAfter'],
    ['an unknown rule', { lineSpacingRule: 'atLeast' }, 'lineSpacingRule'],
    ['a bad value next to a good one', { leftIndent: 10, lineSpacing: 0 }, 'lineSpacing'],
  ] as const)
    test(`refuses ${label} with invalid-argument and changes nothing`, () => {
      const host = open();
      const before = host.revision();
      const response = write(host, paragraph as AutomationParagraphWrite);
      expect(response.ok).toBe(false);
      expect(response.results[2]).toMatchObject({ error: { code: 'invalid-argument', detail } });
      expect(host.revision()).toBe(before);
      expect(read(host)).toMatchObject({ leftIndent: 0, spaceBefore: 0 });
    });
});
