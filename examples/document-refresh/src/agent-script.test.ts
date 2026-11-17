// The scripted agent of the replay page: each step makes the one change its title names, as
// `comparePresentations()` computes it, and the mock route accepts only step numbers in digits.
// The suite runs this file with no global DOM.

import { describe, expect, test } from 'bun:test';
import { comparePresentations } from '@pptx-editor.dev/pro';
import { AGENT_STEPS, agentPptx, agentStepParam } from '../agent-script.ts';

describe('the scripted agent', () => {
  test('each step makes one change of the expected scope and action', () => {
    const expected = [
      { slideIndex: 1, scope: 'text', action: 'changed', aspects: ['text'] },
      { slideIndex: 1, scope: 'shape', action: 'changed', aspects: ['fill', 'text'] },
      { slideIndex: 0, scope: 'text', action: 'changed', aspects: ['text'] },
      { slideIndex: 3, scope: 'slide', action: 'added' },
      { slideIndex: 2, scope: 'notes', action: 'changed', aspects: ['text'] },
      { slideIndex: 0, scope: 'shape', action: 'removed' },
      { slideIndex: 2, scope: 'slide', action: 'moved', previousSlideIndex: 4 },
    ];
    expect(expected).toHaveLength(AGENT_STEPS.length);
    expected.forEach(({ slideIndex, ...computed }, i) => {
      const r = comparePresentations(agentPptx(i), agentPptx(i + 1));
      if (!r.ok) throw new Error(`step ${i + 1}: ${r.message}`);
      expect(r.changes).toHaveLength(1);
      const [change] = r.changes;
      expect(change!.location?.slideIndex).toBe(slideIndex);
      expect(change!.computed).toEqual(computed as never);
    });
  });

  test('the text changes locate the new text', () => {
    const revenue = comparePresentations(agentPptx(0), agentPptx(1));
    const title = comparePresentations(agentPptx(2), agentPptx(3));
    expect(revenue.ok && revenue.changes[0]!.location).toMatchObject({ shapeId: 4, text: '$4.6M' });
    expect(title.ok && title.changes[0]!.location).toMatchObject({ shapeId: 3 });
  });

  test('the route reads only a step number in digits, 0 through the last step', () => {
    expect(agentStepParam('0')).toBe(0);
    expect(agentStepParam(String(AGENT_STEPS.length))).toBe(AGENT_STEPS.length);
    for (const raw of [null, '', ' 1', '1 ', '1e0', '0x1', '-1', '1.0', '+1', '8', '99999999'])
      expect(agentStepParam(raw)).toBeNull();
  });
});
