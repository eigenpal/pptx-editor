// The automation `readTransition` and `setTransition` operations over `p:transition`
// (`CT_SlideTransition`, ECMA-376 Part 1, §19.3.1.50).

import { describe, expect, test } from 'bun:test';
import { buildPptx } from '../../../../../scripts/fixtures/builder.ts';
import { createPptxEditor } from '../../editor/index.ts';
import { presentationInfo, readPptx, serializeXmlPart, xmlPartOf } from '../../store/index.ts';
import {
  createServerAutomationHost,
  type AutomationBatchResponse,
  type AutomationHandle,
  type AutomationHost,
  type AutomationOperation,
  type AutomationValue,
} from '../index.ts';

const SOUND = '<p:sndAc><p:endSnd/></p:sndAc>';
const EXT = '<p:extLst><p:ext uri="{00000000-0000-0000-0000-000000000176}"/></p:extLst>';

/** A deck of one slide whose slide XML after `clrMapOvr` is `ext`. */
const deckWith = (ext = ''): Uint8Array =>
  buildPptx({ slides: [{ layout: 'blank', shapes: '', ext }] });

/** The saved XML of the first slide. */
function slideXml(bytes: Uint8Array): string {
  const r = readPptx(bytes);
  if (!r.ok) throw new Error('reopen');
  return serializeXmlPart(xmlPartOf(r.pkg, presentationInfo(r.pkg).slides[0]!.partName)!);
}

/** The saved XML of the first slide of the host's presentation. */
function saved(host: AutomationHost): string {
  const result = host.save();
  if (!result.ok) throw new Error(JSON.stringify(result.error));
  return slideXml(result.bytes);
}

function open(bytes = deckWith()): AutomationHost {
  const opened = createServerAutomationHost(bytes);
  if (!opened.ok) throw new Error(JSON.stringify(opened.error));
  return opened.host;
}

function exec(host: AutomationHost, operations: AutomationOperation[]): AutomationBatchResponse {
  return host.execute({ operations });
}

/** Run a batch that must succeed and return its values. */
function ok(host: AutomationHost, operations: AutomationOperation[]): AutomationValue[] {
  const response = exec(host, operations);
  if (!response.ok) throw new Error(JSON.stringify(response.results));
  return response.results.map((r) => (r.status === 'ok' ? r.value : { kind: 'null' as const }));
}

function firstSlide(host: AutomationHost): AutomationHandle {
  const [value] = ok(host, [{ op: 'listSlides' }]);
  if (value?.kind !== 'handles') throw new Error('no slides');
  return value.handles[0]!;
}

describe('automation host: slide transitions', () => {
  test('reads, sets, and clears a transition', () => {
    const host = open();
    const slide = firstSlide(host);
    const read = () =>
      (ok(host, [{ op: 'readTransition', slide }])[0] as { transition: unknown }).transition;
    expect(read()).toEqual({
      kind: 'none',
      durationMs: 750,
      advanceOnClick: true,
      advanceAfterMs: null,
      element: null,
      direction: null,
      throughBlack: false,
    });
    ok(host, [{ op: 'setTransition', slide, transition: { kind: 'fade', durationMs: 1200 } }]);
    ok(host, [{ op: 'setTransition', slide, transition: { advanceAfterMs: 3000 } }]);
    expect(read()).toEqual({
      kind: 'fade',
      durationMs: 1200,
      advanceOnClick: true,
      advanceAfterMs: 3000,
      element: 'fade',
      direction: null,
      throughBlack: false,
    });
    ok(host, [{ op: 'setTransition', slide, transition: { kind: 'wipe', direction: 'd' } }]);
    expect(read()).toMatchObject({ kind: 'wipe', element: 'wipe', direction: 'd' });
    ok(host, [{ op: 'setTransition', slide, transition: { kind: 'fade' } }]);
    // The same settings again write nothing, and answer applied.
    const revision = exec(host, [{ op: 'readTransition', slide }]).revision;
    ok(host, [{ op: 'setTransition', slide, transition: { kind: 'fade' } }]);
    expect(exec(host, [{ op: 'readTransition', slide }]).revision).toBe(revision);
    ok(host, [
      {
        op: 'setTransition',
        slide,
        transition: { kind: 'none', advanceOnClick: true, advanceAfterMs: null },
      },
    ]);
    expect(read()).toMatchObject({ kind: 'none', advanceAfterMs: null });
  });

  test('refuses bad arguments, and the batch writes nothing', () => {
    const host = open();
    const slide = firstSlide(host);
    const refusals: [unknown, string][] = [
      [{ kind: 'dissolve' }, 'transitionKind'],
      [{ kind: 'other' }, 'transitionKind'],
      [{ durationMs: 9 }, 'durationMs'],
      [{ durationMs: 59_001 }, 'durationMs'],
      [{ durationMs: Number.POSITIVE_INFINITY }, 'durationMs'],
      [{ advanceOnClick: 1 }, 'advanceOnClick'],
      [{ advanceAfterMs: -1 }, 'advanceAfterMs'],
      [{ advanceAfterMs: 3_600_001 }, 'advanceAfterMs'],
      [{ direction: 'left' }, 'direction'],
      [{ throughBlack: 'yes' }, 'throughBlack'],
      [{ spd: 'fast' }, 'transition'],
      [null, 'transition'],
    ];
    for (const [transition, detail] of refusals) {
      const refused = exec(host, [
        { op: 'setTransition', slide, transition: { kind: 'cut' } },
        { op: 'setTransition', slide, transition } as unknown as AutomationOperation,
      ]);
      expect(refused.results[1]).toMatchObject({ error: { code: 'invalid-argument', detail } });
    }
    expect(ok(host, [{ op: 'readTransition', slide }])[0]).toMatchObject({
      transition: { kind: 'none' },
    });
    // No effect: a duration has nothing to keep it.
    const noEffect = exec(host, [{ op: 'setTransition', slide, transition: { durationMs: 500 } }]);
    expect(noEffect.results[0]).toMatchObject({
      error: { code: 'unsupported-content', detail: 'transition-no-effect' },
    });
    // A direction belongs to push, wipe, and split; through black to fade and cut.
    for (const transition of [{ direction: 'l' }, { kind: 'push', throughBlack: true }] as const) {
      const refused = exec(host, [{ op: 'setTransition', slide, transition }]);
      expect(refused.results[0]).toMatchObject({
        error: { code: 'unsupported-content', detail: 'transition-variant' },
      });
    }
  });

  test('a write that does not name the kind keeps a transition that holds only a sound', () => {
    const host = open(deckWith(`<p:transition advTm="3000">${SOUND}</p:transition>`));
    const slide = firstSlide(host);
    ok(host, [{ op: 'setTransition', slide, transition: { advanceAfterMs: null } }]);
    expect(saved(host)).toContain(`<p:transition>${SOUND}</p:transition>`);
    ok(host, [{ op: 'setTransition', slide, transition: { advanceOnClick: false } }]);
    expect(saved(host)).toContain(SOUND);
  });

  test('a new effect keeps the sound and the extensions; an explicit none drops them', () => {
    const host = open(
      deckWith(`<p:transition advTm="3000"><p:fade/>${SOUND}${EXT}</p:transition>`)
    );
    const slide = firstSlide(host);
    ok(host, [{ op: 'setTransition', slide, transition: { kind: 'push', durationMs: 900 } }]);
    const pushed = saved(host);
    expect(
      pushed.match(new RegExp(`<p:push dir="u"/>${SOUND}${EXT.replace(/[{}]/g, '\\$&')}`, 'g'))
    ).toHaveLength(2);
    ok(host, [{ op: 'setTransition', slide, transition: { kind: 'none' } }]);
    expect(saved(host)).toContain('<p:transition advTm="3000"/>');
    ok(host, [{ op: 'setTransition', slide, transition: { kind: 'none', advanceAfterMs: null } }]);
    expect(saved(host)).not.toContain('transition');
  });

  test("the editor's own command keeps a sound when the pane clears the advance time", async () => {
    // The animations pane sends `setTransition` with `advanceAfterMs: null` for its check box.
    const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    expect(
      (await editor.load(deckWith(`<p:transition advTm="3000">${SOUND}</p:transition>`))).ok
    ).toBe(true);
    expect(editor.exec({ type: 'setTransition', advanceAfterMs: null }).ok).toBe(true);
    expect(slideXml(editor.save()!)).toContain(`<p:transition>${SOUND}</p:transition>`);
  });
});
