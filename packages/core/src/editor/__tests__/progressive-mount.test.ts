// The rail's mount limit while a long deck opens: the first screen with the first slide, more
// after each painted frame, the current slide always in, and every item at once for a deck
// that opened whole. A grouped rail mounts whole sections.

import { describe, expect, test } from 'bun:test';
import {
  createProgressiveMount,
  MOUNT_CHUNK_ITEMS,
  MOUNT_FIRST_ITEMS,
  MOUNT_FIRST_STEP_ITEMS,
  MOUNT_FRAME_MS,
  MOUNT_MIN_STEP_ITEMS,
  mountedSections,
} from '../progressive-mount.ts';
import type { SlideSection } from '../../contracts/index.ts';

const section = (id: string, firstSlide: number, slideCount: number): SlideSection =>
  ({ id, name: id, firstSlide, slideCount }) as SlideSection;

const tick = (): Promise<void> => new Promise((r) => setTimeout(r, 5));

describe('the rail mount limit', () => {
  test('a deck that opened whole mounts every item', () => {
    const mount = createProgressiveMount(() => {});
    expect(mount.limit(5000, 0, false)).toBe(5000);
    mount.dispose();
  });

  test('an open mounts the first screen, then grows a step at a time until every item is in', async () => {
    let grown = 0;
    const mount = createProgressiveMount(() => (grown += 1));
    expect(mount.limit(5000, 0, true)).toBe(MOUNT_FIRST_ITEMS);
    mount.follow(5000);
    mount.follow(5000); // One growth waits at a time.
    await tick();
    expect(grown).toBe(1);
    expect(mount.limit(5000, 0, true)).toBe(MOUNT_FIRST_ITEMS + MOUNT_FIRST_STEP_ITEMS);
    // The open can end before the rail is whole: it keeps growing.
    for (let i = 0; i < 200 && mount.limit(5000, 0, false) < 5000; i += 1) {
      mount.follow(5000);
      await tick();
    }
    expect(mount.limit(5000, 0, false)).toBe(5000);
    mount.dispose();
  });

  test('each step mounts what fits in the frame budget, past the cost every render pays', async () => {
    /** The steps a rail takes when a render costs `fixed` ms plus `per` ms an item. */
    const steps = async (fixed: number, per: number, count: number): Promise<number[]> => {
      let clock = 0;
      const mount = createProgressiveMount(
        () => {},
        () => clock
      );
      let at = mount.limit(50_000, 0, true);
      const out: number[] = [];
      for (let i = 0; i < count; i += 1) {
        mount.follow(50_000); // Times the render of the step before, and waits for a frame.
        await tick();
        const next = mount.limit(50_000, 0, true);
        out.push(next - at);
        clock += fixed + (next - at) * per; // The render of the added items.
        at = next;
      }
      mount.dispose();
      return out;
    };
    // 20 ms for the list and 0.1 ms an item: 8 ms holds 80 more items, whatever the list costs.
    expect(await steps(20, 0.1, 4)).toEqual([MOUNT_FIRST_STEP_ITEMS, 26, 80, 80]);
    // Items that cost nothing: the step goes to its cap.
    expect(await steps(5, 0, 3)).toEqual([MOUNT_FIRST_STEP_ITEMS, 160, MOUNT_CHUNK_ITEMS]);
    // Very slow items: the step never stops the rail.
    expect((await steps(0, 5, 2)).at(-1)).toBe(MOUNT_MIN_STEP_ITEMS);
    expect(MOUNT_FRAME_MS).toBe(8);
  });

  test('the current slide is always mounted, and a new open starts at the first screen again', () => {
    const mount = createProgressiveMount(() => {});
    expect(mount.limit(5000, 3000, true)).toBe(3001);
    expect(mount.limit(5000, 10, false)).toBe(3001);
    expect(mount.limit(800, 0, true)).toBe(MOUNT_FIRST_ITEMS);
    mount.dispose();
  });

  test('dispose cancels a waiting growth', async () => {
    let grown = 0;
    const mount = createProgressiveMount(() => (grown += 1));
    mount.limit(5000, 0, true);
    mount.follow(5000);
    mount.dispose();
    await tick();
    expect(grown).toBe(0);
  });

  test('a grouped rail mounts whole sections and no header without its slides', () => {
    const grouped = [
      section('a', 0, 30),
      section('b', 30, 15),
      section('empty', 45, 0),
      section('c', 45, 255),
    ];
    const shown = mountedSections(grouped, MOUNT_FIRST_ITEMS);
    // `b` holds item 39, so its other slides mount too; `c` starts past the limit and waits.
    expect(shown.limit).toBe(45);
    expect(shown.sections.map((s) => s.id)).toEqual(['a', 'b', 'empty']);
    expect(mountedSections(grouped, 300).sections).toHaveLength(4);
  });

  test('a long section still mounts step by step under its header', () => {
    const grouped = [section('a', 0, 10), section('b', 10, 4990)];
    const shown = mountedSections(grouped, MOUNT_FIRST_ITEMS);
    expect(shown.limit).toBe(MOUNT_FIRST_ITEMS);
    expect(shown.sections.map((s) => s.id)).toEqual(['a', 'b']);
  });
});
