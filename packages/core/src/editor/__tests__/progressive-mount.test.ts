// The rail's mount limit while a long deck opens: the first screen with the first slide, more
// after each painted frame, the current slide always in, and every item at once for a deck
// that opened whole. A grouped rail mounts whole sections.

import { describe, expect, test } from 'bun:test';
import {
  createProgressiveMount,
  MOUNT_CHUNK_ITEMS,
  MOUNT_FIRST_ITEMS,
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

  test('an open mounts the first screen, then grows a chunk at a time until every item is in', async () => {
    let grown = 0;
    const mount = createProgressiveMount(() => (grown += 1));
    expect(mount.limit(5000, 0, true)).toBe(MOUNT_FIRST_ITEMS);
    mount.follow(5000);
    mount.follow(5000); // One growth waits at a time.
    await tick();
    expect(grown).toBe(1);
    expect(mount.limit(5000, 0, true)).toBe(MOUNT_FIRST_ITEMS + MOUNT_CHUNK_ITEMS);
    // The open can end before the rail is whole: it keeps growing.
    for (let i = 0; i < 20; i += 1) {
      mount.follow(5000);
      await tick();
    }
    expect(mount.limit(5000, 0, false)).toBe(5000);
    mount.dispose();
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

  test('a section longer than one chunk still mounts chunk by chunk under its header', () => {
    const grouped = [section('a', 0, 10), section('b', 10, 4990)];
    const shown = mountedSections(grouped, MOUNT_FIRST_ITEMS);
    expect(shown.limit).toBe(MOUNT_FIRST_ITEMS);
    expect(shown.sections.map((s) => s.id)).toEqual(['a', 'b']);
  });
});
