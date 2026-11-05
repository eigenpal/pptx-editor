// The slide and section slots of the chrome registry (`CHROME_GROUPS`): the Home >
// Slides group, Slide Show > Hide Slide, and Home > Section. Kept apart from
// `chrome-controls.ts` so that file stays under the line cap; the registry spreads them in.

import type { EditorCommand, EditorSnapshot } from '../contracts/index.ts';
import type { ChromeControl, ChromeGroup, ChromeSlotId } from './chrome-controls.ts';
import { ICON_PATHS } from './icon-paths.ts';

const icon = (name: string): readonly string[] => ICON_PATHS[name] ?? [];

/** New slide, Layout, Reset, and Hide: controls of the `slide` group after Format Background. */
export const SLIDE_STRUCTURE_CONTROLS: readonly ChromeControl[] = [
  // The Home > Slides group and Slide Show > Hide Slide. Menus, the rail, and a
  // ribbon place these slots; the formatting bar does not show them.
  {
    // A split button: the button adds a slide, the gallery picks its layout.
    id: 'add',
    shape: 'gallery',
    labelKey: 'toolbar.newSlide',
    paths: icon('add'),
    state: { kind: 'value' },
  },
  {
    // A gallery of the deck's layouts, each drawn by the engine. The slot's value is the
    // current slide's layout, which the gallery marks.
    id: 'layout',
    shape: 'gallery',
    labelKey: 'toolbar.slideLayout',
    paths: icon('dashboard'),
    state: { kind: 'value' },
  },
  {
    id: 'reset',
    labelKey: 'toolbar.resetSlide',
    paths: icon('restart_alt'),
    state: { kind: 'command' },
  },
  {
    id: 'hide',
    labelKey: 'toolbar.hideSlide',
    toggle: 'pressed',
    paths: icon('visibility_off'),
    state: { kind: 'command' },
  },
];

/** The `section` group: Home > Section, one slot per row of its menu. */
export const SECTION_GROUP: ChromeGroup = {
  // Sections of the slide list (Home > Section): one slot per row of its menu.
  // The group is named after the section object; a ribbon can show it as one menu button.
  id: 'section',
  labelKey: 'toolbar.groups.section',
  contextual: true,
  controls: [
    {
      id: 'add',
      labelKey: 'sections.add',
      paths: icon('view_agenda'),
      state: { kind: 'command' },
    },
    {
      id: 'rename',
      labelKey: 'sections.rename',
      paths: icon('edit_note'),
      state: { kind: 'command' },
    },
    {
      id: 'remove',
      labelKey: 'sections.remove',
      paths: icon('remove'),
      state: { kind: 'command' },
    },
    {
      id: 'removeAll',
      labelKey: 'sections.removeAll',
      paths: icon('delete'),
      state: { kind: 'command' },
    },
    {
      id: 'collapseAll',
      labelKey: 'sections.collapseAll',
      paths: icon('unfold_less'),
      state: { kind: 'command' },
    },
    {
      id: 'expandAll',
      labelKey: 'sections.expandAll',
      paths: icon('unfold_more'),
      state: { kind: 'command' },
    },
  ],
};

/** The section of the current slide, or undefined. */
function currentSection(snap: EditorSnapshot) {
  return (snap.sections ?? []).find(
    (x) => snap.currentSlide >= x.firstSlide && snap.currentSlide < x.firstSlide + x.slideCount
  );
}

/**
 * The command that probes a slide or section slot whose probe depends on the deck, or null for
 * every other slot (`toolbarCommandState`). @internal
 */
export function slideSlotProbe(id: ChromeSlotId, snap: EditorSnapshot): EditorCommand | null {
  if (id === 'slide.layout') {
    // A layout pick is probed with a layout the deck lists: the slide's own one when its master
    // lists it (§19.3.1.41 `sldLayoutIdLst`), else the first one. An unlisted layout is refused.
    // The probe sets `enabled` only: the slot's `value` is always the slide's own layout.
    const own = snap.slides[snap.currentSlide]?.layoutId;
    const probe = snap.layouts?.some((l) => l.id === own) ? own : snap.layouts?.[0]?.id;
    return probe ? { type: 'applyLayout', layoutId: probe } : { type: 'resetSlides' };
  }
  if (id === 'section.rename') {
    // Rename Section needs a name: it is probed with a fixed valid name, because the section's
    // own name can come from a file and be invalid.
    const own = currentSection(snap);
    return own
      ? { type: 'renameSection', sectionId: own.id, name: 'x' }
      : { type: 'removeSection' };
  }
  return null;
}

/** Whether a slide slot shows as pressed: Hide Slide when every selected slide is hidden. @internal */
export function slideSlotActive(id: ChromeSlotId, snap: EditorSnapshot): boolean {
  return (
    id === 'slide.hide' &&
    snap.selectedSlides.length > 0 &&
    snap.selectedSlides.every((i) => !!snap.slides[i]?.hidden)
  );
}

/** The value of a slide or view slot: the current slide's layout, or the view. @internal */
export function slideSlotValue(id: ChromeSlotId, snap: EditorSnapshot): string | undefined {
  if (id === 'slide.layout') return snap.slides[snap.currentSlide]?.layoutId;
  if (id === 'view.presentationView') return snap.presentationView ?? 'normal';
  return undefined;
}

/** The fixed commands of the slide and section command slots. @internal */
export const SLIDE_SLOT_COMMANDS: Partial<Record<ChromeSlotId, EditorCommand>> = {
  'slide.add': { type: 'addSlide' },
  'slide.reset': { type: 'resetSlides' },
  'slide.hide': { type: 'setSlidesHidden' },
  'section.add': { type: 'addSection' },
  'section.remove': { type: 'removeSection' },
  'section.removeAll': { type: 'removeAllSections' },
  'section.collapseAll': { type: 'setSectionsCollapsed', collapsed: true },
  'section.expandAll': { type: 'setSectionsCollapsed', collapsed: false },
};

/**
 * Slots that adapter chrome runs, not one engine command: the refusal names the public way to
 * run them, as DOCX Editor's dialog slots do (`DIALOG_SLOT_REASONS`). @internal
 */
export const CHROME_SLOT_REASONS: Partial<Record<ChromeSlotId, string>> = {
  'section.rename':
    'renaming a section is not a toolbar command; call requestSectionRename(editor, sectionId) ' +
    'from the adapter, or run renameSection with a name',
};
