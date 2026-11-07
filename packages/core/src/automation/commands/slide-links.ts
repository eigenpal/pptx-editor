// Links to slides of the same presentation (ECMA-376 Part 1, 21.1.2.3.5 `a:hlinkClick` and its
// `action`). A link to one slide relates to the slide part and takes the action
// `ppaction://hlinksldjump`. A link to the next, previous, first, or last slide has no target and
// takes `ppaction://hlinkshowjump?jump=...`. The editor names both in the form of a link in the
// page: `#slide-3` (the slide's number now) and `#next-slide`, `#previous-slide`, `#first-slide`,
// `#last-slide`. The file keeps the relationship, so a link follows its slide when slides move.

import { NS, presentationInfo, REL, sanitizeHref, type NodeSpec } from '../../store/index.ts';
import {
  customShowLinkSpec,
  fileLinkTarget,
  isFileLink,
  parseCustomShowLink,
} from './custom-show-links.ts';
import { CommandRefusal, type OpBatch } from './op-batch.ts';
import { addExternalRelationship, addRelationship } from './package-edits.ts';

/**
 * The `a:hlinkClick` of a link from `part` to `url`, with its relationship: a slide or a custom
 * show in its `#` form, a file in its `./` form (`custom-show-links.ts`), or an address that
 * passes `sanitizeHref`. A bare host ("example.com") is a web address. Anything else is refused
 * with `errors.unsafeLink`, a `#` form that names no slide with `errors.slideNotFound`, one that
 * names no custom show with `errors.customShowNotFound`, and a file path that is not safe with
 * `errors.unsafeFileLink`.
 */
export function linkSpec(batch: OpBatch, part: string, url: string): NodeSpec {
  const raw = url.trim();
  if (!raw || raw.length > 2048) throw new CommandRefusal('errors.unsafeLink');
  if (raw.startsWith('#custom-show-')) {
    const show = parseCustomShowLink(raw);
    if (!show) throw new CommandRefusal('errors.customShowNotFound');
    return customShowLinkSpec(batch.pkg, show);
  }
  if (isFileLink(raw)) {
    // A relative path, written as the file's own hyperlink: no action, never fetched.
    const rId = addExternalRelationship(batch, part, REL.hyperlink, fileLinkTarget(raw));
    return {
      ns: NS.a,
      local: 'hlinkClick',
      attrs: [{ ns: NS.r, local: 'id', prefix: 'r', value: rId }],
    };
  }
  if (raw.startsWith('#')) {
    const target = parseSlideLink(raw);
    if (!target) throw new CommandRefusal('errors.slideNotFound');
    return slideLinkSpec(batch, part, target);
  }
  const withScheme = /^[a-z][a-z0-9+.-]*:/i.test(raw) ? raw : `https://${raw}`;
  const safe = sanitizeHref(withScheme);
  if (!safe.ok || !/^[a-z][a-z0-9+.-]*:/i.test(safe.href))
    throw new CommandRefusal('errors.unsafeLink');
  const rId = addExternalRelationship(batch, part, REL.hyperlink, safe.href);
  return {
    ns: NS.a,
    local: 'hlinkClick',
    attrs: [{ ns: NS.r, local: 'id', prefix: 'r', value: rId }],
  };
}

/** A slide a show jumps to by its place in the show, not by its part. @public */
export type SlideJump = 'next' | 'previous' | 'first' | 'last';

/** A parsed slide link: one slide (0-based) or a jump. */
export type SlideLinkTarget = { readonly slide: number } | { readonly jump: SlideJump };

const SLIDE_JUMP_ACTION = 'ppaction://hlinksldjump';
const SHOW_JUMP_ACTION = 'ppaction://hlinkshowjump?jump=';

const JUMP_OF_KEYWORD: Readonly<Record<string, SlideJump>> = {
  nextslide: 'next',
  previousslide: 'previous',
  firstslide: 'first',
  lastslide: 'last',
};

/** `#slide-3` or `#next-slide`, `#previous-slide`, `#first-slide`, `#last-slide`, or null. */
export function parseSlideLink(url: string): SlideLinkTarget | null {
  const one = /^#slide-([1-9]\d{0,5})$/.exec(url);
  if (one) return { slide: Number(one[1]) - 1 };
  const jump = /^#(next|previous|first|last)-slide$/.exec(url);
  return jump ? { jump: jump[1] as SlideJump } : null;
}

/** The `#` form of a slide link. */
export function slideLinkUrl(target: SlideLinkTarget): string {
  return 'slide' in target ? `#slide-${target.slide + 1}` : `#${target.jump}-slide`;
}

/**
 * The `a:hlinkClick` of a slide link from `part`, with the relationship it needs. A slide that
 * is not in the deck is refused.
 */
export function slideLinkSpec(batch: OpBatch, part: string, target: SlideLinkTarget): NodeSpec {
  if ('jump' in target) {
    const keyword = Object.keys(JUMP_OF_KEYWORD).find((k) => JUMP_OF_KEYWORD[k] === target.jump)!;
    return {
      ns: NS.a,
      local: 'hlinkClick',
      attrs: [
        { ns: NS.r, local: 'id', prefix: 'r', value: '' },
        { ns: '', local: 'action', value: SHOW_JUMP_ACTION + keyword },
      ],
    };
  }
  const slide = presentationInfo(batch.pkg).slides[target.slide];
  if (!slide) throw new CommandRefusal('errors.slideNotFound');
  const rId = addRelationship(batch, part, REL.slide, slide.partName);
  return {
    ns: NS.a,
    local: 'hlinkClick',
    attrs: [
      { ns: NS.r, local: 'id', prefix: 'r', value: rId },
      { ns: '', local: 'action', value: SLIDE_JUMP_ACTION },
    ],
  };
}

/**
 * The `#` form of a run's slide link, from its `action` and the part its relationship names, or
 * undefined when it is no slide link or names no slide of the deck.
 */
export function slideLinkOf(
  action: string | undefined,
  targetPart: string | undefined,
  slideParts: readonly string[]
): string | undefined {
  if (!action) return undefined;
  if (action.startsWith(SHOW_JUMP_ACTION)) {
    const keyword = action.slice(SHOW_JUMP_ACTION.length);
    // An own key only: a file's `jump=constructor` names no jump.
    return Object.hasOwn(JUMP_OF_KEYWORD, keyword)
      ? slideLinkUrl({ jump: JUMP_OF_KEYWORD[keyword]! })
      : undefined;
  }
  if (action !== SLIDE_JUMP_ACTION || !targetPart) return undefined;
  const index = slideParts.indexOf(targetPart);
  return index < 0 ? undefined : slideLinkUrl({ slide: index });
}
