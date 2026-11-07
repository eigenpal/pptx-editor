// The link popup's edit form, framework-free: a link goes to a web address, an email address,
// or a slide of this presentation. The form keeps one draft string, the URL that `setLink`
// takes: `mailto:` for email, and the `#` form for a slide (`#slide-3`, `#next-slide`). So the
// controller and the command stay the same for every kind. `@internal`: adapter glue.

import type { Translate } from '@pptx-editor.dev/i18n';
import { parseSlideLink } from '../automation/commands/slide-links.ts';

/** What a link goes to. @internal */
export type LinkKind = 'web' | 'email' | 'slide';

/** The kinds in the order the form offers them. @internal */
export const LINK_KINDS: readonly LinkKind[] = ['web', 'email', 'slide'];

/** The kind of a link draft: `mailto:` is email, the `#` form is a slide, else a web address. */
export function linkKindOf(url: string): LinkKind {
  const raw = url.trim();
  if (/^mailto:/i.test(raw)) return 'email';
  return raw.startsWith('#') ? 'slide' : 'web';
}

/** The address and subject of a `mailto:` draft; any other draft gives empty fields. */
export function mailtoParts(url: string): { address: string; subject: string } {
  const raw = url.trim();
  if (!/^mailto:/i.test(raw)) return { address: '', subject: '' };
  const [to = '', query = ''] = raw.slice('mailto:'.length).split('?', 2);
  const subject = new URLSearchParams(query).get('subject') ?? '';
  let address = to;
  try {
    address = decodeURIComponent(to);
  } catch {
    // A malformed escape keeps the text as written.
  }
  return { address, subject };
}

/**
 * The `mailto:` draft of an address and a subject. The subject is percent-encoded, so its text
 * cannot add other fields. An empty address gives an empty draft, which Apply cannot send.
 */
export function mailtoUrl(address: string, subject: string): string {
  const to = address.trim();
  if (!to) return '';
  const s = subject.trim();
  return `mailto:${to}${s ? `?subject=${encodeURIComponent(s)}` : ''}`;
}

/** The slide targets the form lists: the four jumps, then each slide by number. */
export function slideLinkChoices(slideCount: number): readonly string[] {
  const out = ['#first-slide', '#last-slide', '#next-slide', '#previous-slide'];
  for (let i = 1; i <= slideCount; i += 1) out.push(`#slide-${i}`);
  return out;
}

/** The words for a slide link, such as "Slide 3" or "Next slide"; '' for any other draft. */
export function slideLinkLabel(t: Translate, url: string): string {
  const target = parseSlideLink(url);
  if (!target) return '';
  if ('slide' in target) return t('link.slideNumber', { number: target.slide + 1 });
  const keys = {
    next: 'link.nextSlide',
    previous: 'link.previousSlide',
    first: 'link.firstSlide',
    last: 'link.lastSlide',
  } as const;
  return t(keys[target.jump]);
}

/**
 * The slide (0-based) a slide link goes to from slide `current` of `count`, or null for a draft
 * that is no slide link or a slide past the end.
 */
export function slideLinkIndex(url: string, current: number, count: number): number | null {
  const target = parseSlideLink(url);
  if (!target || count <= 0) return null;
  if ('slide' in target) return target.slide < count ? target.slide : null;
  const last = count - 1;
  if (target.jump === 'first') return 0;
  if (target.jump === 'last') return last;
  return Math.max(0, Math.min(last, current + (target.jump === 'next' ? 1 : -1)));
}
