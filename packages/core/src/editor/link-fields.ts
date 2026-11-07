// The link popup's edit form, framework-free: a link goes to a web address, an email address,
// a slide or a custom show of this presentation, or another file. The form keeps one draft
// string, the URL that `setLink` takes: `mailto:` for email, the `#` form for a slide
// (`#slide-3`, `#next-slide`) and a custom show (`#custom-show-2?return=true`), and the `./`
// form for a file (`./report.pdf`). So the controller and the command stay the same for every
// kind. `@internal`: adapter glue.

import type { Translate } from '@pptx-editor.dev/i18n';
import type { CustomShowSummary } from '../contracts/index.ts';
import {
  customShowLinkUrl,
  isFileLink,
  parseCustomShowLink,
} from '../automation/commands/custom-show-links.ts';
import { parseSlideLink } from '../automation/commands/slide-links.ts';

/** What a link goes to. @internal */
export type LinkKind = 'web' | 'email' | 'slide' | 'customShow' | 'file';

/** The kinds in the order the form offers them. @internal */
export const LINK_KINDS: readonly LinkKind[] = ['web', 'email', 'slide', 'customShow', 'file'];

/**
 * The kinds the form offers for a deck: Custom show only when the deck has one, or when the
 * draft already links to one. @internal
 */
export function linkKindsFor(
  shows: readonly CustomShowSummary[],
  draft: string
): readonly LinkKind[] {
  if (shows.length > 0 || linkKindOf(draft) === 'customShow') return LINK_KINDS;
  return LINK_KINDS.filter((k) => k !== 'customShow');
}

/**
 * The kind of a link draft: `mailto:` is email, `#custom-show-` a custom show, any other `#`
 * form a slide, `./` and `../` a file, else a web address.
 */
export function linkKindOf(url: string): LinkKind {
  const raw = url.trim();
  if (/^mailto:/i.test(raw)) return 'email';
  if (raw.startsWith('#custom-show-')) return 'customShow';
  if (isFileLink(raw)) return 'file';
  return raw.startsWith('#') ? 'slide' : 'web';
}

/** The first draft of a kind, when the form switches to it. @internal */
export function linkKindDraft(kind: LinkKind, shows: readonly CustomShowSummary[]): string {
  if (kind === 'slide') return '#next-slide';
  if (kind === 'customShow')
    return shows[0] ? customShowLinkUrl({ id: shows[0].id, returns: true }) : '';
  return '';
}

/** The custom show and the return choice of a custom show draft, or null. @internal */
export function customShowDraft(
  url: string
): { readonly id: number; readonly returns: boolean } | null {
  return parseCustomShowLink(url.trim());
}

/** The draft of a custom show and a return choice. @internal */
export function customShowDraftUrl(id: number, returns: boolean): string {
  return customShowLinkUrl({ id, returns });
}

/**
 * The words for a custom show link, such as "Custom show: Short", or '' for any other draft.
 * A show the deck does not have shows its id. @internal
 */
export function customShowLinkLabel(
  t: Translate,
  url: string,
  shows: readonly CustomShowSummary[]
): string {
  const target = parseCustomShowLink(url.trim());
  if (!target) return '';
  const show = shows.find((s) => s.id === target.id);
  return t('link.customShowName', { name: show?.name || String(target.id) });
}

/** The fields of an email link the form edits, and the fields it keeps as they are. */
export interface MailtoParts {
  readonly address: string;
  readonly subject: string;
  /** The other fields of the link (`cc`, `body`), still encoded, joined with `&`. */
  readonly rest: string;
}

/** Decode a `mailto:` part as RFC 6068 writes it: `+` stays a plus sign. */
function decodePart(text: string): string {
  try {
    return decodeURIComponent(text);
  } catch {
    // A malformed escape keeps the text as written.
    return text;
  }
}

/** The address, subject, and other fields of a `mailto:` draft; any other draft gives none. */
export function mailtoParts(url: string): MailtoParts {
  const raw = url.trim();
  if (!/^mailto:/i.test(raw)) return { address: '', subject: '', rest: '' };
  const body = raw.slice('mailto:'.length);
  const q = body.indexOf('?');
  const to = q < 0 ? body : body.slice(0, q);
  let subject = '';
  const rest: string[] = [];
  for (const field of q < 0 ? [] : body.slice(q + 1).split('&')) {
    if (!field) continue;
    const eq = field.indexOf('=');
    const key = (eq < 0 ? field : field.slice(0, eq)).toLowerCase();
    if (key === 'subject' && !subject) subject = decodePart(field.slice(eq + 1));
    else rest.push(field);
  }
  return { address: decodePart(to), subject, rest: rest.join('&') };
}

/**
 * The `mailto:` draft of an address, a subject, and the other fields of the link. The address
 * and the subject are percent-encoded where needed, so their text cannot add other fields. An
 * empty address gives an empty draft, which Apply cannot send.
 */
export function mailtoUrl(address: string, subject: string, rest = ''): string {
  const to = address.trim().replace(/[%?#&\s]/gu, (c) => encodeURIComponent(c));
  if (!to) return '';
  const s = subject.trim();
  const fields = [...(s ? [`subject=${encodeURIComponent(s)}`] : []), ...(rest ? [rest] : [])];
  return `mailto:${to}${fields.length ? `?${fields.join('&')}` : ''}`;
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
