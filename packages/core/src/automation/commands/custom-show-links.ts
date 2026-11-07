// Links to custom shows and to other files.
//
// A custom show (ECMA-376 Part 1, 19.2.1.7 `custShow`) is a named list of the deck's slides in
// `p:presentation/p:custShowLst`. A link to one has no relationship (`r:id=""`) and takes the
// action `ppaction://customshow?id=<id>`, with `&return=true` when the show returns to the
// linking slide after its last slide (ECMA-376 Part 1, 21.1.2.3.5 `hlinkClick`, `action`). The
// editor names it in the `#` form of a link in the page: `#custom-show-<id>`, or
// `#custom-show-<id>?return=true`.
//
// A link to another file is an external hyperlink relationship whose target is a relative path,
// such as `report.pdf`. The editor names it `./report.pdf` (or `../shared/report.pdf`), so a
// bare host such as `example.com` stays a web address. A path is checked strictly: no scheme,
// no absolute path, no backslash, no control character, at most two leading `../` steps, and
// no program or script file. The editor never opens such a link: a browser cannot reach the
// file. Links to programs (`ppaction://program`) are never written.

import type { CustomShowSummary } from '../../contracts/neutral.ts';
import { cleanText } from './text-units.ts';
import {
  attr,
  child,
  elements,
  NS,
  presentationInfo,
  xmlPartOf,
  type NodeSpec,
  type PptxPackage,
} from '../../store/index.ts';
import { CommandRefusal } from './op-batch.ts';

/** A parsed custom show link. */
export interface CustomShowLinkTarget {
  readonly id: number;
  /** Return to the linking slide after the show's last slide. */
  readonly returns: boolean;
}

const MAX_SHOWS = 1000;
const CUSTOM_SHOW_ACTION = 'ppaction://customshow?';

/** The custom shows of a deck, in file order: each with an `xsd:unsignedInt` id and a name. */
export function deckCustomShows(pkg: PptxPackage): CustomShowSummary[] {
  const root = xmlPartOf(pkg, presentationInfo(pkg).presentationPart)?.root;
  const out: CustomShowSummary[] = [];
  for (const show of elements(child(root, 'p:custShowLst'))) {
    if (show.ns !== NS.p || show.local !== 'custShow') continue;
    const raw = attr(show, 'id') ?? '';
    if (!/^\d{1,10}$/.test(raw) || Number(raw) > 4_294_967_295) continue;
    const name = cleanText(attr(show, 'name') ?? '')
      .replace(/[\n\u000B]/g, ' ')
      .trim()
      .slice(0, 255);
    out.push({ id: Number(raw), name });
    if (out.length >= MAX_SHOWS) break;
  }
  return out;
}

/** Custom show summaries for the snapshot, reference-stable while the shows stay the same. */
export function createCustomShowSummaries(): (
  pkg: PptxPackage | undefined
) => readonly CustomShowSummary[] {
  let last: readonly CustomShowSummary[] = [];
  return (pkg) => {
    const next = pkg ? deckCustomShows(pkg) : [];
    const same =
      next.length === last.length &&
      next.every((s, i) => s.id === last[i].id && s.name === last[i].name);
    if (!same) last = next;
    return last;
  };
}

/** `#custom-show-3` or `#custom-show-3?return=true`, or null. */
export function parseCustomShowLink(url: string): CustomShowLinkTarget | null {
  const m = /^#custom-show-(\d{1,10})(\?return=true)?$/.exec(url);
  if (!m || Number(m[1]) > 4_294_967_295) return null;
  return { id: Number(m[1]), returns: !!m[2] };
}

/** The `#` form of a custom show link. */
export function customShowLinkUrl(target: CustomShowLinkTarget): string {
  return `#custom-show-${target.id}${target.returns ? '?return=true' : ''}`;
}

/** The `a:hlinkClick` of a link to a custom show of the deck; another id is refused. */
export function customShowLinkSpec(pkg: PptxPackage, target: CustomShowLinkTarget): NodeSpec {
  if (!deckCustomShows(pkg).some((s) => s.id === target.id))
    throw new CommandRefusal('errors.customShowNotFound');
  const action = `${CUSTOM_SHOW_ACTION}id=${target.id}${target.returns ? '&return=true' : ''}`;
  return {
    ns: NS.a,
    local: 'hlinkClick',
    attrs: [
      { ns: NS.r, local: 'id', prefix: 'r', value: '' },
      { ns: '', local: 'action', value: action },
    ],
  };
}

/** The `#` form of a run's custom show link, from its `action`, or undefined. */
export function customShowLinkOf(action: string | undefined): string | undefined {
  if (!action?.startsWith(CUSTOM_SHOW_ACTION)) return undefined;
  let id: number | undefined;
  let returns = false;
  for (const field of action.slice(CUSTOM_SHOW_ACTION.length).split('&')) {
    const [key, value = ''] = field.split('=');
    if (key === 'id' && /^\d{1,10}$/.test(value) && Number(value) <= 4_294_967_295)
      id = Number(value);
    else if (key === 'return') returns = value === 'true' || value === '1';
  }
  return id === undefined ? undefined : customShowLinkUrl({ id, returns });
}

/** Extensions of programs and scripts: a file link never names one. */
const PROGRAM_EXTENSIONS = new Set([
  'app',
  'bat',
  'cmd',
  'com',
  'cpl',
  'dll',
  'exe',
  'gadget',
  'hta',
  'inf',
  'jar',
  'js',
  'jse',
  'lnk',
  'msc',
  'msi',
  'msp',
  'pif',
  'ps1',
  'psm1',
  'reg',
  'scf',
  'scr',
  'sh',
  'url',
  'vb',
  'vbe',
  'vbs',
  'ws',
  'wsc',
  'wsf',
  'wsh',
  'command',
  'pkg',
  'dmg',
  'desktop',
  'appimage',
]);
const MAX_UP_STEPS = 2;
/** One path segment as a person types it: letters, digits, spaces, and common punctuation. */
const SEGMENT = /^[\p{L}\p{N} ._()~+,'&@!$=-]+$/u;

/** Whether a link draft is in the file form: `./name` or `../name`. */
export function isFileLink(url: string): boolean {
  return url.startsWith('./') || url.startsWith('../');
}

/**
 * The relationship target of a file link draft (`./docs/report.pdf` gives `docs/report.pdf`),
 * percent-encoded segment by segment, or a refusal with `errors.unsafeFileLink`.
 */
export function fileLinkTarget(url: string): string {
  const bad = (): never => {
    throw new CommandRefusal('errors.unsafeFileLink');
  };
  if (url.length > 1024 || !isFileLink(url)) bad();
  const segments = url.split('/');
  let up = 0;
  while (segments[0] === '..') {
    segments.shift();
    up += 1;
  }
  if (up === 0 && segments[0] === '.') segments.shift();
  if (up > MAX_UP_STEPS || segments.length === 0) bad();
  for (const s of segments)
    if (!SEGMENT.test(s) || s === '.' || s === '..' || s.trim() !== s) bad();
  const name = segments[segments.length - 1];
  const dot = name.lastIndexOf('.');
  if (dot > 0 && PROGRAM_EXTENSIONS.has(name.slice(dot + 1).toLowerCase())) bad();
  return [...Array<string>(up).fill('..'), ...segments.map(encodeURIComponent)].join('/');
}

/**
 * The file form of an external hyperlink target that is a relative path (`report.pdf` reads
 * `./report.pdf`), or undefined for any other target: a URL with a scheme, an absolute path, a
 * fragment, or a path with a backslash.
 */
export function fileLinkUrl(target: string | undefined): string | undefined {
  if (!target || /^[a-z][a-z0-9+.-]*:/i.test(target) || /^[/#?\\]/.test(target)) return undefined;
  if (target.includes('\\') || /\p{Cc}/u.test(target)) return undefined;
  let text = target;
  try {
    text = decodeURIComponent(target);
  } catch {
    // A malformed escape keeps the target as written.
  }
  return text.startsWith('../') ? text : `./${text}`;
}
