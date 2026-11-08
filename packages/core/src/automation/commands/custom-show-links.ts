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
// only a document, picture, sound, or video file. The editor never opens such a link: a browser cannot reach the
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
  const seen = new Set<number>();
  for (const show of elements(child(root, 'p:custShowLst'))) {
    if (show.ns !== NS.p || show.local !== 'custShow') continue;
    const raw = attr(show, 'id') ?? '';
    if (!/^\d{1,10}$/.test(raw) || Number(raw) > 4_294_967_295) continue;
    // A second show with the same id is not one a link can name: the first one wins.
    if (seen.has(Number(raw))) continue;
    seen.add(Number(raw));
    // A format character (a direction override, a zero-width space) can disguise a name.
    const name = cleanText(attr(show, 'name') ?? '')
      .replace(/\p{Cf}/gu, '')
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

/**
 * The file types a file link can name: documents, pictures, sound, and video. A program, a
 * script, a page that runs script (`.html`, `.svg`), a document with macros, a disk image, and
 * a name with no extension are never one of them.
 */
const DOCUMENT_EXTENSIONS = new Set([
  // Documents.
  'pdf',
  'txt',
  'rtf',
  'md',
  'csv',
  'tsv',
  'odt',
  'ods',
  'odp',
  'odg',
  'doc',
  'docx',
  'dotx',
  'xls',
  'xlsx',
  'xltx',
  'ppt',
  'pptx',
  'potx',
  'ppsx',
  'pages',
  'numbers',
  'key',
  'epub',
  'xps',
  'oxps',
  // Pictures.
  'png',
  'jpg',
  'jpeg',
  'gif',
  'bmp',
  'tif',
  'tiff',
  'webp',
  'heic',
  // Sound and video.
  'mp3',
  'wav',
  'm4a',
  'aac',
  'ogg',
  'flac',
  'mp4',
  'm4v',
  'mov',
  'avi',
  'wmv',
  'webm',
  'mkv',
]);
const MAX_UP_STEPS = 2;
/** One path segment as a person types it: letters, marks, digits, spaces, and punctuation. */
const SEGMENT = /^[\p{L}\p{M}\p{N} ._()~+,'&@!$=-]+$/u;
/** Names that a file system opens as a device, with any extension (`nul.txt`). */
const RESERVED = /^(?:con|prn|aux|nul|com[0-9¹²³]|lpt[0-9¹²³]|conin\$|conout\$)$/iu;

/** Whether a link draft is in the file form: `./name` or `../name`. */
export function isFileLink(url: string): boolean {
  return url.startsWith('./') || url.startsWith('../');
}

/**
 * The draft of the text in a file path field: a path with no `./` or `../` prefix gets `./`,
 * so `report.pdf` stays a file link, not a web address. `.` and `..` stay as they are while the
 * user types a prefix. @internal
 */
export function fileDraftUrl(text: string): string {
  if (!text || isFileLink(text) || text === '.' || text === '..') return text;
  return `./${text}`;
}

/** Whether one path segment is safe: no `.` or `..` step, no trailing dot or space, no device. */
function safeSegment(s: string): boolean {
  if (!SEGMENT.test(s) || s === '.' || s === '..' || s.trim() !== s) return false;
  // A file system drops a trailing dot when it opens the file: `a.exe.` opens `a.exe`.
  if (s.endsWith('.')) return false;
  return !RESERVED.test(s.split('.')[0].trim());
}

/**
 * The relationship target of a file link draft (`./docs/report.pdf` gives `docs/report.pdf`),
 * percent-encoded segment by segment, or a refusal with `errors.unsafeFileLink`. The name ends
 * in a document extension (`DOCUMENT_EXTENSIONS`).
 */
export function fileLinkTarget(url: string): string {
  const bad = (): never => {
    throw new CommandRefusal('errors.unsafeFileLink');
  };
  const path = url.normalize('NFC');
  if (path.length > 1024 || !isFileLink(path)) bad();
  const segments = path.split('/');
  let up = 0;
  while (segments[0] === '..') {
    segments.shift();
    up += 1;
  }
  if (up === 0 && segments[0] === '.') segments.shift();
  if (up > MAX_UP_STEPS || segments.length === 0) bad();
  for (const s of segments) if (!safeSegment(s)) bad();
  const name = segments[segments.length - 1];
  const dot = name.lastIndexOf('.');
  if (dot <= 0 || !DOCUMENT_EXTENSIONS.has(name.slice(dot + 1).toLowerCase())) bad();
  return [...Array<string>(up).fill('..'), ...segments.map(encodeURIComponent)].join('/');
}

/** Whether a decoded relative path is plain: no hidden character, no drive, no extra up step. */
function plainPath(text: string): boolean {
  if (/[\\:]|[\p{Cc}\p{Cf}]/u.test(text) || text.startsWith('/')) return false;
  const steps = text.split('/');
  let i = 0;
  while (i < MAX_UP_STEPS && steps[i] === '..') i += 1;
  return steps.slice(i).every((s) => s !== '' && s !== '.' && s !== '..');
}

/**
 * The file form of an external hyperlink target that is a relative path (`report.pdf` reads
 * `./report.pdf`), or undefined for any other target: a URL with a scheme, an absolute path, a
 * fragment, or a path with a backslash. A target whose decoded text hides a character, names a
 * drive, or steps up too far shows as it is written, still encoded.
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
  if (!plainPath(text)) text = target;
  return text.startsWith('../') ? text : `./${text}`;
}
