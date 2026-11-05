// The Bullets and Numbering galleries and dialog: the bullet format of the selected paragraphs
// (Office.js `ParagraphFormat.bulletFormat`).
//
// ECMA-376 Part 1, DrawingML `CT_TextParagraphProperties` (`a:pPr`) bullet children, in schema
// order: `a:buClr` or `a:buClrTx`, `a:buSzPct` (`ST_TextBulletSizePercent`, 25 % to 400 %;
// written in thousandths, `150000`) or `a:buSzTx`, `a:buFont` or `a:buFontTx`, and
// `a:buChar` or `a:buAutoNum` (`ST_TextAutonumberScheme`, `startAt` 1 to 32767). A paragraph
// with no list turns its list on first (`setListKind`: margins and inheritance as
// `list-commands.ts` writes them); one that shows a list keeps its margins. Then each stated
// field changes the child in place, so two peers who restyle one paragraph change one
// element. A bullet in a symbol font carries `charset="2"` (symbol charset), so a reader takes
// its character as a symbol.

import { isValidXmlText } from '@eigenpal/ooxml-core';
import { toHex, type ResolvedParagraph } from '../cascade/index.ts';
import type { BulletFormat, BulletFormatRead } from '../contracts/index.ts';
import type { BulletGalleryKey, NumberGalleryKey } from './toolbar-values.ts';
import { isSymbolFont } from '../layout/index.ts';
import { attr, type Element, type NodeSpec } from '../store/index.ts';
import { FONT_NAME } from '../automation/commands/font-format.ts';
import {
  resolvedParagraphs,
  setListKind,
  type ListTarget,
} from '../automation/commands/list-commands.ts';
import { CommandRefusal, type OpBatch } from '../automation/commands/op-batch.ts';
import { specFromXml } from '../automation/commands/shape-commands.ts';
import { colorXml, isColorChoice } from '../automation/commands/style-commands.ts';
import { aSpec, ensureParagraphProps, paragraphsOf } from '../automation/commands/text-model.ts';

/** One gallery item: its key, and the bullet format it applies (null turns the list off). @public */
export interface ListGalleryItem<K extends string = string> {
  /** A stable key, never renamed: `none`, a bullet name (`filledRound`), or a scheme. */
  readonly key: K;
  /** The bullet format the item applies, the command's `format`; null turns the list off. */
  readonly format: BulletFormat | null;
}

/** The Bullets gallery, in its order (Home > Bullets). @public */
export const BULLET_GALLERY: readonly ListGalleryItem<BulletGalleryKey>[] = [
  { key: 'none', format: null },
  { key: 'filledRound', format: { kind: 'bullet', char: '•', font: 'Arial' } },
  { key: 'hollowRound', format: { kind: 'bullet', char: 'o', font: 'Courier New' } },
  { key: 'filledSquare', format: { kind: 'bullet', char: '§', font: 'Wingdings' } },
  { key: 'hollowSquare', format: { kind: 'bullet', char: 'q', font: 'Wingdings' } },
  { key: 'star', format: { kind: 'bullet', char: 'v', font: 'Wingdings' } },
  { key: 'arrow', format: { kind: 'bullet', char: 'Ø', font: 'Wingdings' } },
  { key: 'checkmark', format: { kind: 'bullet', char: 'ü', font: 'Wingdings' } },
];

/** The Numbering gallery, in its order (Home > Numbering). @public */
export const NUMBER_GALLERY: readonly ListGalleryItem<NumberGalleryKey>[] = [
  { key: 'none', format: null },
  ...(
    [
      'arabicPeriod',
      'arabicParenR',
      'romanUcPeriod',
      'alphaUcPeriod',
      'alphaLcParenR',
      'alphaLcPeriod',
      'romanLcPeriod',
    ] as const
  ).map((scheme) => ({ key: scheme, format: { kind: 'number', scheme } as const })),
];

/** The attributes written with each gallery font: `panose`, `pitchFamily`, and `charset`. */
const FONT_ATTRS: Readonly<Record<string, Readonly<Record<string, string>>>> = Object.assign(
  Object.create(null) as Record<string, Readonly<Record<string, string>>>,
  {
    Arial: { panose: '020B0604020202020204', pitchFamily: '34', charset: '0' },
    'Courier New': { panose: '02070309020205020404', pitchFamily: '49', charset: '0' },
    Wingdings: { panose: '05000000000000000000', pitchFamily: '2', charset: '2' },
  }
);
/** A bullet font: a family name (the `setFontFamily` bound) or a theme font reference. */
const THEME_FONT = /^\+m[jn]-(lt|ea|cs)$/;
/** `ST_TextAutonumberScheme` values the editor writes: `Plain` exists for Arabic numbers only. */
const WRITABLE_SCHEME =
  /^((alphaLc|alphaUc|romanLc|romanUc)(ParenBoth|ParenR|Period)|arabic(ParenBoth|ParenR|Period|Plain))$/;
/** `CT_TextParagraphProperties` children in schema order; a group shares one rank. */
const ORDER: readonly (readonly string[])[] = [
  ['lnSpc'],
  ['spcBef'],
  ['spcAft'],
  ['buClrTx', 'buClr'],
  ['buSzTx', 'buSzPct', 'buSzPts'],
  ['buFontTx', 'buFont'],
  ['buNone', 'buAutoNum', 'buChar', 'buBlip'],
  ['tabLst'],
  ['defRPr'],
  ['extLst'],
];
const rankOf = (local: string): number => ORDER.findIndex((group) => group.includes(local));
/** The first gallery bullet and scheme: what a paragraph without one takes. */
const FIRST_BULLET = { char: '•', font: 'Arial' } as const;
const FIRST_SCHEME = 'arabicPeriod';

/**
 * Whether a bullet character is one the command writes: one code point, not a control
 * character, and valid in XML (no lone surrogate). @internal
 */
export function isBulletChar(char: unknown): boolean {
  return (
    typeof char === 'string' &&
    [...char].length === 1 &&
    !/\p{Cc}/u.test(char) &&
    isValidXmlText(char)
  );
}

/** The fields each kind takes, in a null-prototype record so a file or host value never reads `toString`. */
const FIELDS: Readonly<Record<string, readonly string[]>> = Object.assign(Object.create(null), {
  bullet: ['kind', 'char', 'font', 'sizePercent', 'color'],
  number: ['kind', 'scheme', 'startAt', 'sizePercent', 'color'],
});
/** A format with no kind takes the size and the color only. */
const KINDLESS_FIELDS: readonly string[] = ['kind', 'sizePercent', 'color'];

/**
 * Check a bullet format, or refuse it. The detail names the field: one the kind does not take
 * (a scheme on a bullet), or `kind` for a format with no kind and nothing to change.
 */
export function checkBulletFormat(style: BulletFormat | null): void {
  if (style === null) return;
  const bad = (name: string): never => {
    throw new CommandRefusal('errors.unsupported', name);
  };
  if (typeof style !== 'object') bad('kind');
  const allowed =
    style.kind === undefined
      ? KINDLESS_FIELDS
      : typeof style.kind === 'string'
        ? FIELDS[style.kind]
        : undefined;
  if (!allowed) bad('kind');
  for (const key of Object.keys(style))
    if (!allowed!.includes(key) && (style as Record<string, unknown>)[key] !== undefined) bad(key);
  if (style.kind === undefined && style.sizePercent === undefined && style.color === undefined)
    bad('kind');
  if (style.kind === 'bullet') {
    if (style.char !== undefined && !isBulletChar(style.char)) bad('char');
    if (
      style.font !== undefined &&
      !(
        typeof style.font === 'string' &&
        style.font.trim() === style.font &&
        (FONT_NAME.test(style.font) || THEME_FONT.test(style.font))
      )
    )
      bad('font');
  } else if (style.kind === 'number') {
    if (
      style.scheme !== undefined &&
      !(typeof style.scheme === 'string' && WRITABLE_SCHEME.test(style.scheme))
    )
      bad('scheme');
    if (
      style.startAt !== undefined &&
      !(Number.isInteger(style.startAt) && style.startAt >= 1 && style.startAt <= 32_767)
    )
      bad('startAt');
  }
  const { sizePercent: size, color } = style;
  if (size != null && !(typeof size === 'number' && size >= 25 && size <= 400)) bad('sizePercent');
  if (color != null && !isColorChoice(color)) bad('color');
}

/** The bullet children of `a:pPr` with these local names. */
const kids = (pPr: Element, names: readonly string[]): Element[] =>
  pPr.children.filter((c): c is Element => c.type === 'element' && names.includes(c.local));

/**
 * Make the `a:pPr` hold exactly one child of a schema slot: change the one that is there when
 * it has the same name (an attribute write, which two peers merge), otherwise replace the slot.
 * Null empties the slot. `replace` always writes a new child: a bullet font that changes with
 * the list kind goes with its `a:buChar` or `a:buAutoNum`, so two peers who change the kind at
 * once each keep a whole bullet, and the read shows one peer's font and kind together.
 */
function setSlot(
  batch: OpBatch,
  part: string,
  pPrId: string,
  slot: readonly string[],
  spec: NodeSpec | null,
  replace = false
): void {
  const present = kids(batch.element(part, pPrId), slot);
  const same = spec && !replace ? present.find((c) => c.local === spec.local) : undefined;
  for (const c of present) if (c !== same) batch.push({ op: 'removeNode', part, node: c.id });
  if (!spec) return;
  if (same) return morph(batch, part, same, spec);
  const rank = rankOf(spec.local);
  let index = 0;
  batch.element(part, pPrId).children.forEach((c, k) => {
    if (c.type === 'element' && rankOf(c.local) !== -1 && rankOf(c.local) <= rank) index = k + 1;
  });
  batch.insert({ op: 'insertNode', part, parent: pPrId, index, spec });
}

/**
 * Turn an element into `spec`, which has the same name, with as few ops as it takes: attribute
 * writes, and the same for children whose names match one by one (`a:buClr` keeps its
 * `a:srgbClr` and changes `val`). Two peers' attribute writes merge to one value; a child that
 * changes its name is replaced.
 */
function morph(batch: OpBatch, part: string, el: Element, spec: NodeSpec): void {
  const wanted = new Map((spec.attrs ?? []).filter((x) => !x.ns).map((x) => [x.local, x.value]));
  const set = (name: string, value: string | null) =>
    batch.push({ op: 'setAttribute', part, node: el.id, name: { ns: '', local: name }, value });
  for (const x of el.attrs) if (!x.ns && !wanted.has(x.local)) set(x.local, null);
  for (const [name, value] of wanted)
    if (el.attrs.find((x) => !x.ns && x.local === name)?.value !== value) set(name, value);
  const have = el.children.filter((c): c is Element => c.type === 'element');
  const want = (spec.children ?? []).filter((c): c is NodeSpec => 'local' in c);
  if (have.length === want.length && have.every((c, k) => c.local === want[k]!.local)) {
    have.forEach((c, k) => morph(batch, part, c, want[k]!));
    return;
  }
  for (const c of have) batch.push({ op: 'removeNode', part, node: c.id });
  for (const [k, c] of want.entries())
    batch.insert({ op: 'insertNode', part, parent: el.id, index: k, spec: c });
}

/** The `a:buFont` a bullet font writes, with the known attributes for the gallery fonts. */
function fontSpec(font: string): NodeSpec {
  const known = FONT_ATTRS[font];
  const attrs = known ?? (isSymbolFont(font) ? { pitchFamily: '2', charset: '2' } : {});
  return aSpec('buFont', { typeface: font, ...attrs });
}

/**
 * Apply a bullet format to the target paragraphs. `resolve` is false for a probe, which only
 * needs ops (`setListKind`).
 */
export function setBulletFormat(
  batch: OpBatch,
  part: string,
  slide: number,
  target: ListTarget,
  style: BulletFormat | null,
  resolve = true
): boolean {
  checkBulletFormat(style);
  // No kind: only the size and color, of the paragraphs that show a list. Which ones do is
  // read also for a probe, so `can` and `exec` agree.
  const kind = style ? style.kind : 'none';
  const read = resolve || !kind;
  const before = read ? resolvedParagraphs(batch, slide, target) : [];
  // A paragraph that already shows a list keeps its margins: only its bullet changes. One
  // without a list, and None, go through Home > Bullets (`setListKind`).
  const fresh =
    style && read ? target.paragraphs.filter((i) => !before[i]?.bullet) : target.paragraphs;
  if (kind && fresh.length)
    setListKind(batch, part, slide, { ...target, paragraphs: fresh }, kind, resolve);
  if (!style) return true;
  if (!kind && fresh.length === target.paragraphs.length) return false;
  // What a paragraph with no list shows once it has one: the scheme its placeholder inherits.
  const after =
    resolve && style.kind === 'number' && fresh.length
      ? resolvedParagraphs(batch, slide, target)
      : before;
  const all = paragraphsOf(batch.element(part, target.txBodyId));
  for (const i of kind ? target.paragraphs : target.paragraphs.filter((j) => !fresh.includes(j))) {
    const pPrId = ensureParagraphProps(batch, part, all[i].id);
    const shown = before[i]?.bullet;
    const sameKind = !!shown && (shown.type === 'char') === (style.kind === 'bullet');
    const set = (slot: readonly string[], spec: NodeSpec | null, replace = false) =>
      setSlot(batch, part, pPrId, slot, spec, replace);
    if (style.color !== undefined)
      set(
        ['buClrTx', 'buClr'],
        style.color === null
          ? aSpec('buClrTx')
          : aSpec('buClr', {}, [specFromXml(colorXml(style.color))]),
        // A color and its luminance are one pick: a new `a:buClr`, never a merge of two.
        true
      );
    if (style.sizePercent !== undefined)
      set(
        ['buSzTx', 'buSzPct', 'buSzPts'],
        style.sizePercent === null
          ? aSpec('buSzTx')
          : aSpec('buSzPct', { val: String(Math.round(style.sizePercent * 1000)) })
      );
    if (style.kind === 'bullet') {
      // An absent character keeps the paragraph's own bullet, or takes the first gallery one.
      const char = style.char ?? (sameKind || fresh.includes(i) ? undefined : FIRST_BULLET.char);
      const font = style.char === undefined && style.font === undefined ? undefined : style.font;
      if (char !== undefined) {
        // A character with no font follows the text font (`a:buFontTx`), also over a list
        // style's symbol font.
        const fontEl = font
          ? fontSpec(font)
          : style.char
            ? aSpec('buFontTx')
            : fontSpec(FIRST_BULLET.font);
        // A new character in a new font is one bullet: both are written as new children, so a
        // peer's concurrent new bullet cannot keep its character with this font. A change of
        // the character alone, or of the font alone, is an attribute write, and two peers who
        // change one each merge field by field.
        const pPr = batch.element(part, pPrId);
        const nowChar = attr(kids(pPr, ['buChar'])[0], 'char');
        const nowFont = kids(pPr, ['buFontTx', 'buFont'])[0];
        const pair =
          nowChar !== char &&
          (nowFont?.local !== fontEl.local ||
            attr(nowFont, 'typeface') !== fontEl.attrs?.find((x) => x.local === 'typeface')?.value);
        set(['buNone', 'buAutoNum', 'buChar', 'buBlip'], aSpec('buChar', { char }), pair);
        set(['buFontTx', 'buFont'], fontEl, !sameKind || pair);
      } else if (font !== undefined) set(['buFontTx', 'buFont'], fontSpec(font));
    } else if (style.kind === 'number') {
      const now = after[i]?.bullet;
      const own = now?.type === 'autonum' ? now : undefined;
      // An absent start keeps the paragraph's own; a paragraph with no numbering starts at 1.
      const startAt = style.startAt ?? own?.startAt ?? 1;
      const ownNum = kids(batch.element(part, pPrId), ['buAutoNum'])[0];
      // The asked scheme; the shown one when only the start of an inherited scheme changes;
      // the first gallery scheme for a bulleted paragraph.
      const scheme =
        style.scheme ??
        (own
          ? style.startAt !== undefined && !ownNum
            ? own.scheme
            : undefined
          : fresh.includes(i)
            ? undefined
            : FIRST_SCHEME);
      if (scheme !== undefined) {
        if (style.scheme !== undefined || !own)
          set(['buFontTx', 'buFont'], aSpec('buFont', { typeface: '+mj-lt' }), !sameKind);
        set(
          ['buNone', 'buAutoNum', 'buChar', 'buBlip'],
          aSpec('buAutoNum', { type: scheme, ...(startAt > 1 ? { startAt: String(startAt) } : {}) })
        );
      } else if (style.startAt !== undefined && ownNum) {
        batch.push({
          op: 'setAttribute',
          part,
          node: ownNum.id,
          name: { ns: '', local: 'startAt' },
          value: style.startAt > 1 ? String(style.startAt) : null,
        });
      }
    }
  }
  return true;
}

/** The bullet format of a resolved paragraph, as the gallery and dialog show it. */
export function bulletFormatOf(p: ResolvedParagraph): BulletFormatRead | null {
  const b = p.bullet;
  if (!b) return null;
  // A point size (`a:buSzPts`) reads as a percent of the first text's size, as layout sizes it.
  const text = (p.runs.find((r) => r.kind !== 'break') ?? { props: p.endProps }).props.size;
  const sizePercent = Math.round(
    'pct' in b.size ? b.size.pct * 100 : text > 0 ? (b.size.pts / text) * 100 : 100
  );
  const color = b.color ? { color: toHex(b.color).toUpperCase() } : {};
  if (b.type === 'char')
    return {
      kind: 'bullet',
      char: b.char,
      ...(b.font ? { font: b.font } : {}),
      sizePercent,
      ...color,
    };
  return { kind: 'number', scheme: b.scheme, startAt: b.startAt, sizePercent, ...color };
}

/**
 * The gallery key of a bullet format: `none` for no list, undefined when the gallery has no such
 * item. @public
 */
export function galleryKeyOf(
  gallery: readonly ListGalleryItem[],
  state: BulletFormatRead | null | undefined
): string | undefined {
  if (state === undefined) return undefined;
  if (state === null) return 'none';
  return gallery.find((item) => {
    const s = item.format;
    if (!s || s.kind !== state.kind) return false;
    return s.kind === 'bullet'
      ? state.kind === 'bullet' && s.char === state.char && (s.font ?? '') === (state.font ?? '')
      : state.kind === 'number' && s.scheme === state.scheme;
  })?.key;
}

/** The bullet format of a gallery key, or undefined for a key the gallery does not have. @public */
export function galleryFormat(
  gallery: readonly ListGalleryItem[],
  key: unknown
): BulletFormat | null | undefined {
  return gallery.find((item) => item.key === key)?.format;
}

/** Whether two reads show the same thing, for adapters that keep the last read. @internal */
export function sameBulletFormat(
  a: BulletFormatRead | null | undefined,
  b: BulletFormatRead | null | undefined
): boolean {
  return a === b || JSON.stringify(a) === JSON.stringify(b);
}
