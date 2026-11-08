// The font family picker's groups, shared by the React and Vue adapters: the theme fonts of the
// current slide (labeled Headings and Body), the families the deck uses, then every family the
// editor offers. After DOCX Editor's font picker
// (`packages/react/src/editor/toolbar/FontFamily.tsx`), which lists the document's fonts in their
// own group above the full list and draws each name in its own face. DOCX Editor commit adf0e178.

import type { TranslationKey } from '@pptx-editor.dev/i18n';
import { xmlPartOf, slideRefAt } from '../store/index.ts';
import { parseTheme } from '../cascade/index.ts';
import { namedFamilies, type DocumentSession } from './document-session.ts';
import { FONT_NAME, LATIN_THEME_FONTS } from '../automation/commands/font-format.ts';
import { filterFontFamilies } from './toolbar-controls.ts';

/**
 * A font picker group: the slide's theme fonts (`theme`), the other families the deck names
 * (`document`), or every family the editor offers (`all`). @public
 */
export type FontPickerGroupId = 'theme' | 'document' | 'all';

/**
 * One row of the font picker, as `useFontFamily().groups` gives it. Pass `value` to
 * `useFontFamily().setValue`. @public
 */
export interface FontPickerRow {
  /** The family name the row shows. */
  readonly family: string;
  /**
   * What the row applies: the theme font reference (`+mj-lt`, `+mn-lt`) on a theme row, so the
   * text follows a later theme change (ECMA-376 Part 1, 20.1.4.1.17 `latin`), else the family.
   */
  readonly value: string;
  /** The theme font of a theme row, or null on a row that names a literal family. */
  readonly theme: 'major' | 'minor' | null;
  /**
   * The CSS value that draws the name in its own face (a quoted family), or null for a name
   * that is not a valid family. A face that is not loaded falls back to the chrome font.
   */
  readonly face: string | null;
  /** The i18n key of the theme role (`Headings`, `Body`), on theme font rows only. */
  readonly roleKey: TranslationKey | null;
  /**
   * True on a theme row whose theme font is empty in the slide's theme: the row shows its role
   * with no family and applies nothing. Only a host's `Item` gives such a row.
   */
  readonly disabled: boolean;
}

/**
 * One group of the font picker, with its i18n heading key (`toolbar.fontGroups.theme`,
 * `toolbar.fontGroups.document`, `toolbar.fontGroups.all`). @public
 */
export interface FontPickerGroup {
  /** Which group this is: the theme fonts, the fonts the deck uses, or the other fonts. */
  readonly id: FontPickerGroupId;
  /** The i18n key of the group heading; show it with `t(labelKey)`. */
  readonly labelKey: TranslationKey;
  /** The rows of the group, in the order the picker shows them. */
  readonly rows: readonly FontPickerRow[];
}

/** What the picker reads from the editor. @internal */
export interface FontPickerSource {
  readonly session: DocumentSession | null;
  fontFamilies(): readonly string[];
}

const GROUP_LABELS: Record<FontPickerGroupId, TranslationKey> = {
  theme: 'toolbar.fontGroups.theme',
  document: 'toolbar.fontGroups.document',
  all: 'toolbar.fontGroups.all',
};

/**
 * The CSS value that draws a family name in its own face. The name passes the font catalog
 * rule (letters, digits, marks, space, `-`, `.`, `+`, `_`), so it holds no quote or backslash.
 * @internal
 */
export function fontFaceValue(family: string): string | null {
  return FONT_NAME.test(family) ? `"${family}"` : null;
}

/**
 * The Latin heading and body fonts of a slide's theme (ECMA-376 Part 1, 20.1.4.1.24
 * `majorFont`, 20.1.4.1.25 `minorFont`). @internal
 */
export function slideThemeFonts(
  session: DocumentSession | null,
  slide: number
): { readonly headings: string | null; readonly body: string | null } {
  const themePart = session && slideRefAt(session.store.package, slide)?.themePart;
  const part = session && themePart ? xmlPartOf(session.store.package, themePart) : undefined;
  if (!part) return { headings: null, body: null };
  const theme = parseTheme(part);
  const valid = (f: string | undefined): string | null => (f && FONT_NAME.test(f) ? f : null);
  return { headings: valid(theme.majorFont.latin), body: valid(theme.minorFont.latin) };
}

const row = (family: string): FontPickerRow => ({
  family,
  value: family,
  theme: null,
  face: fontFaceValue(family),
  roleKey: null,
  disabled: false,
});

/** A theme font row; an empty `family` gives the disabled row of a theme font the theme leaves empty. */
const themeRow = (family: string, theme: 'major' | 'minor'): FontPickerRow => ({
  family,
  value: LATIN_THEME_FONTS[theme],
  theme,
  face: family ? fontFaceValue(family) : null,
  roleKey: theme === 'major' ? 'toolbar.fontHeadings' : 'toolbar.fontBody',
  disabled: !family,
});

/**
 * Whether a row is the font the text uses, for the check mark. A theme font marks only its
 * theme row (`+mj-lt` marks Headings); a literal family marks only the rows that name it, never
 * a theme row with the same name. `family` and `theme` are the selection's `fontFamily` and
 * `fontFamilyTheme`. @internal
 */
export function isCurrentFontRow(
  row: FontPickerRow,
  family: string | null | undefined,
  theme: 'major' | 'minor' | null | undefined
): boolean {
  if (theme) return row.theme === theme;
  return row.theme === null && !!family && row.family.toLowerCase() === family.toLowerCase();
}

/**
 * The picker's groups for a search `query`, empty groups left out: the theme fonts of `slide`,
 * the other families the deck names, then every family the editor offers. @internal
 */
export function fontPickerGroups(
  source: FontPickerSource | null,
  slide: number,
  query: string
): readonly FontPickerGroup[] {
  const session = source?.session ?? null;
  const theme = slideThemeFonts(session, slide);
  const themeRows: FontPickerRow[] = [];
  if (theme.headings) themeRows.push(themeRow(theme.headings, 'major'));
  if (theme.body) themeRows.push(themeRow(theme.body, 'minor'));
  const inTheme = new Set(themeRows.map((r) => r.family.toLowerCase()));
  const used = session ? namedFamilies(session.store.package) : [];
  const documentRows = used
    .filter((f) => FONT_NAME.test(f) && !inTheme.has(f.toLowerCase()))
    .map((f) => row(f));
  const q = query.trim().toLowerCase();
  const match = (r: FontPickerRow): boolean => !q || r.family.toLowerCase().includes(q);
  const groups: FontPickerGroup[] = [
    { id: 'theme', labelKey: GROUP_LABELS.theme, rows: themeRows.filter(match) },
    { id: 'document', labelKey: GROUP_LABELS.document, rows: documentRows.filter(match) },
    {
      id: 'all',
      labelKey: GROUP_LABELS.all,
      rows: filterFontFamilies(source?.fontFamilies() ?? [], query).map((f) => row(f)),
    },
  ];
  return groups.filter((g) => g.rows.length > 0);
}

/**
 * The row for one value a host passes to `FontFamily.Item`: a theme font reference (`+mj-lt`,
 * `+mn-lt`) gives the theme row of `slide`, with its family name and role, or a disabled row
 * with its role only when the slide's theme leaves that font empty; any other valid family
 * gives a plain row. Null for a name that is not a valid family. @internal
 */
export function fontPickerRow(
  source: FontPickerSource | null,
  slide: number,
  value: string
): FontPickerRow | null {
  const theme =
    value === LATIN_THEME_FONTS.major
      ? 'major'
      : value === LATIN_THEME_FONTS.minor
        ? 'minor'
        : null;
  if (!theme) return FONT_NAME.test(value) ? row(value) : null;
  const fonts = slideThemeFonts(source?.session ?? null, slide);
  return themeRow((theme === 'major' ? fonts.headings : fonts.body) ?? '', theme);
}

/**
 * What Enter in the search field applies: the value of the first row shown (a theme font
 * reference on a theme row). @internal
 */
export function firstFontPickerFamily(groups: readonly FontPickerGroup[]): string | null {
  return groups[0]?.rows[0]?.value ?? null;
}

/**
 * Focus the current option of a font list that has no search field (a list a host composed
 * from `Item` parts), else its first enabled option, so the arrow keys reach the options. Call
 * it once, after the list opens and its popover took focus. A list with a search field keeps
 * focus in the field. @internal
 */
export function focusFontOption(list: HTMLElement | null): void {
  if (!list || list.closest('.pptx-font-picker')?.querySelector('input')) return;
  const options = [...list.querySelectorAll<HTMLElement>('[role="option"]:not(:disabled)')];
  const current = options.find((o) => o.getAttribute('aria-selected') === 'true');
  (current ?? options[0])?.focus();
}

/**
 * Scroll the font list so the first row of the current family is in view, centered when it was
 * out of view. Call it once, when the list opens. @internal
 */
export function revealCurrentFont(list: HTMLElement | null): void {
  const current = list?.querySelector<HTMLElement>('[role="option"][aria-selected="true"]');
  if (!list || !current) return;
  const box = list.getBoundingClientRect();
  const at = current.getBoundingClientRect();
  if (at.top >= box.top && at.bottom <= box.bottom) return;
  list.scrollTop += at.top - box.top - (list.clientHeight - at.height) / 2;
}
