// The script fonts of a theme font collection (ECMA-376 Part 1, 20.1.4.1.16 `font`): the
// `a:font script="Jpan" typeface="…"` entries of `a:majorFont` and `a:minorFont`. A theme
// whose `a:ea` or `a:cs` typeface is empty names the face for each script there, and a run's
// language (`a:rPr/@lang`, `@altLang`, 21.1.2.3.9 `rPr`) picks the entry. The script codes are
// ISO 15924 codes.

/** The ISO 15924 script of the East Asian slot for a language tag, or undefined. */
export function eastAsianScriptOf(lang: string | undefined): string | undefined {
  if (!lang) return undefined;
  const tag = lang.toLowerCase();
  const primary = tag.split('-')[0];
  if (primary === 'ja') return 'Jpan';
  if (primary === 'ko') return 'Hang';
  if (primary === 'zh' || primary === 'yue' || primary === 'cmn') {
    // Traditional Chinese: an explicit `Hant` script, or a region that writes it.
    if (/-hant(-|$)/.test(tag) || /-(tw|hk|mo)(-|$)/.test(tag)) return 'Hant';
    return 'Hans';
  }
  return undefined;
}

/** The ISO 15924 script of the complex-script slot for each primary language subtag. */
const COMPLEX_SCRIPTS: ReadonlyMap<string, string> = new Map([
  ['ar', 'Arab'],
  ['fa', 'Arab'],
  ['ur', 'Arab'],
  ['ps', 'Arab'],
  ['sd', 'Arab'],
  ['ug', 'Uigh'],
  ['he', 'Hebr'],
  ['yi', 'Hebr'],
  ['th', 'Thai'],
  ['lo', 'Laoo'],
  ['km', 'Khmr'],
  ['my', 'Mymr'],
  ['hi', 'Deva'],
  ['mr', 'Deva'],
  ['ne', 'Deva'],
  ['sa', 'Deva'],
  ['kok', 'Deva'],
  ['bn', 'Beng'],
  ['as', 'Beng'],
  ['gu', 'Gujr'],
  ['pa', 'Guru'],
  ['ta', 'Taml'],
  ['te', 'Telu'],
  ['kn', 'Knda'],
  ['ml', 'Mlym'],
  ['or', 'Orya'],
  ['si', 'Sinh'],
  ['bo', 'Tibt'],
  ['dv', 'Thaa'],
  ['syr', 'Syrc'],
  ['am', 'Ethi'],
  ['ti', 'Ethi'],
]);

/** The ISO 15924 script of the complex-script slot for a language tag, or undefined. */
export function complexScriptOf(lang: string | undefined): string | undefined {
  if (!lang) return undefined;
  return COMPLEX_SCRIPTS.get(lang.toLowerCase().split('-')[0]);
}

/**
 * The language that picks a script font for a slot: the run's language when it is one of the
 * slot's languages, else its alternate language when that one is. Undefined when neither is.
 */
export function slotLanguage(
  slot: 'ea' | 'cs',
  lang: string | undefined,
  altLang: string | undefined
): string | undefined {
  const of = slot === 'ea' ? eastAsianScriptOf : complexScriptOf;
  if (of(lang)) return lang;
  if (of(altLang)) return altLang;
  return undefined;
}
