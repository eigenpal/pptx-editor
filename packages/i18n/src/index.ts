/**
 * `@pptx-editor.dev/i18n` — the chrome's user-facing strings.
 *
 * `en.json` is the source of truth. A locale catalog mirrors its shape; a `null` leaf falls
 * back to English. Keys are dotted paths (`toolbar.bold`). Parameters use `{name}`.
 *
 * Mirrors docx-editor/packages/i18n: one catalog, typed keys, no hardcoded English in chrome.
 *
 * @packageDocumentation
 * @public
 */

import en from '../en.json' with { type: 'json' };

/** The English catalog, the shape every locale mirrors. @public */
export type Catalog = typeof en;

type Leaves<T, Prefix extends string = ''> = {
  [K in keyof T & string]: T[K] extends string
    ? `${Prefix}${K}`
    : T[K] extends Record<string, unknown>
      ? Leaves<T[K], `${Prefix}${K}.`>
      : never;
}[keyof T & string];

/** Every translation key, derived from `en.json`. @public */
export type TranslationKey = Leaves<Catalog>;

/** Values substituted into `{name}` placeholders. @public */
export type TranslationParams = Readonly<Record<string, string | number>>;

// An interface, not a type alias: a declaration that infers a prop of this type (a Vue
// component's props) then prints `LocaleCatalog`, not the whole expanded catalog.
/** A locale catalog: the English shape, where `null` means "fall back to English". @public */
export interface LocaleCatalog extends DeepNullable<Catalog> {}

type DeepNullable<T> = {
  [K in keyof T]?: T[K] extends string ? string | null : DeepNullable<T[K]>;
};

/** A function that turns a key into display text. @public */
export type Translate = (key: TranslationKey, params?: TranslationParams) => string;

/** The English catalog. @public */
export const englishCatalog: Catalog = en;

function lookup(catalog: unknown, key: string): string | null {
  let node: unknown = catalog;
  for (const part of key.split('.')) {
    if (node === null || typeof node !== 'object') return null;
    if (!Object.prototype.hasOwnProperty.call(node, part)) return null;
    node = (node as Record<string, unknown>)[part];
  }
  return typeof node === 'string' ? node : null;
}

function interpolate(text: string, params: TranslationParams | undefined): string {
  if (!params) return text;
  return text.replace(/\{(\w+)\}/g, (match, name: string) =>
    Object.prototype.hasOwnProperty.call(params, name) ? String(params[name]) : match
  );
}

/**
 * Whether a string is a key of the English catalog. Engine refusal reasons and other keys that
 * arrive as plain strings pass through this guard before they reach a `Translate`, so no caller
 * casts a string to `TranslationKey`.
 *
 * @public
 */
export function isTranslationKey(key: string): key is TranslationKey {
  return lookup(en, key) !== null;
}

/**
 * Create a translator over an optional locale catalog. Missing or `null` entries fall back to
 * English; an unknown key returns the key itself, so a gap is visible instead of blank.
 *
 * @public
 */
export function createTranslator(locale?: LocaleCatalog): Translate {
  return (key, params) => {
    const text = (locale ? lookup(locale, key) : null) ?? lookup(en, key) ?? key;
    return interpolate(text, params);
  };
}
