// The i18n CLI for packages/i18n, after docx-editor/scripts/validate-i18n.mjs.
//
//   node scripts/validate-i18n.mjs validate [--fix]  Check every locale against en.json.
//   node scripts/validate-i18n.mjs status            Show translation coverage per locale.
//   node scripts/validate-i18n.mjs new <lang>        Scaffold locales/<lang>.json (all null).
//
// en.json is the source of truth: every leaf is a non-empty string. A locale mirrors its shape;
// `null` means "fall back to English". A translated string keeps the English placeholders.
// `--fix` adds missing keys as null and removes unknown keys.
import { existsSync, readdirSync, readFileSync, writeFileSync } from 'node:fs';

const DIR = 'packages/i18n';
const LOCALES = `${DIR}/locales`;
const en = JSON.parse(readFileSync(`${DIR}/en.json`, 'utf8'));
const [command = 'validate', ...rest] = process.argv.slice(2);
/** BCP 47 file names: `de`, `pt-BR`, `zh-Hans`. */
const LOCALE_NAME = /^[a-z]{2,3}(?:-(?:[A-Z]{2}|[A-Z][a-z]{3}))?$/;

/** Every leaf of a catalog: `[dotted key, value]`. */
function leaves(obj, path = '', out = []) {
  for (const [k, v] of Object.entries(obj ?? {})) {
    if (v !== null && typeof v === 'object') leaves(v, `${path}${k}.`, out);
    else out.push([`${path}${k}`, v]);
  }
  return out;
}

const placeholders = (text) =>
  [...text.matchAll(/\{(\w+)\}/g)]
    .map((m) => m[1])
    .sort()
    .join(',');

/** Compare a locale with en.json; returns the locale with the English shape. */
function check(ref, loc, path, problems) {
  const out = {};
  for (const [k, v] of Object.entries(ref)) {
    const here = `${path}${k}`;
    if (typeof v === 'string') {
      if (!(k in loc)) problems.push(`missing ${here}`);
      const value = k in loc ? loc[k] : null;
      if (value !== null && typeof value !== 'string') problems.push(`not a string ${here}`);
      else if (typeof value === 'string' && placeholders(value) !== placeholders(v))
        problems.push(`placeholders differ ${here}`);
      out[k] = value;
    } else {
      if (!(k in loc) || typeof loc[k] !== 'object' || loc[k] === null)
        problems.push(`missing ${here}`);
      out[k] = check(v, loc[k] ?? {}, `${here}.`, problems);
    }
  }
  for (const k of Object.keys(loc)) if (!(k in ref)) problems.push(`unknown ${path}${k}`);
  return out;
}

function skeleton(obj) {
  return Object.fromEntries(
    Object.entries(obj).map(([k, v]) => [k, typeof v === 'string' ? null : skeleton(v)])
  );
}

const localeFiles = () =>
  existsSync(LOCALES)
    ? readdirSync(LOCALES)
        .filter((f) => f.endsWith('.json'))
        .sort()
    : [];

function validate(fix) {
  let failed = false;
  for (const [key, value] of leaves(en)) {
    if (typeof value !== 'string' || value.trim() === '') {
      console.error(`en.json: ${key} must be a non-empty string`);
      failed = true;
    }
  }
  for (const file of localeFiles()) {
    const path = `${LOCALES}/${file}`;
    const loc = JSON.parse(readFileSync(path, 'utf8'));
    const problems = [];
    const fixed = check(en, loc, '', problems);
    if (!problems.length) continue;
    if (fix) writeFileSync(path, `${JSON.stringify(fixed, null, 2)}\n`);
    else {
      console.error(`${file}:\n  ${problems.join('\n  ')}`);
      failed = true;
    }
  }
  console.log(failed ? 'i18n: problems found' : 'i18n: catalogs valid');
  return failed ? 1 : 0;
}

function status() {
  const total = leaves(en).length;
  console.log(`en: ${total} keys (source of truth)`);
  for (const file of localeFiles()) {
    const loc = JSON.parse(readFileSync(`${LOCALES}/${file}`, 'utf8'));
    const done = leaves(loc).filter(([, v]) => typeof v === 'string').length;
    console.log(
      `${file.replace(/\.json$/, '')}: ${done}/${total} (${Math.round((done / total) * 100)}%)`
    );
  }
  if (localeFiles().length === 0) console.log('No other locales. Add one with i18n:new <lang>.');
  return 0;
}

function scaffold(lang) {
  if (!lang || !LOCALE_NAME.test(lang)) {
    console.error(
      'Usage: bun run i18n:new <lang>, where <lang> is a BCP 47 tag such as de or pt-BR.'
    );
    return 1;
  }
  const path = `${LOCALES}/${lang}.json`;
  if (existsSync(path)) {
    console.error(`${path} exists.`);
    return 1;
  }
  writeFileSync(path, `${JSON.stringify(skeleton(en), null, 2)}\n`);
  console.log(
    `Created ${path}. Replace null values with translations; null falls back to English.`
  );
  return 0;
}

const code =
  command === 'validate'
    ? validate(rest.includes('--fix'))
    : command === 'status'
      ? status()
      : command === 'new'
        ? scaffold(rest[0])
        : (console.error(`Unknown command ${command}. Use validate, status, or new <lang>.`), 1);
process.exit(code);
