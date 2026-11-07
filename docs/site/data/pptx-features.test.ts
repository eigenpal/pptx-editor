import { describe, expect, test } from 'bun:test';
import { existsSync, readFileSync, readdirSync } from 'node:fs';
import { join } from 'node:path';
import { resolveSlide } from '../../../packages/core/src/cascade/index.ts';
import { APPROXIMATIONS } from '../../../packages/core/src/layout/approximations.ts';
import { createFixedMeasurer, layoutSlide } from '../../../packages/core/src/layout/index.ts';
import { readPptx } from '../../../packages/core/src/store/index.ts';
import { readManifest, referenceState } from '../../../scripts/fidelity/reference-identity.ts';
import { FIXTURE_KNOWN_GAPS } from '../../../scripts/fidelity/thresholds.ts';
import { FEATURE_EVIDENCE, type FeatureEvidence } from './pptx-feature-evidence.ts';
import {
  FEATURE_CATEGORY_LABELS,
  FEATURE_TIERS,
  pptxFeatureById,
  pptxFeatures,
  type FeatureStatus,
  type PptxFeature,
} from './pptx-features.ts';
import { FEATURE_EDITS } from '../../../scripts/fidelity/feature-edits.ts';
import { SAVE_VERIFIED_FEATURE_EDITS, SAVE_VERIFIED_FIXTURES } from './save-verified-fixtures.ts';

const FIXTURES = join(import.meta.dir, '../../../fixtures');

/** The fields of a DOCX Editor matrix row (`docs/site/data/word-features.ts`). */
const DOCX_ROW_FIELDS = new Set([
  'id',
  'name',
  'category',
  'editing',
  'rendering',
  'roundTrip',
  'tier',
  'notes',
  'docsLink',
]);

/**
 * The page-wide text approximations: the line model, the shaper's glyph advances, and missing
 * glyphs. They apply to all text, so they do not keep a row from full rendering when its fixtures
 * pass the harness. Any other approximation is a rendering limit of its row. A reviewer also
 * checks that the notes of a full row name no rendering gap; no test can read that.
 */
const PAGE_WIDE_APPROXIMATIONS = new Set(['line-model', 'glyph-advances', 'missing-glyphs']);

const NO_EVIDENCE: FeatureEvidence = { fixtures: [] };

/** The round-trip values that claim the content survives a save. */
const CLAIMS_ROUND_TRIP = new Set<FeatureStatus>(['full', 'preserved']);

/** The evidence entry of a row. A missing entry fails its own test below. */
const evidence = (row: PptxFeature): FeatureEvidence =>
  Object.hasOwn(FEATURE_EVIDENCE, row.id) ? FEATURE_EVIDENCE[row.id]! : NO_EVIDENCE;

/** Values that occur more than once, sorted. */
const repeats = (values: readonly string[]): string[] =>
  [...new Set(values.filter((value, i) => values.indexOf(value) !== i))].sort();

describe('pptx-features: the matrix itself', () => {
  test('every row has a unique id', () => {
    expect(repeats(pptxFeatures.map((row) => row.id))).toEqual([]);
  });

  test('every id starts with its category', () => {
    const wrong = pptxFeatures
      .filter((row) => !row.id.startsWith(`${row.category}.`))
      .map((row) => row.id);
    expect(wrong).toEqual([]);
  });

  test('the lookup by id holds every row', () => {
    expect(Object.keys(pptxFeatureById)).toEqual(pptxFeatures.map((row) => row.id));
  });

  test('every row has only the fields of a DOCX Editor row', () => {
    // The docs site renders this file with the DOCX Editor components. Evidence goes in
    // pptx-feature-evidence.ts, never in a new row field.
    const extra = pptxFeatures.flatMap((row) =>
      Object.keys(row)
        .filter((key) => !DOCX_ROW_FIELDS.has(key))
        .map((key) => `${row.id}: ${key}`)
    );
    expect(extra).toEqual([]);
  });

  test('the matrix file imports nothing, so the docs site can sync it as it is', () => {
    const source = readFileSync(join(import.meta.dir, 'pptx-features.ts'), 'utf8');
    expect(source.match(/^\s*(import|export \* from|export \{[^}]*\} from)\b.*$/gm) ?? []).toEqual(
      []
    );
  });

  test('every premium row names the EigenPal Pro License in its notes', () => {
    // The docs components show no tier, so the note is where a reader learns it.
    const silent = pptxFeatures
      .filter((row) => row.tier === 'premium' && !row.notes?.includes('EigenPal Pro License'))
      .map((row) => row.id);
    expect(silent).toEqual([]);
  });

  test('notes name no internal process, since the docs page shows them', () => {
    // AGENTS.md "Docs": public pages never name internal process, the corpus, or the harness.
    const internal =
      /reference PDF|harness|fixture|smoke|corpus|save-check|calibrat|#\d+|packages\//i;
    const leaks = pptxFeatures
      .filter((row) => internal.test(row.notes ?? ''))
      .map((row) => `${row.id}: ${internal.exec(row.notes ?? '')![0]}`);
    expect(leaks).toEqual([]);
  });

  test('every row declares a known tier and category', () => {
    const tiers = new Set<string>(FEATURE_TIERS);
    const categories = new Set(Object.keys(FEATURE_CATEGORY_LABELS));
    const invalid = pptxFeatures
      .filter((row) => !tiers.has(row.tier) || !categories.has(row.category))
      .map((row) => row.id);
    expect(invalid).toEqual([]);
  });

  test('shared status and tier values remain compatible with the DOCX matrix', () => {
    const statuses = {
      full: true,
      partial: true,
      'render-only': true,
      preserved: true,
      planned: true,
      none: true,
    } satisfies Record<FeatureStatus, true>;
    expect(FEATURE_TIERS).toEqual(['community', 'premium']);
    const wrong = pptxFeatures.flatMap((row) =>
      (['editing', 'rendering', 'roundTrip'] as const)
        .filter((axis) => !Object.hasOwn(statuses, row[axis]))
        .map((axis) => `${row.id}: ${axis} = ${row[axis]}`)
    );
    expect(wrong).toEqual([]);
  });
});

describe('pptx-features: the evidence file', () => {
  test('every row has an evidence entry', () => {
    const missing = pptxFeatures
      .filter((row) => !Object.hasOwn(FEATURE_EVIDENCE, row.id))
      .map((row) => row.id);
    expect(missing).toEqual([]);
  });

  test('every evidence entry belongs to a row', () => {
    const orphans = Object.keys(FEATURE_EVIDENCE).filter(
      (id) => !Object.hasOwn(pptxFeatureById, id)
    );
    expect(orphans).toEqual([]);
  });

  test('every named fixture exists in fixtures/, once per row', () => {
    const wrong = pptxFeatures.flatMap((row) => {
      const { fixtures } = evidence(row);
      return [
        ...fixtures
          .filter((name) => !existsSync(join(FIXTURES, `${name}.pptx`)))
          .map((name) => `${row.id}: ${name} does not exist`),
        ...repeats(fixtures).map((name) => `${row.id}: ${name} is repeated`),
      ];
    });
    expect(wrong).toEqual([]);
  });

  test('every fixture appears in at least one row', () => {
    const named = new Set(pptxFeatures.flatMap((row) => evidence(row).fixtures));
    const orphans = readdirSync(FIXTURES)
      .filter((file) => file.endsWith('.pptx'))
      .map((file) => file.slice(0, -'.pptx'.length))
      .filter((name) => !named.has(name));
    expect(orphans).toEqual([]);
  });
});

describe('pptx-features: evidence for claims', () => {
  test('rendering "full" needs a current reference PDF for every fixture of the row', () => {
    // Current: the PDF and the deck are the bytes recorded in fixtures/references.json. A stale
    // reference proves nothing about the deck as it is now.
    const manifest = readManifest(FIXTURES);
    const unproven = pptxFeatures
      .filter((row) => row.rendering === 'full')
      .flatMap((row) => {
        const { fixtures } = evidence(row);
        if (fixtures.length === 0) return [`${row.id}: no fixture`];
        return fixtures
          .filter(
            (name) => referenceState(join(FIXTURES, `${name}.pptx`), manifest).status !== 'current'
          )
          .map((name) => `${row.id}: ${name} has no current reference PDF`);
      });
    expect(unproven).toEqual([]);
  });

  test('rendering "full" needs fixtures that pass the harness: no slide of them is a known gap', () => {
    const gapped = new Set(
      Object.keys(FIXTURE_KNOWN_GAPS.slides).map((key) => key.slice(0, key.lastIndexOf('/')))
    );
    const unproven = pptxFeatures
      .filter((row) => row.rendering === 'full')
      .flatMap((row) =>
        evidence(row)
          .fixtures.filter((name) => gapped.has(name))
          .map((name) => `${row.id}: ${name} is in FIXTURE_KNOWN_GAPS`)
      );
    expect(unproven).toEqual([]);
  });

  test('rendering "full" owns no approximation but the page-wide text ones', () => {
    const owned = pptxFeatures
      .filter((row) => row.rendering === 'full')
      .flatMap((row) =>
        (evidence(row).approximations ?? [])
          .filter((id) => !PAGE_WIDE_APPROXIMATIONS.has(id))
          .map((id) => `${row.id}: ${id}`)
      );
    expect(owned).toEqual([]);
  });

  test('round trip "full" or "preserved" needs fixtures whose saved copies opened with no repair prompt', () => {
    const verified = new Set(SAVE_VERIFIED_FIXTURES);
    const unproven = pptxFeatures
      .filter((row) => CLAIMS_ROUND_TRIP.has(row.roundTrip))
      .flatMap((row) => {
        const { fixtures } = evidence(row);
        if (fixtures.length === 0) return [`${row.id}: no fixture`];
        return fixtures
          .filter((name) => !verified.has(name))
          .map((name) => `${row.id}: ${name} is not in SAVE_VERIFIED_FIXTURES`);
      });
    expect(unproven).toEqual([]);
  });

  test('round trip "full" or "preserved" needs an edit on one of the row\'s fixtures, checked for a prompt', () => {
    // feature-edits.test.ts runs each recipe: edit, save, reopen, and find the edit again. For a
    // "preserved" row, the edit is next to the kept content or of its box, and the check finds
    // the content as it was.
    const verified = new Set(Object.keys(SAVE_VERIFIED_FEATURE_EDITS));
    const unproven = pptxFeatures
      .filter((row) => CLAIMS_ROUND_TRIP.has(row.roundTrip))
      .flatMap((row) => {
        const recipe = FEATURE_EDITS[row.id];
        if (!recipe) return [`${row.id}: no recipe in FEATURE_EDITS`];
        if (!evidence(row).fixtures.includes(recipe.fixture))
          return [`${row.id}: the recipe fixture ${recipe.fixture} is not a fixture of the row`];
        if (!verified.has(row.id)) return [`${row.id}: not in SAVE_VERIFIED_FEATURE_EDITS`];
        return [];
      });
    expect(unproven).toEqual([]);
  });

  test('every save-verified feature edit names a row with a recipe', () => {
    const orphans = Object.keys(SAVE_VERIFIED_FEATURE_EDITS).filter(
      (id) => !pptxFeatureById[id] || !FEATURE_EDITS[id]
    );
    expect(orphans).toEqual([]);
  });

  test('every save-verified fixture exists in fixtures/', () => {
    const missing = SAVE_VERIFIED_FIXTURES.filter(
      (name) => !existsSync(join(FIXTURES, `${name}.pptx`))
    );
    expect(missing).toEqual([]);
  });
});

/** The approximation ids that layout reports on the slides of a fixture. */
function fixtureApproximations(name: string): Set<string> {
  const opened = readPptx(new Uint8Array(readFileSync(join(FIXTURES, `${name}.pptx`))));
  if (!opened.ok) throw new Error(`${name}: open failed`);
  const found = new Set<string>();
  for (let i = 0; ; i += 1) {
    const slide = resolveSlide(opened.pkg, i);
    if (!slide) return found;
    for (const d of layoutSlide(slide, createFixedMeasurer()).diagnostics)
      if (d.startsWith('approximation: ')) found.add(d.slice('approximation: '.length));
  }
}

// The fidelity page lists the known approximations, and `layout/__tests__/approximations.test.ts`
// keeps that list equal to `APPROXIMATIONS`. These tests link the matrix rows to the same catalog.
describe('pptx-features: the known approximations', () => {
  const catalog = new Set(APPROXIMATIONS.map((a) => a.id));

  test('every approximation a row names is in the catalog, once', () => {
    const wrong = pptxFeatures.flatMap((row) => {
      const ids = evidence(row).approximations ?? [];
      return [
        ...ids
          .filter((id) => !catalog.has(id))
          .map((id) => `${row.id}: ${id} is not in the catalog`),
        ...repeats(ids).map((id) => `${row.id}: ${id} is repeated`),
      ];
    });
    expect(wrong).toEqual([]);
  });

  test('every approximation of the catalog is named by at least one row', () => {
    const named = new Set(pptxFeatures.flatMap((row) => evidence(row).approximations ?? []));
    expect(APPROXIMATIONS.map((a) => a.id).filter((id) => !named.has(id))).toEqual([]);
  });

  test('every approximation that layout finds on a fixture is named by a row with such a fixture', () => {
    // The owner of an approximation must show it on one of its own fixtures. A row whose fixture
    // shows an approximation of another feature (a slide title shows `line-model`) does not name it.
    const shownOn = new Map<string, string[]>();
    const fixtures = [...new Set(pptxFeatures.flatMap((row) => evidence(row).fixtures))].sort();
    for (const name of fixtures)
      for (const id of fixtureApproximations(name))
        shownOn.set(id, [...(shownOn.get(id) ?? []), name]);
    const unowned = [...shownOn]
      .filter(
        ([id, names]) =>
          !pptxFeatures.some((row) => {
            const { fixtures: own, approximations = [] } = evidence(row);
            return approximations.includes(id) && own.some((name) => names.includes(name));
          })
      )
      .map(([id, names]) => `${id}: shown on ${names.join(', ')}; no row that names it lists one`);
    expect(unowned).toEqual([]);
  });
});
