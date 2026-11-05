// Copied from DOCX Editor packages/editor-api/scripts/lib/manifest-integrity.mjs at commit 5993bd223945cabdaf3900557d9fb59bfd8c2d8c.
// Changes: the presentation namespace and paths. PPTX Editor has no declaration-only fixture
// (DOCX Editor's `compat/docxeditor/declarations.ts`), so the authored-export check
// `validateAuthoredExportsAgainstManifest` and its support-type allowlist are not copied.
/**
 * Internal-consistency checks for `compat/manifest.json`.
 *
 * These are the checks `bun test` can make *offline*, entirely from files
 * already checked into the repository: they catch a manifest that
 * references a symbol/member no longer present in the frozen reference
 * fixture, a category list that drifted from `manifest.symbols`, or an
 * omission entry that is malformed or contradicts an active selection.
 *
 * What this deliberately cannot check offline: whether the manifest is
 * *complete* — i.e. whether some upstream presentation member exists that the
 * manifest neither selected nor recorded as a deliberate omission. That
 * requires the full upstream declaration text, which is only ever fetched
 * transiently by the scheduled drift-check job (never vendored here).
 */

const OMISSION_UID_PATTERN = /^(PowerPoint|OfficeExtension)\.[A-Za-z0-9_]+(#[A-Za-z0-9_]+)?$/;

/** Kept in lockstep with `reference-normalize.mjs`'s/`provenance.mjs`'s own `SCHEMA_VERSION` by convention, not by import — this file validates the *manifest's* schema, a separate document. */
const SUPPORTED_MANIFEST_SCHEMA_VERSION = 1;

function symbolLabel(symbolName) {
  return symbolName;
}

/**
 * Validates `manifest.json`'s own `schemaVersion` field — offline, no
 * reference fixture needed. Catches both a manifest that predates
 * versioning (field missing) and one written against a schema shape this
 * version of the tooling (`reference-normalize.mjs`, `shape-compare.mjs`,
 * etc.) does not understand.
 */
export function validateManifestSchemaVersion(manifest) {
  const version = manifest?.schemaVersion;
  if (version !== SUPPORTED_MANIFEST_SCHEMA_VERSION) {
    return [
      `manifest.schemaVersion: expected ${SUPPORTED_MANIFEST_SCHEMA_VERSION}, got ${JSON.stringify(version)}`,
    ];
  }
  return [];
}

export function validateManifestAgainstReference(manifest, referenceFixture) {
  const issues = [];
  const referenceSymbols = referenceFixture?.symbols ?? {};
  const manifestSymbols = manifest?.symbols ?? {};

  for (const [symbolName, selection] of Object.entries(manifestSymbols)) {
    const referenceSymbol = referenceSymbols[symbolName];
    if (!referenceSymbol) {
      issues.push(
        `manifest.symbols.${symbolLabel(symbolName)}: no corresponding reference symbol (stale manifest entry?)`
      );
      continue;
    }
    if (selection.isFunction) continue;
    const referenceMembers = referenceSymbol.members ?? {};
    for (const memberName of selection.members ?? []) {
      if (!(memberName in referenceMembers)) {
        issues.push(
          `manifest.symbols.${symbolLabel(symbolName)}.members: "${memberName}" has no corresponding reference member (stale manifest entry?)`
        );
      }
    }
  }

  for (const [categoryName, symbolNames] of Object.entries(manifest?.categories ?? {})) {
    for (const symbolName of symbolNames) {
      if (!(symbolName in manifestSymbols)) {
        issues.push(
          `manifest.categories.${categoryName}: "${symbolName}" is not a key in manifest.symbols`
        );
      }
    }
  }

  const selectedMemberUids = new Set();
  for (const [symbolName, selection] of Object.entries(manifestSymbols)) {
    for (const memberName of selection.members ?? []) {
      selectedMemberUids.add(`PowerPoint.${symbolName}#${memberName}`);
    }
  }

  // A class-level omission ("PowerPoint.Table") contradicts selecting that class at all: the reason
  // would describe the class as unpublished while its members are measured as published.
  const selectedSymbolUids = new Set(
    Object.entries(manifestSymbols)
      .filter(([, selection]) => !selection.isFunction)
      .map(([symbolName, selection]) => `${selection.namespace ?? 'PowerPoint'}.${symbolName}`)
  );

  for (const omission of manifest?.omissions ?? []) {
    const uid = omission?.uid ?? '(missing uid)';
    if (typeof omission?.reason !== 'string' || omission.reason.trim().length === 0) {
      issues.push(`manifest.omissions: "${uid}" is missing a non-empty reason`);
    }
    if (typeof omission?.uid !== 'string' || !OMISSION_UID_PATTERN.test(omission.uid)) {
      issues.push(
        `manifest.omissions: "${uid}" does not look like a PowerPoint.*/OfficeExtension.* UID`
      );
      continue;
    }
    if (selectedSymbolUids.has(omission.uid)) {
      issues.push(
        `manifest.omissions: "${omission.uid}" omits a whole class that manifest.symbols selects — record member-level omissions instead`
      );
    }
    if (selectedMemberUids.has(omission.uid)) {
      issues.push(
        `manifest.omissions: "${omission.uid}" contradicts an active selection — it is both selected and recorded as omitted`
      );
    }
  }

  return issues;
}
