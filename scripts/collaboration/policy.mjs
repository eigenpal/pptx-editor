// Copyright (c) 2026 Eigenpal, Inc. Licensed under Apache-2.0.
// Compatibility decisions follow DOCX Editor policy.mjs, commit 13f08556889ef7c192023f92aa865cc9cc8e3fe4.
import assert from 'node:assert/strict';
export const fields = [
  'protocolVersion',
  'sharedSchemaVersion',
  'repairVersion',
  'canonicalModelVersion',
];
export function sourceFormat(source, schemaSource) {
  const schema = schemaSource.match(/SCHEMA_VERSION = (\d+)/)?.[1];
  assert.ok(schema, 'Missing shared schema constant');
  const values = fields.map((field, index) =>
    index === 1 ? Number(schema) : Number(source.match(new RegExp(`${field}: (\\d+)`))?.[1])
  );
  if (!source.includes('COLLABORATION_VERSIONS'))
    return { format: `pptx-collaboration-${schema}`, values: [undefined, Number(schema)] };
  assert.ok(
    values.every((value) => Number.isSafeInteger(value) && value > 0),
    'Invalid format constants'
  );
  return { format: `pptx-collaboration:${values.join('.')}`, values };
}
export function validateDecision(record, before, after) {
  assert.ok(['no-impact', 'compatible', 'migration-required'].includes(record.impact));
  assert.equal(record.before, before.format, 'Decision must name the base format');
  assert.equal(record.after, after.format, 'Decision must name the candidate format');
  assert.ok(typeof record.reason === 'string' && record.reason.trim().length > 0);
  assert.ok(Array.isArray(record.tests) && record.tests.length > 0);
  assert.ok(Array.isArray(record.fields));
  const changed = before.format !== after.format;
  const affected = changed
    ? fields.filter((_, index) => before.values[index] !== after.values[index])
    : [];
  assert.deepEqual(
    [...record.fields].sort(),
    [...affected].sort(),
    'Record the fields that actually change'
  );
  assert.equal(
    record.impact === 'migration-required',
    changed,
    'Migration impact must match the format change'
  );
  before.values.forEach((value, index) => {
    if (value !== undefined)
      assert.ok(after.values[index] >= value, 'Format fields cannot decrease');
  });
}
export function repositoryPath(path) {
  assert.ok(
    typeof path === 'string' && path.length > 0 && !path.startsWith('/') && !path.includes('\\')
  );
  assert.ok(
    path.split('/').every((part) => part !== '..' && part !== '.' && part !== ''),
    'Evidence must stay inside the repository'
  );
  return path;
}
