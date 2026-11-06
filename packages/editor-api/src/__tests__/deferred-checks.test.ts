/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
import { describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import * as api from '../index.ts';
import { ClientObject } from '../runtime/client-object.ts';
import { deferArgumentChecks, NOT_DEFERRED } from '../runtime/deferred-checks.ts';

interface Reference {
  symbols: Record<string, { members?: Record<string, unknown> }>;
}

const reference = JSON.parse(
  readFileSync(
    new URL('../../compat/reference/presentation.reference.json', import.meta.url),
    'utf8'
  )
) as Reference;

/** Public members that this API adds to an Office.js class, and that check their arguments. */
const EXTENSIONS: ReadonlySet<string> = new Set([]);

/** The members of `proto` and the classes above it that check their arguments at the sync. */
function deferredMembers(proto: object): string[] {
  const names: string[] = [];
  for (let p: object | null = proto; p && p !== ClientObject.prototype;) {
    for (const name of Object.getOwnPropertyNames(p)) {
      if (NOT_DEFERRED.has(name)) continue;
      const d = Object.getOwnPropertyDescriptor(p, name)!;
      if (typeof d.value === 'function' || d.set) names.push(name);
    }
    p = Object.getPrototypeOf(p) as object | null;
  }
  return names;
}

const classes = Object.entries(api).filter(
  (entry): entry is [string, abstract new (...args: never[]) => ClientObject] =>
    typeof entry[1] === 'function' && entry[1].prototype instanceof ClientObject
);

describe('argument checks at the sync', () => {
  test('the object model has classes to check', () => {
    expect(classes.length).toBeGreaterThan(40);
  });

  test('every call and property write that refuses at the sync is an Office.js member', () => {
    const unknown: string[] = [];
    for (const [name, cls] of classes) {
      const members = reference.symbols[name]?.members;
      if (!members) continue;
      for (const member of deferredMembers(cls.prototype))
        if (!(member in members) && !EXTENSIONS.has(member)) unknown.push(`${name}#${member}`);
    }
    // A name here is machinery that must stay out of the rule (add it to `NOT_DEFERRED`), or a
    // public member this API adds (add it to `EXTENSIONS`).
    expect(unknown).toEqual([]);
  });

  test('every class is wrapped once, so a refusal is queued once', () => {
    for (const [, cls] of classes) {
      deferArgumentChecks(cls.prototype, ClientObject.prototype);
      const before = Object.getOwnPropertyDescriptors(cls.prototype);
      deferArgumentChecks(cls.prototype, ClientObject.prototype);
      expect(Object.getOwnPropertyDescriptors(cls.prototype)).toEqual(before);
    }
  });
});
