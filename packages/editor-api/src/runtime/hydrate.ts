/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// After DOCX Editor's `packages/editor-api/src/runtime/hydrate.ts` (commit 5993bd223945).
//
// Narrow one host answer to the shape an action asked for. A wrong shape is the host breaking
// its contract, so it is `GeneralException`, never a guess.

import type { AutomationHandle, AutomationValue } from '@pptx-editor.dev/core/automation';
import { PptxEditorError } from './errors.ts';

type ValueOf<K extends AutomationValue['kind']> = Extract<AutomationValue, { kind: K }>;

/** The answer as kind `kind`, or `GeneralException`. */
export function hydrated<K extends AutomationValue['kind']>(
  value: AutomationValue,
  kind: K,
  target: string
): ValueOf<K> {
  if (value.kind !== kind) throw new PptxEditorError({ code: 'GeneralException', target });
  return value as ValueOf<K>;
}

/** A handle answer, or null for a null answer. */
export function hydratedHandleOrNull(
  value: AutomationValue,
  target: string
): AutomationHandle | null {
  if (value.kind === 'null') return null;
  return hydrated(value, 'handle', target).handle;
}
