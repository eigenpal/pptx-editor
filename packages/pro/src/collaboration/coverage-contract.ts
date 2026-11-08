/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/

// The collaboration coverage contract, after DOCX Editor's
// `packages/core/src/store/store/collaboration-coverage-contract.ts`. DOCX Editor commit adf0e178.
//
// Collaboration here is state-based (decision D26): the Yjs mirror diffs the package before and
// after each transaction, so any command that commits through the store replicates by
// construction. What can still go wrong is per command: a write the mirror does not carry, or a
// replica that rebuilds a different package. So the gate is per editing command, not per op kind.
//
// Every command in `EDITOR_COMMAND_TYPES` (`packages/core/src/editor/command-types.ts`) MUST have
// a two-peer convergence fixture (`src/__tests__/collaboration-coverage-fixtures.ts`), or an
// entry here with the reason it does not need one. Only a command that writes nothing to the
// deck belongs here, and its probe in `NO_WRITE_PROBES` proves it. The gate,
// `src/__tests__/collaboration-coverage.test.ts`, fails on a command with neither, on an entry
// for a command that no longer exists, on an entry for a command that has a fixture, and on an
// entry without a probe.
//
// Removing an entry is how a later change promotes a command: delete the line and its probe, and
// add the fixture. Keep each reason specific enough that a reader can decide whether it still
// holds. CONTRIBUTING.md, "Add a collaboration fixture", gives the steps.

import type { EditorCommandType } from '@pptx-editor.dev/core/editor';

const VIEW_SETTING =
  'A view setting of this editor: it writes no part of the deck (not even `viewProps.xml`) and ' +
  'makes no undo step, so there is nothing to replicate. Each participant keeps its own view.';

const SESSION_STATE =
  'Changes session state of this editor only and commits no store transaction, so there is ' +
  'nothing to replicate. The edit that follows it is a command with its own fixture.';

/** Commands declared not to need a convergence fixture, each with its reason. @internal */
export const COLLABORATION_UNCOVERED: ReadonlyMap<EditorCommandType, string> = new Map<
  EditorCommandType,
  string
>([
  ['toggleSmartGuides', VIEW_SETTING],
  ['showAnimationOrder', VIEW_SETTING],
  ['setTableBorderPen', VIEW_SETTING],
  ['setPresentationView', VIEW_SETTING],
  ['setSectionsCollapsed', VIEW_SETTING],
  [
    'toggleFormatPainter',
    `${SESSION_STATE} It arms the format painter; the paint is \`pasteFormat\` or a pointer gesture.`,
  ],
  [
    'copyFormat',
    `${SESSION_STATE} It copies the selection's formatting; \`pasteFormat\` applies it.`,
  ],
  ['toggleCrop', `${SESSION_STATE} It turns crop mode on or off; the crop is \`cropImage\`.`],
  [
    'selectTableRegion',
    `${SESSION_STATE} It selects table cells; the selection is presence, not content.`,
  ],
]);
