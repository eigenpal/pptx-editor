// The example's own English text, in one place, as the chat example keeps its text. The
// editor's controls and the computed change summaries use the published locale catalog.

import type { RefreshChange } from '@pptx-editor.dev/react';

export const TEXT = {
  title: 'Document refresh',
  description:
    'Receive complete PPTX files from a processing server. The editor keeps the current slide ' +
    'and zoom, and flashes what changed.',
  limitations:
    'The mock server returns controlled sample files. Reset the example after you edit it.',
  loading: 'Loading the presentation',
  loadError: 'The presentation did not load.',
  start: 'Simulate server updates',
  processingLabel: 'Processing',
  cancel: 'Cancel updates',
  reset: 'Reset',
  modeLabel: 'Apply results as',
  modeReload: 'Reload',
  modeEdit: 'Undoable edit',
  computeChanges: 'Compute changes with Pro',
  goToChange: 'Go to the first new change',
  lateResult: 'Deliver an older result after the latest update',
  idle: 'Select Simulate server updates to start a job.',
  processing: 'The server is processing the presentation. You can keep working.',
  applied: (sequence: number, mode: string) => `Applied result ${sequence} of 2 as a ${mode}.`,
  movedToChange: 'Moved to the first new change.',
  refused: 'The editor refused a result.',
  refusedLate: 'The editor refused an older result. Newer content stays.',
  refusedEdits: 'You edited the presentation, so the editor kept your edit and refused the result.',
  sampleEdited: 'You edited the sample. Reset the example to request more updates.',
  cancelled: 'Updates cancelled. Accepted results stay in the presentation.',
  failed: 'The update failed.',
  resultCode: 'Result code',
  partial: (count: number) =>
    `The processor reported ${count} failed ${count === 1 ? 'operation' : 'operations'}:`,
  noChangeInformation: 'This result has no change information.',
  reviewChanges: 'Review changes',
  showChange: 'Show change',
  newChange: 'New',
  computed: 'Computed',
  changeStatus: {
    available: 'Available',
    deleted: 'Deleted',
    unavailable: 'No location',
    invalid: 'Invalid location',
    'unsupported-story': 'Unsupported location',
  } satisfies Record<RefreshChange['status'], string>,
  recoveryTitle: 'The previous presentation can be restored.',
  retry: 'Restore the previous presentation',
  download: 'Download the previous presentation',
};
