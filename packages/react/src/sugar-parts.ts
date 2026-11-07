// Small helpers of the `<PptxEditor>` sugar component, kept out of `index.ts` for its line cap.

/** The props an object value gives, or none for `true`. */
export const partProps = <P extends object>(value: boolean | P | undefined): P =>
  typeof value === 'object' ? value : ({} as P);

/** The prop names the open-failure warning names. */
export const OPEN_FAILURE_ACTION_NAMES = {
  actions: 'renderOpenFailureActions',
  off: 'openFailure={false}',
} as const;
