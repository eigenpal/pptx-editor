// The narrow layout of both example pages. Below 760 px the slide rail becomes a row of
// thumbnails under the slide, the notes start closed, and the controls go under the row.
// `styles.css` uses the same width.

import { useSyncExternalStore } from 'react';

const QUERY = '(max-width: 760px)';

function subscribe(onChange: () => void): () => void {
  const list = matchMedia(QUERY);
  list.addEventListener('change', onChange);
  return () => list.removeEventListener('change', onChange);
}

/** True when the window is narrow: the pages lay out in one column. */
export function useNarrow(): boolean {
  return useSyncExternalStore(subscribe, () => matchMedia(QUERY).matches);
}
