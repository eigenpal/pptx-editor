// A wait with a limit, for the open path: a request that never answers must not hold it.

/** Settles with `promise`, or with `late` after `ms`, whichever is first. */
export function within<T, L>(promise: Promise<T>, ms: number, late: L): Promise<T | L> {
  let timer: ReturnType<typeof setTimeout> | undefined;
  const wait = new Promise<L>((resolve) => {
    timer = setTimeout(() => resolve(late), ms);
  });
  return Promise.race<T | L>([promise, wait]).finally(() => clearTimeout(timer));
}
