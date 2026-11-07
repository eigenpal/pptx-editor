// The changed regions between two texts, so a text field (the notes pane) writes back only what
// the user changed and every unchanged run keeps its formatting. Format-neutral: no
// PresentationML here.

/** One changed region: `a[aStart, aEnd)` becomes `b[bStart, bEnd)`. */
export interface TextHunk {
  readonly aStart: number;
  readonly aEnd: number;
  readonly bStart: number;
  readonly bEnd: number;
}

/**
 * The changed regions between `a` and `b`, in order, by UTF-16 code unit: Myers' O((N+M)D)
 * shortest edit script ("An O(ND) Difference Algorithm and Its Variations", 1986). Texts that
 * differ in more than `maxEdits` places give one region over the whole changed middle, which
 * bounds the time and memory a large paste costs.
 */
export function diffHunks(a: string, b: string, maxEdits = 400): TextHunk[] {
  let pre = 0;
  while (pre < a.length && pre < b.length && a[pre] === b[pre]) pre += 1;
  let suf = 0;
  while (
    suf < a.length - pre &&
    suf < b.length - pre &&
    a[a.length - 1 - suf] === b[b.length - 1 - suf]
  )
    suf += 1;
  const n = a.length - pre - suf;
  const m = b.length - pre - suf;
  if (n === 0 && m === 0) return [];
  const whole: TextHunk[] = [{ aStart: pre, aEnd: pre + n, bStart: pre, bEnd: pre + m }];
  if (n === 0 || m === 0) return whole;
  const max = Math.min(n + m, maxEdits);
  const off = max + 1;
  const v = new Int32Array(2 * max + 3);
  const trace: Int32Array[] = [];
  let done = false;
  for (let d = 0; d <= max && !done; d += 1) {
    trace.push(v.slice());
    for (let k = -d; k <= d; k += 2) {
      let x =
        k === -d || (k !== d && v[off + k - 1] < v[off + k + 1])
          ? v[off + k + 1]
          : v[off + k - 1] + 1;
      let y = x - k;
      while (x < n && y < m && a[pre + x] === b[pre + y]) {
        x += 1;
        y += 1;
      }
      v[off + k] = x;
      if (x >= n && y >= m) {
        done = true;
        break;
      }
    }
  }
  if (!done) return whole;
  // Walk back through the snapshots: each step is one deletion or insertion and a diagonal.
  const ops: number[] = []; // 0 equal, 1 delete from a, 2 insert from b; built backwards
  let x = n;
  let y = m;
  for (let d = trace.length - 1; d >= 0; d -= 1) {
    const vd = trace[d];
    const k = x - y;
    const prevK = k === -d || (k !== d && vd[off + k - 1] < vd[off + k + 1]) ? k + 1 : k - 1;
    const prevX = vd[off + prevK];
    const prevY = prevX - prevK;
    while (x > prevX && y > prevY) {
      ops.push(0);
      x -= 1;
      y -= 1;
    }
    if (d > 0) {
      if (x === prevX) {
        ops.push(2);
        y -= 1;
      } else {
        ops.push(1);
        x -= 1;
      }
    }
  }
  const out: TextHunk[] = [];
  let ia = pre;
  let ib = pre;
  let open: { aStart: number; bStart: number } | null = null;
  for (let i = ops.length - 1; i >= 0; i -= 1) {
    const op = ops[i];
    if (op === 0) {
      if (open) out.push({ ...open, aEnd: ia, bEnd: ib });
      open = null;
      ia += 1;
      ib += 1;
      continue;
    }
    open ??= { aStart: ia, bStart: ib };
    if (op === 1) ia += 1;
    else ib += 1;
  }
  if (open) out.push({ ...open, aEnd: ia, bEnd: ib });
  return out;
}
