// Generic reference-equality array windowing, shared by the track-changes
// diff engine (lib/trackChangesDiff.ts) and the version-history delta
// encoder (lib/versionHistory.ts). Product-neutral: nothing here knows what
// the elements are — only that callers keep untouched elements at the same
// object identity across an edit.

/** The trimmed middle of a before/after array pair: the common leading and
 * trailing runs' lengths, plus what remains of each side between them. */
export interface Window<T> {
  prefixLen: number;
  suffixLen: number;
  beforeMid: T[];
  afterMid: T[];
}

/** Trims the common leading/trailing run of *reference-equal* elements —
 * cheap and correct because untouched elements keep their exact object
 * identity across a command (see lib/trackChangesDiff.ts's header comment),
 * so only the elements a command actually touched ever land in the returned
 * "middle" window. */
export function windowOf<T>(before: T[], after: T[]): Window<T> {
  let start = 0;
  while (start < before.length && start < after.length && before[start] === after[start]) {
    start++;
  }
  let bEnd = before.length;
  let aEnd = after.length;
  while (bEnd > start && aEnd > start && before[bEnd - 1] === after[aEnd - 1]) {
    bEnd--;
    aEnd--;
  }
  return { prefixLen: start, suffixLen: before.length - bEnd, beforeMid: before.slice(start, bEnd), afterMid: after.slice(start, aEnd) };
}
