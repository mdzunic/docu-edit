// The playback key map — the NORMATIVE data shared by the play view
// (src/views/PresentationPlayView.vue) and, downstream, SPEC-088's HTML export
// template so their navigation cannot drift: the same `advanceKeys`/`backKeys`
// arrays drive both surfaces. SPEC-085/086 insert transition/build steps into
// the advance pipeline but keep this key map; SPEC-079's editor imports these
// tables for its two-key (PageUp/PageDown) slide navigation, so the maps stay
// in sync. Plain data + a pure classifier — no Vue, no DOM (a KeyboardEvent's
// `.key` string in, an action out). SPEC-100 adds the blank-screen rows (B/W,
// which SPEC-084 deferred) and its presenter window classifies through this
// same function; SPEC-088's export deliberately ignores the blank rows (see
// htmlExport.ts's keyTableJson).

/** The navigation actions the key map resolves to. "advance"/"back" step one
 * step (this spec: one slide); "first"/"last" jump to the ends. Number-then-
 * Enter jumps are handled by the caller (it owns the digit buffer) before the
 * Enter reaches this classifier. SPEC-100's "blackout"/"whiteout" toggle the
 * blank screen; they resolve after every navigation row, so no existing key
 * changes its action. */
export type PlaybackKeyAction = "advance" | "back" | "first" | "last" | "blackout" | "whiteout";

// Single-letter entries are matched case-insensitively (see `normaliseKey`), so
// N and P appear here lowercased. Space arrives as " " on modern browsers and
// "Spacebar" on legacy ones; both advance.
export const advanceKeys: readonly string[] = [
  "ArrowRight",
  "ArrowDown",
  " ",
  "Spacebar",
  "PageDown",
  "Enter",
  "n",
];

export const backKeys: readonly string[] = [
  "ArrowLeft",
  "ArrowUp",
  "PageUp",
  "p",
  "Backspace",
];

export const firstKeys: readonly string[] = ["Home"];

export const lastKeys: readonly string[] = ["End"];

/** SPEC-100: blank the audience to black (PowerPoint's B and period). */
export const blackoutKeys: readonly string[] = ["b", "."];

/** SPEC-100: blank the audience to white (PowerPoint's W and comma). */
export const whiteoutKeys: readonly string[] = ["w", ","];

/** Lower-case a single-character key so letter bindings (n, p) match whether or
 * not Shift is held; multi-character keys (ArrowRight, PageDown, …) pass
 * through unchanged. */
export function normaliseKey(key: string): string {
  return key.length === 1 ? key.toLowerCase() : key;
}

/** The navigation action a bare key resolves to, or null if it owns no action.
 * The caller is responsible for having already filtered out modifier combos it
 * does not own (Ctrl/Cmd/Alt) and events from text inputs. */
export function playbackActionForKey(key: string): PlaybackKeyAction | null {
  const k = normaliseKey(key);
  if (advanceKeys.includes(k)) return "advance";
  if (backKeys.includes(k)) return "back";
  if (firstKeys.includes(k)) return "first";
  if (lastKeys.includes(k)) return "last";
  if (blackoutKeys.includes(k)) return "blackout";
  if (whiteoutKeys.includes(k)) return "whiteout";
  return null;
}
