// Keyboard keys, per platform.
//
// The editor intercepts Ctrl/Cmd+Home/End, PageUp/PageDown and Ctrl/Cmd+A
// itself, and those behave the same everywhere. Line-start/end and word jumps
// it leaves to the browser — and the browser binds them per platform. On
// macOS `Home` is a scroll key: pressing it in a contenteditable moves no
// caret at all, so a test that uses it silently does nothing and then asserts
// against a caret that never moved.
//
// `process.platform` is the machine running the browser, which is what picks
// the binding: darwin locally, linux in the factory's container.
const mac = process.platform === "darwin";

/** Every registry binding declares `mod: true` — Cmd on macOS, Ctrl
 * elsewhere (src/lib/shortcuts.ts) — so a spec that hard-codes "Control+z"
 * presses a combination the editor never claims on a Mac. The keystroke
 * then falls through to the browser, the command never runs, and the test
 * fails for a reason that has nothing to do with the behaviour it covers.
 * Compose every app shortcut from this. */
export const MOD = mac ? "Meta" : "Control";

/** How that same modifier is *spelled* in the UI — tooltips and the
 * shortcuts dialog render it through modifierKeyLabel() (SPEC-021 AC-3). */
export const MOD_LABEL = mac ? "Cmd" : "Ctrl";

export const LINE_START = mac ? "Meta+ArrowLeft" : "Home";
export const LINE_END = mac ? "Meta+ArrowRight" : "End";
export const WORD_LEFT = mac ? "Alt+ArrowLeft" : "Control+ArrowLeft";
export const WORD_RIGHT = mac ? "Alt+ArrowRight" : "Control+ArrowRight";

/** Shift-extended form of any of the combos above. */
export const extend = (combo: string): string => `Shift+${combo}`;
