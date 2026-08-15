// Mac-vs-everyone-else detection, isolated in its own module so both the
// toolbar tooltips (SPEC-012 AC-2) and any future shortcut-reference UI can
// share one answer for "which modifier key does this platform use".

/** Whether the app is running on a Mac (Cmd-based shortcuts) rather than
 * Windows/Linux (Ctrl-based). Reads `navigator.platform`, falling back to
 * `userAgent` for browsers that no longer expose `platform`; defaults to
 * false (Ctrl) when neither is available, e.g. in a non-browser test. */
export function isMacPlatform(): boolean {
  const nav = typeof navigator === "undefined" ? undefined : navigator;
  const signal = nav?.platform || nav?.userAgent || "";
  return /Mac|iPhone|iPad|iPod/.test(signal);
}

/** The modifier key label to show in a shortcut hint: "Cmd" on Mac, "Ctrl"
 * elsewhere (SPEC-012 AC-2). */
export function modifierKeyLabel(): string {
  return isMacPlatform() ? "Cmd" : "Ctrl";
}
