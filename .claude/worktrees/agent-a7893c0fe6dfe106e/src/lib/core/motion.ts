// Motion helpers (SPEC-060): the library highlights a returning document's
// row for one `--motion-base` interval, but stays still under
// `prefers-reduced-motion` (where the token is 0). Framework-free DOM reads,
// so a second product can reuse them. See lib/core/README.md (app plumbing).
export function prefersReducedMotion(): boolean {
  return (
    typeof window.matchMedia === "function" &&
    window.matchMedia("(prefers-reduced-motion: reduce)").matches
  );
}

// The app's motion interval from the `--motion-base` token, in milliseconds.
// Falls back to 200ms if the token is absent or unparseable. The token is set
// to 0ms under reduced motion (tokens.css), so callers may skip animating
// entirely when this returns 0.
export function motionBaseMs(): number {
  const raw = getComputedStyle(document.documentElement).getPropertyValue("--motion-base").trim();
  const value = parseFloat(raw);
  if (!Number.isFinite(value)) {
    return 200;
  }
  return raw.endsWith("ms") ? value : value * 1000;
}
