// Default-skip heuristics the spell checker applies beyond language/noProof
// (SPEC-035 AC-12.1, AC-12.2). The document model has no character-style or
// distinguished "Code Block" paragraph-style primitive (see
// documentModel.ts: only NamedStyle, paragraph-scoped, and both readers
// bake all resolved formatting — including a style's font — directly onto
// each run). A "Code" character style or "Code Block" paragraph style's
// monospace font is therefore already baked onto every run it applies to,
// so recognising common monospace font families covers both ACs with one
// rule instead of two.
const MONOSPACE_FONT_PATTERN =
  /courier|consolas|monaco|menlo|mono|source code|fira code|jetbrains|cascadia|inconsolata|lucida console/i;

export function isCodeStyledFont(fontFamily: string | undefined): boolean {
  return fontFamily !== undefined && MONOSPACE_FONT_PATTERN.test(fontFamily);
}
