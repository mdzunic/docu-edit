// The odp theme wire mapping (SPEC-094). The mechanics — the `de-theme-*`
// named-style scheme under `office:styles`, the parent-style ref carriers and
// the splice discipline — are shared with the word processor since SPEC-077
// and live in core/themeOdf.ts (AC-37: the text product never imports
// lib/presentation); this module re-exports them plus the deck-typed reader.
import { readThemeStyles, spliceThemeStyles, odpThemeBlockXml } from "../../core/themeOdf";
import type { DocumentTheme } from "../../core/theme";
import type { DeckTheme } from "../model";

export { themeParentSlot, themeGraphicParent, themeTextParent, odpThemeBlockXml } from "../../core/themeOdf";

/** Rewrites styles.xml's theme block in place: any existing `de-theme-*`
 * styles are removed and the canonical deck block for `theme` spliced into
 * `office:styles` — every other byte of styles.xml survives. Returns the
 * input unchanged when there is no `office:styles` element to splice into. */
export function spliceOdpThemeStyles(stylesXml: string, theme: DocumentTheme): string {
  return spliceThemeStyles(stylesXml, odpThemeBlockXml(theme));
}

/** Reads the deck theme back from styles.xml's `de-theme-*` block, or
 * undefined when the deck carries none (an unthemed odp — `model.theme` stays
 * absent and every colour stays a literal, AC-7). */
export function readOdpTheme(stylesXmlText: string, fallback: DocumentTheme): DeckTheme | undefined {
  return readThemeStyles(stylesXmlText, fallback);
}
