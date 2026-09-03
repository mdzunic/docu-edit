// The OOXML ↔ model border-style mapping (SPEC-023 cell borders, SPEC-074
// paragraph and page borders). One module so the reader's cell borders, the
// paragraph-properties parser and the writer all agree on which `w:val` a
// model style is and vice versa — three copies of this table is exactly how
// a round trip starts losing styles.
import { resolveColor } from "../core/theme";
import type { DocumentTheme, ModelColor } from "../core/theme";
import { themeAttrsXml, themeRefFromAttrs } from "./theme";
import { roundWidthPt, type BorderLineStyle, type BorderSide } from "../core/borders";
import { eighthPointsToPt, ptToEighthPoints } from "../core/units";
import { cssColorToOoxml, ooxmlColorToCss } from "../core/colors";

/** The `w:val` this writes for each model style. OOXML's ~180-value
 * `ST_Border` includes both plain lines and clip-art "art" borders; the
 * dash-dot pair maps to the plain line values (`dotDash`/`dotDotDash`)
 * rather than the `dashDotStroked` art variant, so a saved file draws the
 * line the user picked in every consumer. Both spellings read back (see
 * OOXML_BORDER_STYLES). */
const BORDER_VAL: Record<BorderLineStyle, string> = {
  none: "nil",
  single: "single",
  double: "double",
  dashed: "dashed",
  dotted: "dotted",
  thick: "thick",
  dashDot: "dotDash",
  dashDotDot: "dotDotDash",
};

/** The `w:val`s that read back to each model style exactly. Everything else
 * falls through the pattern rules in `mapBorderLineStyle`. */
const OOXML_BORDER_STYLES: Record<string, BorderLineStyle> = {
  nil: "none",
  none: "none",
  single: "single",
  thick: "thick",
  double: "double",
  triple: "double",
  dotted: "dotted",
  dashed: "dashed",
  dashsmallgap: "dashed",
  dotdash: "dashDot",
  dashdotstroked: "dashDot",
  dotdotdash: "dashDotDot",
  dashdotdotstroked: "dashDotDot",
};

/** An OOXML `w:val` as the model's small enum (SPEC-023, widened by
 * SPEC-074 AC-3). Anything outside the table — including every art border,
 * which SPEC-074 keeps out of scope — reads as the nearest plain line the
 * model has, never as nothing. */
export function mapBorderLineStyle(val: string): BorderLineStyle {
  const v = val.trim().toLowerCase();
  const exact = OOXML_BORDER_STYLES[v];
  if (exact) {
    return exact;
  }
  // Word's remaining dash family: everything whose name says dot-dash is a
  // dash-dot pattern; the rest of the "dash*" names are plain dashes.
  if (v.includes("dotdotdash")) {
    return "dashDotDot";
  }
  if (v.includes("dotdash") || v.includes("dashdot")) {
    return "dashDot";
  }
  if (v.startsWith("dash")) {
    return "dashed";
  }
  if (v.startsWith("dot")) {
    return "dotted";
  }
  return "single";
}

/** The `w:val` for a model style. */
export function borderStyleVal(style: BorderLineStyle): string {
  return BORDER_VAL[style] ?? "single";
}

/** One `CT_Border` element as a `BorderSide`: `@w:sz` is eighths of a point,
 * `@w:color` an RRGGBB or "auto". Undefined when the element is missing or
 * carries no `w:val` at all. */
export function parseBorderElement(el: Element | undefined, attr: (el: Element, name: string) => string | null): BorderSide | undefined {
  if (!el) {
    return undefined;
  }
  const val = attr(el, "val");
  if (!val) {
    return undefined;
  }
  const style = mapBorderLineStyle(val);
  const sz = Number(attr(el, "sz") ?? "");
  const widthPt = style !== "none" && Number.isFinite(sz) && sz > 0 ? eighthPointsToPt(sz) : 0;
  // A @w:themeColor (with @w:themeTint/@w:themeShade) becomes a theme
  // reference (SPEC-077 AC-27); the literal @w:color stays the fallback.
  const ref = themeRefFromAttrs(attr(el, "themeColor"), attr(el, "themeTint"), attr(el, "themeShade"));
  const color: ModelColor = ref ?? ooxmlColorToCss(attr(el, "color")) ?? "auto";
  return { style, widthPt, color };
}

/** One `CT_Border` element from a `BorderSide`. `spacePt` is the OOXML
 * `@w:space` — the distance from the text, in whole points (SPEC-074 AC-28);
 * cell borders pass 0, the value SPEC-023 has always written. */
export function borderElementXml(tag: string, side: BorderSide, theme: DocumentTheme, spacePt = 0): string {
  const sz = side.style === "none" ? 0 : Math.max(1, ptToEighthPoints(roundWidthPt(side.widthPt)));
  // A theme reference writes the theme attributes beside the resolved
  // literal (SPEC-077 AC-30); a reference with no WordprocessingML spelling
  // (lumMod/lumOff, D8) writes the resolved literal alone (AC-32).
  const resolved = resolveColor(side.color, theme);
  const color = resolved !== "auto" ? (cssColorToOoxml(resolved) ?? "auto") : "auto";
  const themeAttrs = typeof side.color === "object" ? (themeAttrsXml(side.color, "themeColor") ?? "") : "";
  // @w:space is an unsigned byte capped at 31 points by the schema.
  const space = Math.min(31, Math.max(0, Math.round(spacePt)));
  return `<w:${tag} w:val="${borderStyleVal(side.style)}" w:sz="${sz}" w:space="${space}" w:color="${color}"${themeAttrs}/>`;
}
