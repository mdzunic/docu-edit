// The ODF ↔ model border mapping (SPEC-023 cell borders, SPEC-074 paragraph
// and page borders). ODF expresses style, width and colour in one `fo:border`
// shorthand ("0.5pt solid #000000"), drawn from CSS2's keyword set — which
// has no dash-dot pattern, so those two styles write as their nearest
// equivalent and are reported dropped (SPEC-074 AC-32).
import { roundWidthPt, type BorderLineStyle, type BorderSide } from "../core/borders";
import { odfColorToCss, odfLengthToPt } from "./units";

/** The extension namespace this product's own attributes live in — the
 * convention SPEC-073 established for ODF-inexpressible model fields (see
 * odf/shape.ts's `de:shadow-blur-pt`). Used here for the paragraph
 * `between` edge and the page border's reference edge, neither of which ODF
 * has a standard attribute for; a consumer that does not know the namespace
 * ignores them, and this product reads them back exactly. */
export const DE_BORDER_NAMESPACE = "urn:docu-edit:borders";

/** The CSS2 keyword each model style writes as. `thick` is a solid line
 * whose weight is its width — ODF has no separate weight class, and unlike
 * the dash-dot pair that is not a loss worth reporting (the line drawn is
 * the line asked for). */
const ODF_BORDER_STYLE: Record<BorderLineStyle, string> = {
  none: "none",
  single: "solid",
  double: "double",
  dashed: "dashed",
  dotted: "dotted",
  thick: "solid",
  dashDot: "dashed",
  dashDotDot: "dashed",
};

/** The styles ODF cannot express, which write as their nearest equivalent
 * and are reported as a `paragraphFormatting` DropRecord (SPEC-074 AC-32). */
export function odfLosesBorderStyle(style: BorderLineStyle): boolean {
  return style === "dashDot" || style === "dashDotDot";
}

/** ODF's CSS2-derived border style keywords as the model's enum — the same
 * simplification the .docx reader applies. `groove`/`ridge`/`inset`/`outset`
 * are 3-D single lines with no model equivalent and read as `single`. */
export function mapOdfBorderStyle(word: string): BorderLineStyle {
  const w = word.trim().toLowerCase();
  if (w === "none" || w === "hidden") {
    return "none";
  }
  if (w === "double") {
    return "double";
  }
  if (w === "dashed") {
    return "dashed";
  }
  if (w === "dotted") {
    return "dotted";
  }
  return "single";
}

/** ODF's `fo:border`-family shorthand ("0.5pt solid #ff0000") or
 * "none"/"hidden" as a `BorderSide` (SPEC-023, SPEC-074 AC-29). The three
 * components may arrive in any order per CSS2, but every producer this
 * product has met writes width-style-colour, and a colour is required for
 * the model's CSS-ready `color` — a shorthand without one reads as black. */
export function parseBorderShorthand(value: string | null): BorderSide | undefined {
  if (!value) {
    return undefined;
  }
  const trimmed = value.trim();
  if (trimmed === "none" || trimmed === "hidden") {
    return { style: "none", widthPt: 0, color: "#000000" };
  }
  const match = /^(\S+)\s+(\S+)\s+(#[0-9a-fA-F]{6})$/.exec(trimmed);
  if (!match) {
    return undefined;
  }
  const widthPt = odfLengthToPt(match[1]) ?? 1;
  const style = mapOdfBorderStyle(match[2]);
  const color = odfColorToCss(match[3]) ?? "#000000";
  return { style, widthPt: roundWidthPt(widthPt), color };
}

/** A `BorderSide` as ODF's `fo:border`-family shorthand. */
export function borderShorthand(side: BorderSide): string {
  if (side.style === "none" || side.widthPt <= 0) {
    return "none";
  }
  const color = side.color === "auto" ? "#000000" : side.color;
  return `${roundWidthPt(side.widthPt)}pt ${ODF_BORDER_STYLE[side.style] ?? "solid"} ${color}`;
}
