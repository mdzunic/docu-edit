// Maps ODF `style:text-properties` and `style:paragraph-properties` onto the
// document model's RunFormatting/ParagraphFormatting — the same model the
// .docx reader produces, so nothing downstream cares which format a document
// came from. Shared by the style-catalog parser (named and automatic styles
// carry the same property elements) and used for every resolved style chain.
import type { ParagraphAlignment, ParagraphFormatting, RunFormatting } from "../documentModel";
import { firstChildElement } from "../docx/xml";
import { foAttr, styleAttr, STYLE_NAMESPACE } from "./xml";
import { odfColorToCss, odfLengthToCm, odfLengthToPt, odfPercentage } from "./units";

/** Font families keyed by `style:font-face`'s name — built from the
 * `office:font-face-decls` of content.xml and styles.xml, so a
 * `style:font-name` reference resolves to the real family. */
export type FontFaceMap = Map<string, string>;

/** ODF heading/text styles often give the font size as a percentage of the
 * parent style's size ("130%"), which can only be resolved while walking the
 * inheritance chain — so it is carried alongside the absolute formatting
 * until resolution (see styles.ts). */
export interface OdfTextProperties {
  formatting: RunFormatting;
  fontSizePercent?: number;
}

function parseFontWeight(value: string | null): boolean | undefined {
  if (!value) {
    return undefined;
  }
  if (value === "normal") {
    return false;
  }
  if (value === "bold" || value === "bolder") {
    return true;
  }
  const numeric = Number(value);
  return Number.isNaN(numeric) ? undefined : numeric >= 600;
}

function parseFontStyle(value: string | null): boolean | undefined {
  if (!value) {
    return undefined;
  }
  return value === "italic" || value === "oblique";
}

/** `style:text-underline-style` / `style:text-line-through-style`: any line
 * style ("solid", "dotted", ...) means on; "none" explicitly turns it off. */
function parseLineStyleToggle(value: string | null): boolean | undefined {
  if (!value) {
    return undefined;
  }
  return value !== "none";
}

// The 15 standard OOXML highlight colours (the document model stores the
// OOXML name, matching the .docx reader) keyed by the hex value ODF's
// `fo:background-color` uses for the same colour. An ODF highlight outside
// this palette stays a raw CSS colour, which the renderer also accepts.
// Exported so the writer can invert the mapping when emitting.
export const HIGHLIGHT_NAMES: Record<string, string> = {
  "#ffff00": "yellow",
  "#00ff00": "green",
  "#00ffff": "cyan",
  "#ff00ff": "magenta",
  "#0000ff": "blue",
  "#ff0000": "red",
  "#00008b": "darkBlue",
  "#008b8b": "darkCyan",
  "#006400": "darkGreen",
  "#8b008b": "darkMagenta",
  "#8b0000": "darkRed",
  "#808000": "darkYellow",
  "#a9a9a9": "darkGray",
  "#d3d3d3": "lightGray",
  "#000000": "black",
};

/** Strips quoting from a font-family value ("'Liberation Serif'") and takes
 * the first entry of a comma-separated fallback list. */
export function cleanFontFamily(value: string): string | undefined {
  const first = value.split(",")[0].trim().replace(/^['"]|['"]$/g, "").trim();
  return first.length > 0 ? first : undefined;
}

/** Parses a `style:text-properties` element (on a style definition) into run
 * formatting. `fontFaces` resolves `style:font-name` references. */
export function parseTextProperties(
  textPropsEl: Element | undefined,
  fontFaces: FontFaceMap,
): OdfTextProperties {
  if (!textPropsEl) {
    return { formatting: {} };
  }
  const formatting: RunFormatting = {};
  const result: OdfTextProperties = { formatting };

  const bold = parseFontWeight(foAttr(textPropsEl, "font-weight"));
  if (bold !== undefined) {
    formatting.bold = bold;
  }

  const italic = parseFontStyle(foAttr(textPropsEl, "font-style"));
  if (italic !== undefined) {
    formatting.italic = italic;
  }

  const underline = parseLineStyleToggle(styleAttr(textPropsEl, "text-underline-style"));
  if (underline !== undefined) {
    formatting.underline = underline;
  }

  const strikethrough = parseLineStyleToggle(styleAttr(textPropsEl, "text-line-through-style"));
  if (strikethrough !== undefined) {
    formatting.strikethrough = strikethrough;
  }

  const fontName = styleAttr(textPropsEl, "font-name");
  const fontFamilyAttr = foAttr(textPropsEl, "font-family");
  const family = (fontName ? fontFaces.get(fontName) : undefined)
    ?? (fontName ? cleanFontFamily(fontName) : undefined)
    ?? (fontFamilyAttr ? cleanFontFamily(fontFamilyAttr) : undefined);
  if (family) {
    formatting.fontFamily = family;
  }

  const fontSize = foAttr(textPropsEl, "font-size");
  const fontSizePt = odfLengthToPt(fontSize);
  if (fontSizePt !== undefined) {
    formatting.fontSizePt = fontSizePt;
  } else {
    const percent = odfPercentage(fontSize);
    if (percent !== undefined) {
      result.fontSizePercent = percent;
    }
  }

  const color = odfColorToCss(foAttr(textPropsEl, "color"));
  if (color) {
    formatting.color = color;
  }

  const background = odfColorToCss(foAttr(textPropsEl, "background-color"));
  if (background) {
    formatting.highlight = HIGHLIGHT_NAMES[background] ?? background;
  }

  return result;
}

const ALIGNMENT_MAP: Record<string, ParagraphAlignment> = {
  left: "left",
  start: "left",
  right: "right",
  end: "right",
  center: "center",
  justify: "justify",
};

/** Parses a `style:paragraph-properties` element into paragraph formatting. */
export function parseParagraphProperties(paragraphPropsEl: Element | undefined): ParagraphFormatting {
  if (!paragraphPropsEl) {
    return {};
  }
  const formatting: ParagraphFormatting = {};

  const align = foAttr(paragraphPropsEl, "text-align");
  if (align) {
    const alignment = ALIGNMENT_MAP[align.toLowerCase()];
    if (alignment) {
      formatting.alignment = alignment;
    }
  }

  const marginLeft = odfLengthToCm(foAttr(paragraphPropsEl, "margin-left"));
  if (marginLeft !== undefined) {
    formatting.indentLeftCm = marginLeft;
  }
  const marginRight = odfLengthToCm(foAttr(paragraphPropsEl, "margin-right"));
  if (marginRight !== undefined) {
    formatting.indentRightCm = marginRight;
  }
  const textIndent = odfLengthToCm(foAttr(paragraphPropsEl, "text-indent"));
  if (textIndent !== undefined) {
    formatting.indentFirstLineCm = textIndent;
  }

  const marginTop = odfLengthToPt(foAttr(paragraphPropsEl, "margin-top"));
  if (marginTop !== undefined) {
    formatting.spaceBeforePt = marginTop;
  }
  const marginBottom = odfLengthToPt(foAttr(paragraphPropsEl, "margin-bottom"));
  if (marginBottom !== undefined) {
    formatting.spaceAfterPt = marginBottom;
  }

  // fo:line-height is either a percentage of the font size (proportional,
  // like OOXML's lineRule="auto"), an absolute length (fixed rule), or
  // "normal" (inherit). style:line-height-at-least maps to the same absolute
  // slot, matching how the .docx reader folds atLeast into lineSpacingPt.
  const lineHeight = foAttr(paragraphPropsEl, "line-height");
  const lineHeightPercent = odfPercentage(lineHeight);
  if (lineHeightPercent !== undefined) {
    formatting.lineSpacing = lineHeightPercent / 100;
  } else {
    const lineHeightPt = odfLengthToPt(lineHeight)
      ?? odfLengthToPt(styleAttr(paragraphPropsEl, "line-height-at-least"));
    if (lineHeightPt !== undefined) {
      formatting.lineSpacingPt = lineHeightPt;
    }
  }

  return formatting;
}

/** The `style:text-properties` child of a style definition element. */
export function textPropertiesOf(styleEl: Element): Element | undefined {
  return firstChildElement(styleEl, STYLE_NAMESPACE, "text-properties");
}

/** The `style:paragraph-properties` child of a style definition element. */
export function paragraphPropertiesOf(styleEl: Element): Element | undefined {
  return firstChildElement(styleEl, STYLE_NAMESPACE, "paragraph-properties");
}
