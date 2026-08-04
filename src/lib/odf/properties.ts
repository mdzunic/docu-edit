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

// The OOXML highlight palette (the document model stores the OOXML name,
// matching the .docx reader) keyed by the hex value ODF's
// `fo:background-color` uses for the same colour. 16 entries: the 15
// colours Word's highlight picker offers plus `white`, which
// ST_HighlightColor admits even though the picker never shows it — mapping
// it keeps an incoming `w:highlight w:val="white"` alive across a
// `.docx` → `.odt` → `.docx` trip (SPEC-013). A highlight outside this
// palette stays a raw CSS colour, which the renderer also accepts. This is
// the single source for the palette: the `.docx` reader/writer, the `.odt`
// writer and the canvas renderer all derive from it.
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
  "#ffffff": "white",
};

/** The inverse palette: OOXML highlight name → hex colour. */
export const HIGHLIGHT_HEX: Record<string, string> = Object.fromEntries(
  Object.entries(HIGHLIGHT_NAMES).map(([hex, name]) => [name, hex]),
);

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

  const backgroundAttr = foAttr(textPropsEl, "background-color");
  if (backgroundAttr?.trim().toLowerCase() === "transparent") {
    // An explicit "no highlight", distinct from the attribute being absent —
    // it must survive to override an inherited background (SPEC-013).
    formatting.highlight = false;
  } else {
    const background = odfColorToCss(backgroundAttr);
    if (background) {
      formatting.highlight = HIGHLIGHT_NAMES[background] ?? background;
    }
  }

  const vertAlign = parseTextPosition(styleAttr(textPropsEl, "text-position"));
  if (vertAlign !== undefined) {
    formatting.vertAlign = vertAlign;
  }

  return result;
}

/** `style:text-position`'s first value is the baseline shift: the keywords
 * "sub"/"super" or a signed percentage ("-33%" is below the baseline,
 * "33%" above, "0%" on it). The optional second value (font height) and a
 * zero shift both leave the model's vertAlign unset. */
function parseTextPosition(value: string | null): "subscript" | "superscript" | undefined {
  if (!value) {
    return undefined;
  }
  const shift = value.trim().split(/\s+/)[0];
  if (shift === "sub") {
    return "subscript";
  }
  if (shift === "super") {
    return "superscript";
  }
  const percent = Number.parseFloat(shift);
  if (Number.isNaN(percent) || !shift.endsWith("%") || percent === 0) {
    return undefined;
  }
  return percent < 0 ? "subscript" : "superscript";
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

  // ODF's fo:margin-left/right follow XSL and are PHYSICAL edges. They are
  // parsed into the logical start/end slots as if the paragraph were LTR;
  // the reader swaps start/end once the style chain is fully resolved and
  // the effective writing mode is known (SPEC-014, see swapRtlIndents in
  // reader.ts). fo:text-indent is start-relative already and needs no swap.
  const marginLeft = odfLengthToCm(foAttr(paragraphPropsEl, "margin-left"));
  if (marginLeft !== undefined) {
    formatting.indentStartCm = marginLeft;
  }
  const marginRight = odfLengthToCm(foAttr(paragraphPropsEl, "margin-right"));
  if (marginRight !== undefined) {
    formatting.indentEndCm = marginRight;
  }
  const textIndent = odfLengthToCm(foAttr(paragraphPropsEl, "text-indent"));
  if (textIndent !== undefined) {
    formatting.indentFirstLineCm = textIndent;
  }

  // style:writing-mode: "rl-tb"/"rl" are the right-to-left modes LibreOffice
  // writes. An explicit LTR mode is recorded so it can override an inherited
  // RTL during chain merging; the reader normalises it away — the model only
  // ever holds "rtl" (SPEC-014, see ParagraphDirection).
  const writingMode = styleAttr(paragraphPropsEl, "writing-mode");
  if (writingMode !== null) {
    const mode = writingMode.trim().toLowerCase();
    if (mode === "rl-tb" || mode === "rl") {
      formatting.direction = "rtl";
    } else if (mode === "lr-tb" || mode === "lr") {
      formatting.direction = "ltr";
    }
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
  // like OOXML's lineRule="auto"), an absolute length (a FIXED height —
  // OOXML's "exact"), or "normal" (inherit). style:line-height-at-least is
  // the minimum-height rule, which is the model's absent-rule default
  // (SPEC-014, see ParagraphFormatting.lineSpacingRule).
  const lineHeight = foAttr(paragraphPropsEl, "line-height");
  const lineHeightPercent = odfPercentage(lineHeight);
  if (lineHeightPercent !== undefined) {
    formatting.lineSpacing = lineHeightPercent / 100;
  } else {
    const exactPt = odfLengthToPt(lineHeight);
    if (exactPt !== undefined) {
      formatting.lineSpacingPt = exactPt;
      formatting.lineSpacingRule = "exact";
    } else {
      const atLeastPt = odfLengthToPt(styleAttr(paragraphPropsEl, "line-height-at-least"));
      if (atLeastPt !== undefined) {
        formatting.lineSpacingPt = atLeastPt;
      }
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
