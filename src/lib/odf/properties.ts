// Maps ODF `style:text-properties` and `style:paragraph-properties` onto the
// document model's RunFormatting/ParagraphFormatting — the same model the
// .docx reader produces, so nothing downstream cares which format a document
// came from. Shared by the style-catalog parser (named and automatic styles
// carry the same property elements) and used for every resolved style chain.
import type {
  ParagraphAlignment,
  ParagraphBorders,
  ParagraphFormatting,
  RunFormatting,
  TabLeader,
  TabStop,
  TabStopAlignment,
} from "../documentModel";
import { firstChildElement } from "../core/xml";
import { HIGHLIGHT_NAMES } from "../core/colors";
import { composeLanguageTag, ODF_NO_LINGUISTIC_CONTENT } from "../core/language";
import { foAttr, styleAttr, STYLE_NAMESPACE } from "../core/odfXml";
import { odfColorToCss, odfLengthToCm, odfLengthToPt, odfPercentage } from "./units";
import { DE_BORDER_NAMESPACE, parseBorderShorthand as parseOdfBorderShorthand } from "./borders";
import { explicitNoBorders, normaliseParagraphBorders, NO_SHADING_FILL } from "../paragraphBorders";

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

// The OOXML/ODF highlight palette lives in src/lib/core/colors.ts (it is
// product-neutral colour data, not ODF parsing).
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

  // SPEC-035: language, and "do not check spelling" via ODF's "no
  // linguistic content" convention (fo:language="zxx") — ODF has no
  // separate boolean flag the way .docx's w:noProof is, so a run marked
  // this way maps to `noProof: true` with no `lang` (see lib/language.ts's
  // ODF_NO_LINGUISTIC_CONTENT).
  const language = foAttr(textPropsEl, "language");
  if (language) {
    if (language.trim().toLowerCase() === ODF_NO_LINGUISTIC_CONTENT) {
      formatting.noProof = true;
    } else {
      formatting.lang = composeLanguageTag(language, foAttr(textPropsEl, "country"));
    }
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

  // Pagination flags (SPEC-028 AC-22..29). Only the "on" values are stored:
  // ODF's own defaults ("auto" keeps, non-zero widows/orphans) read as
  // absent, so LibreOffice's routine fo:keep-with-next="auto" adds no model
  // noise — the model's absent value already means the same thing.
  if (foAttr(paragraphPropsEl, "keep-with-next") === "always") {
    formatting.keepNext = true;
  }
  if (foAttr(paragraphPropsEl, "keep-together") === "always") {
    formatting.keepLines = true;
  }
  // Only "page" maps here: the model has no column-break formatting flag
  // (columns are SPEC-030), so fo:break-before="column" is surfaced through
  // the style definition's columnBreakBefore (styles.ts) and materialises a
  // column BreakBlock on open (SPEC-028 AC-10/11) instead of a formatting
  // flag; fo:break-after is out of scope. Whether a "page" value
  // materialises a BreakBlock or stays this formatting flag is the
  // reader's call (SPEC-028 AC-26/32, see reader.ts) — this parser reports
  // what the properties element says.
  if (foAttr(paragraphPropsEl, "break-before") === "page") {
    formatting.pageBreakBefore = true;
  }
  // Widow/orphan control is off only when BOTH minimums are an explicit 0
  // (AC-28/29) — the shape the writer emits for widowControl: false; any
  // other value keeps the model's absent-means-on default.
  if (
    foAttr(paragraphPropsEl, "widows")?.trim() === "0" &&
    foAttr(paragraphPropsEl, "orphans")?.trim() === "0"
  ) {
    formatting.widowControl = false;
  }

  // Custom tab stops (SPEC-032 AC-19..25): style:tab-stops/style:tab-stop.
  // "char" is ODF's decimal alignment; the leader is recovered from
  // style:leader-text when present, else from style:leader-style, so
  // foreign files that only write the keyword keep their leaders too.
  const tabStopsEl = firstChildElement(paragraphPropsEl, STYLE_NAMESPACE, "tab-stops");
  if (tabStopsEl) {
    const stops: TabStop[] = [];
    for (const stopEl of Array.from(tabStopsEl.children)) {
      if (stopEl.namespaceURI !== STYLE_NAMESPACE || stopEl.localName !== "tab-stop") {
        continue;
      }
      const positionCm = odfLengthToCm(styleAttr(stopEl, "position"));
      if (positionCm === undefined || positionCm < 0) {
        continue;
      }
      const type = styleAttr(stopEl, "type") ?? "left";
      const alignment: TabStopAlignment =
        type === "char"
          ? "decimal"
          : type === "right" || type === "center"
            ? type
            : "left";
      const stop: TabStop = { positionCm, alignment };
      const leader = tabLeaderOf(styleAttr(stopEl, "leader-text"), styleAttr(stopEl, "leader-style"));
      if (leader) {
        stop.leader = leader;
      }
      stops.push(stop);
    }
    if (stops.length > 0) {
      formatting.tabStops = stops.sort((a, b) => a.positionCm - b.positionCm);
    }
  }

  // Paragraph borders and shading (SPEC-074 AC-29). `style:join-border` is
  // read but needs no model slot: its ODF default is `true`, which is the
  // merge behaviour borderRuns already implements for every paragraph.
  const borders = parseOdfParagraphBorders(paragraphPropsEl);
  if (borders !== undefined) {
    formatting.borders = borders;
  }
  // `fo:background-color="transparent"` is ODF's "no fill", and reads as the
  // explicit no-fill override rather than as absent: the attribute's
  // presence is this style saying its background is nothing, which is what
  // switches an inherited fill back off (AC-17). resolveParagraphStyle drops
  // it again where the chain underneath had no fill to override.
  const background = foAttr(paragraphPropsEl, "background-color");
  const shadingFill = background === NO_SHADING_FILL ? NO_SHADING_FILL : odfColorToCss(background);
  if (shadingFill !== undefined) {
    formatting.shadingFill = shadingFill;
  }

  return formatting;
}

/** A `style:paragraph-properties` element's borders (SPEC-074 AC-29): each
 * `fo:border-*` edge, falling back to the `fo:border` shorthand ODF lets a
 * style set all four with, plus `fo:padding-*` (or `fo:padding`) for the
 * border-to-text space. The `between` edge has no ODF attribute and rides in
 * this product's extension namespace (see odf/borders.ts). */
function parseOdfParagraphBorders(props: Element): ParagraphBorders | undefined {
  const shorthand = foAttr(props, "border");
  const paddingAll = odfLengthToPt(foAttr(props, "padding"));
  const borders: ParagraphBorders = {};
  const paddingPt: { top?: number; bottom?: number; left?: number; right?: number } = {};
  let anyPadding = false;
  for (const edge of ["top", "bottom", "left", "right"] as const) {
    const side = parseOdfBorderShorthand(foAttr(props, `border-${edge}`) ?? shorthand);
    if (!side) {
      continue;
    }
    borders[edge] = side;
    const padding = odfLengthToPt(foAttr(props, `padding-${edge}`)) ?? paddingAll;
    if (padding !== undefined) {
      paddingPt[edge] = padding;
      anyPadding = true;
    }
  }
  const between = parseOdfBorderShorthand(props.getAttributeNS(DE_BORDER_NAMESPACE, "border-between"));
  if (between) {
    borders.between = between;
  }
  if (anyPadding) {
    borders.paddingPt = paddingPt;
  }
  const painted = normaliseParagraphBorders(borders);
  if (painted !== undefined) {
    return painted;
  }
  // `fo:border="none"` (or a "none" on every edge) is a border stated and
  // stated as nothing — ODF's explicit override of an inherited box
  // (SPEC-074 AC-17), which reads back as a borders object that is present
  // and paints nothing. An element that mentions no border at all still
  // reads as undefined, i.e. "inherit". resolveParagraphStyle drops the
  // override again where the chain underneath had no box to override.
  return Object.keys(borders).length > 0 ? explicitNoBorders() : undefined;
}

function tabLeaderOf(leaderText: string | null, leaderStyle: string | null): TabLeader | undefined {
  if (leaderText === ".") {
    return "dot";
  }
  if (leaderText === "-") {
    return "hyphen";
  }
  if (leaderText === "_") {
    return "underscore";
  }
  if (leaderStyle === "dotted") {
    return "dot";
  }
  if (leaderStyle === "dash" || leaderStyle === "long-dash") {
    return "hyphen";
  }
  if (leaderStyle === "solid") {
    return "underscore";
  }
  return undefined;
}

/** The `style:text-properties` child of a style definition element. */
export function textPropertiesOf(styleEl: Element): Element | undefined {
  return firstChildElement(styleEl, STYLE_NAMESPACE, "text-properties");
}

/** The `style:paragraph-properties` child of a style definition element. */
export function paragraphPropertiesOf(styleEl: Element): Element | undefined {
  return firstChildElement(styleEl, STYLE_NAMESPACE, "paragraph-properties");
}
