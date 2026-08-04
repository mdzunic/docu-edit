// Maps OOXML run properties (`w:rPr`) and paragraph properties (`w:pPr`)
// onto the document model's `RunFormatting`/`ParagraphFormatting`. Shared by
// the style-catalog parser (styles.xml carries the same properties on each
// named style) and the document-body parser (direct formatting on `w:r`/`w:p`).
import type { ParagraphAlignment, ParagraphFormatting, RunFormatting } from "../documentModel";
import { HIGHLIGHT_NAMES } from "../odf/properties";
import { firstChildElement, wordAttr as attr, WORD_NAMESPACE } from "./xml";
import { halfPointsToPt, ooxmlColorToCss, twentiethsOfPointToPt, twipsToCm } from "./units";

const FALSE_VALUES = new Set(["0", "false", "off"]);

/** OOXML boolean toggle elements (`w:b`, `w:i`, `w:strike`, ...): present
 * with no `w:val` means true; `w:val="0"/"false"/"off"` means false. */
function parseBooleanToggle(el: Element | undefined): boolean | undefined {
  if (!el) {
    return undefined;
  }
  const val = attr(el, "val");
  if (val === null) {
    return true;
  }
  return !FALSE_VALUES.has(val.toLowerCase());
}

export function parseRunProperties(rPr: Element | undefined): RunFormatting {
  if (!rPr) {
    return {};
  }
  const formatting: RunFormatting = {};

  const bold = parseBooleanToggle(firstChildElement(rPr, WORD_NAMESPACE, "b"));
  if (bold !== undefined) {
    formatting.bold = bold;
  }

  const italic = parseBooleanToggle(firstChildElement(rPr, WORD_NAMESPACE, "i"));
  if (italic !== undefined) {
    formatting.italic = italic;
  }

  const underlineEl = firstChildElement(rPr, WORD_NAMESPACE, "u");
  if (underlineEl) {
    const val = attr(underlineEl, "val");
    formatting.underline = val !== null && val.toLowerCase() !== "none";
  }

  const strike = parseBooleanToggle(firstChildElement(rPr, WORD_NAMESPACE, "strike"));
  const doubleStrike = parseBooleanToggle(firstChildElement(rPr, WORD_NAMESPACE, "dstrike"));
  if (strike !== undefined || doubleStrike !== undefined) {
    formatting.strikethrough = strike === true || doubleStrike === true;
  }

  const rFonts = firstChildElement(rPr, WORD_NAMESPACE, "rFonts");
  if (rFonts) {
    const family = attr(rFonts, "ascii") ?? attr(rFonts, "hAnsi") ?? attr(rFonts, "cs");
    if (family) {
      formatting.fontFamily = family;
    }
  }

  const sz = firstChildElement(rPr, WORD_NAMESPACE, "sz");
  const szVal = sz ? attr(sz, "val") : null;
  if (szVal) {
    const halfPoints = Number(szVal);
    if (!Number.isNaN(halfPoints)) {
      formatting.fontSizePt = halfPointsToPt(halfPoints);
    }
  }

  const color = firstChildElement(rPr, WORD_NAMESPACE, "color");
  const colorVal = color ? attr(color, "val") : null;
  const css = ooxmlColorToCss(colorVal);
  if (css) {
    formatting.color = css;
  }

  const highlight = firstChildElement(rPr, WORD_NAMESPACE, "highlight");
  const highlightVal = highlight ? attr(highlight, "val") : null;
  if (highlightVal) {
    // "none" is an explicit off (SPEC-013): it must survive as `false` to
    // keep overriding a highlight inherited from the style chain.
    formatting.highlight = highlightVal.toLowerCase() === "none" ? false : highlightVal;
  }

  // Run shading (w:shd). w:highlight wins when both are present — Word
  // paints the highlight over the shading — and the shading is kept in the
  // shadingFill preservation slot; a shading alone IS the run's background,
  // so it becomes the highlight, normalised to the palette name when the
  // fill matches one (SPEC-013).
  const shd = firstChildElement(rPr, WORD_NAMESPACE, "shd");
  const fill = shd ? attr(shd, "fill") : null;
  if (fill && /^[0-9a-f]{6}$/i.test(fill)) {
    const hex = `#${fill.toLowerCase()}`;
    if (formatting.highlight !== undefined) {
      formatting.shadingFill = hex;
    } else {
      formatting.highlight = HIGHLIGHT_NAMES[hex] ?? hex;
    }
  }

  const vertAlign = firstChildElement(rPr, WORD_NAMESPACE, "vertAlign");
  const vertAlignVal = vertAlign ? attr(vertAlign, "val") : null;
  if (vertAlignVal === "subscript" || vertAlignVal === "superscript") {
    formatting.vertAlign = vertAlignVal;
  }

  return formatting;
}

const ALIGNMENT_MAP: Record<string, ParagraphAlignment> = {
  left: "left",
  start: "left",
  right: "right",
  end: "right",
  center: "center",
  both: "justify",
  distribute: "justify",
};

export function parseParagraphProperties(pPr: Element | undefined): ParagraphFormatting {
  if (!pPr) {
    return {};
  }
  const formatting: ParagraphFormatting = {};

  const jc = firstChildElement(pPr, WORD_NAMESPACE, "jc");
  const jcVal = jc ? attr(jc, "val") : null;
  if (jcVal) {
    const alignment = ALIGNMENT_MAP[jcVal.toLowerCase()];
    if (alignment) {
      formatting.alignment = alignment;
    }
  }

  const ind = firstChildElement(pPr, WORD_NAMESPACE, "ind");
  if (ind) {
    const left = attr(ind, "left") ?? attr(ind, "start");
    const right = attr(ind, "right") ?? attr(ind, "end");
    const firstLine = attr(ind, "firstLine");
    const hanging = attr(ind, "hanging");
    if (left) {
      formatting.indentLeftCm = twipsToCm(Number(left));
    }
    if (right) {
      formatting.indentRightCm = twipsToCm(Number(right));
    }
    if (firstLine) {
      formatting.indentFirstLineCm = twipsToCm(Number(firstLine));
    } else if (hanging) {
      formatting.indentFirstLineCm = -twipsToCm(Number(hanging));
    }
  }

  const spacing = firstChildElement(pPr, WORD_NAMESPACE, "spacing");
  if (spacing) {
    const before = attr(spacing, "before");
    const after = attr(spacing, "after");
    if (before) {
      formatting.spaceBeforePt = twentiethsOfPointToPt(Number(before));
    }
    if (after) {
      formatting.spaceAfterPt = twentiethsOfPointToPt(Number(after));
    }
    const line = attr(spacing, "line");
    if (line) {
      const lineRule = (attr(spacing, "lineRule") ?? "auto").toLowerCase();
      const lineValue = Number(line);
      if (lineRule === "auto") {
        formatting.lineSpacing = lineValue / 240;
      } else {
        formatting.lineSpacingPt = twentiethsOfPointToPt(lineValue);
      }
    }
  }

  return formatting;
}

/** `w:pPr/w:outlineLvl` is 0-based in OOXML; the document model's
 * `headingLevel` is 1-based (matching `<h1>`..`<h9>`). */
export function parseOutlineLevel(pPr: Element | undefined): number | undefined {
  if (!pPr) {
    return undefined;
  }
  const outlineLvl = firstChildElement(pPr, WORD_NAMESPACE, "outlineLvl");
  const val = outlineLvl ? attr(outlineLvl, "val") : null;
  if (val === null) {
    return undefined;
  }
  const level = Number(val);
  return Number.isNaN(level) ? undefined : level + 1;
}

export function mergeRunFormatting(base: RunFormatting, override: RunFormatting): RunFormatting {
  return { ...base, ...override };
}

export function mergeParagraphFormatting(
  base: ParagraphFormatting,
  override: ParagraphFormatting,
): ParagraphFormatting {
  return { ...base, ...override };
}
