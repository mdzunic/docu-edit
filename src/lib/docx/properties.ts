// Maps OOXML run properties (`w:rPr`) and paragraph properties (`w:pPr`)
// onto the document model's `RunFormatting`/`ParagraphFormatting`. Shared by
// the style-catalog parser (styles.xml carries the same properties on each
// named style) and the document-body parser (direct formatting on `w:r`/`w:p`).
import type {
  ParagraphAlignment,
  ParagraphFormatChange,
  ParagraphFormatting,
  RunFormatChange,
  RunFormatting,
  TabLeader,
  TabStop,
  TabStopAlignment,
} from "../documentModel";
import { HIGHLIGHT_NAMES } from "../core/colors";
import { wordAttr as attr, WORD_NAMESPACE } from "./xml";
import { firstChildElement } from "../core/xml";
import { ooxmlColorToCss } from "../core/colors";
import { halfPointsToPt, twentiethsOfPointToPt, twipsToCm } from "../core/units";

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

  // SPEC-035: language and "do not check spelling". w:lang carries three
  // attributes (val/eastAsia/bidi); only w:val is modelled, matching the
  // one-language-per-run scope of the spell checker.
  const lang = firstChildElement(rPr, WORD_NAMESPACE, "lang");
  const langVal = lang ? attr(lang, "val") : null;
  if (langVal) {
    formatting.lang = langVal;
  }

  const noProof = parseBooleanToggle(firstChildElement(rPr, WORD_NAMESPACE, "noProof"));
  if (noProof !== undefined) {
    formatting.noProof = noProof;
  }

  return formatting;
}

const UNKNOWN_TRACK_CHANGE_AUTHOR = "Unknown";

/** A `w:rPrChange`/`w:pPrChange`'s own author/date/id (SPEC-039) — both
 * extend `CT_TrackChange` just like `w:ins`/`w:del`, carrying the same
 * `w:id`/`w:author`/`w:date` attributes. Ids are derived from the source
 * `w:id`, not randomised, so re-parsing the same file is structurally
 * stable across a same-format round trip. */
function trackedChangeInfo(el: Element, kind: string) {
  const id = attr(el, "id") ?? `${Math.random()}`;
  return {
    changeId: `wtc-${kind}-${id}`,
    author: attr(el, "author") || UNKNOWN_TRACK_CHANGE_AUTHOR,
    date: attr(el, "date") || new Date(0).toISOString(),
  };
}

/** A run's pending formatting change (SPEC-039, `w:rPrChange`): the nested
 * `w:rPr` is the formatting the run had *before* the tracked change,
 * resolved the same way any other `w:rPr` is. Undefined when the run has
 * no pending formatting change. */
export function parseRunPropertiesChange(rPr: Element | undefined): RunFormatChange | undefined {
  if (!rPr) {
    return undefined;
  }
  const change = firstChildElement(rPr, WORD_NAMESPACE, "rPrChange");
  if (!change) {
    return undefined;
  }
  const previous = parseRunProperties(firstChildElement(change, WORD_NAMESPACE, "rPr"));
  return { ...trackedChangeInfo(change, "rPrChange"), previous };
}

/** A paragraph's pending formatting change (SPEC-039, `w:pPrChange`) — the
 * `w:pPrChange`/`w:pPr` counterpart of parseRunPropertiesChange. */
export function parseParagraphPropertiesChange(pPr: Element | undefined): ParagraphFormatChange | undefined {
  if (!pPr) {
    return undefined;
  }
  const change = firstChildElement(pPr, WORD_NAMESPACE, "pPrChange");
  if (!change) {
    return undefined;
  }
  const previous = parseParagraphProperties(firstChildElement(change, WORD_NAMESPACE, "pPr"));
  return { ...trackedChangeInfo(change, "pPrChange"), previous };
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

const TAB_ALIGNMENT_MAP: Record<string, TabStopAlignment> = {
  left: "left",
  start: "left",
  right: "right",
  end: "right",
  center: "center",
  decimal: "decimal",
};

const TAB_LEADER_MAP: Record<string, TabLeader> = {
  dot: "dot",
  middledot: "dot",
  hyphen: "hyphen",
  underscore: "underscore",
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

  // `w:ind`'s left/right (start/end under the strict schema names) are the
  // LOGICAL start/end edges — with w:bidi they paint on the right/left — so
  // they map onto the model's logical indents unchanged (SPEC-014).
  const ind = firstChildElement(pPr, WORD_NAMESPACE, "ind");
  if (ind) {
    const start = attr(ind, "left") ?? attr(ind, "start");
    const end = attr(ind, "right") ?? attr(ind, "end");
    const firstLine = attr(ind, "firstLine");
    const hanging = attr(ind, "hanging");
    if (start) {
      formatting.indentStartCm = twipsToCm(Number(start));
    }
    if (end) {
      formatting.indentEndCm = twipsToCm(Number(end));
    }
    if (firstLine) {
      formatting.indentFirstLineCm = twipsToCm(Number(firstLine));
    } else if (hanging) {
      formatting.indentFirstLineCm = -twipsToCm(Number(hanging));
    }
  }

  // Custom tab stops (SPEC-032 AC-19..25). `clear`/`bar`/`num` stops have
  // no model equivalent and are skipped; `start`/`end` are the strict-schema
  // names for left/right.
  const tabsEl = firstChildElement(pPr, WORD_NAMESPACE, "tabs");
  if (tabsEl) {
    const stops: TabStop[] = [];
    for (const tabEl of Array.from(tabsEl.children)) {
      if (tabEl.namespaceURI !== WORD_NAMESPACE || tabEl.localName !== "tab") {
        continue;
      }
      const alignment = TAB_ALIGNMENT_MAP[(attr(tabEl, "val") ?? "").toLowerCase()];
      const pos = Number(attr(tabEl, "pos"));
      if (!alignment || !Number.isFinite(pos) || pos < 0) {
        continue;
      }
      const stop: TabStop = { positionCm: twipsToCm(pos), alignment };
      const leader = TAB_LEADER_MAP[(attr(tabEl, "leader") ?? "").toLowerCase()];
      if (leader) {
        stop.leader = leader;
      }
      stops.push(stop);
    }
    if (stops.length > 0) {
      formatting.tabStops = stops.sort((a, b) => a.positionCm - b.positionCm);
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
        // "exact" is the marked rule; "atLeast" is the model's absent
        // default (see ParagraphFormatting.lineSpacingRule).
        if (lineRule === "exact") {
          formatting.lineSpacingRule = "exact";
        }
      }
    }
  }

  // w:bidi is an OOXML boolean toggle. An explicit off ("ltr") is recorded
  // here so it can override a style chain's RTL during merging; the reader
  // normalises "ltr" away before the model is assembled — the model itself
  // only ever holds "rtl" (SPEC-014, see ParagraphDirection).
  const bidi = parseBooleanToggle(firstChildElement(pPr, WORD_NAMESPACE, "bidi"));
  if (bidi !== undefined) {
    formatting.direction = bidi ? "rtl" : "ltr";
  }

  // The SPEC-028 pagination flags, all ST_OnOff toggles. Like w:bidi above,
  // the non-canonical value (an explicit off on the keeps, an explicit on
  // for widow control) is recorded here so it can override a style chain
  // during merging; the reader normalises it away before the model is
  // assembled — the model only ever holds keepNext/keepLines/
  // pageBreakBefore as true and widowControl as false (absent means on,
  // both file formats' default; see ParagraphFormatting).
  const keepNext = parseBooleanToggle(firstChildElement(pPr, WORD_NAMESPACE, "keepNext"));
  if (keepNext !== undefined) {
    formatting.keepNext = keepNext;
  }
  const keepLines = parseBooleanToggle(firstChildElement(pPr, WORD_NAMESPACE, "keepLines"));
  if (keepLines !== undefined) {
    formatting.keepLines = keepLines;
  }
  const pageBreakBefore = parseBooleanToggle(firstChildElement(pPr, WORD_NAMESPACE, "pageBreakBefore"));
  if (pageBreakBefore !== undefined) {
    formatting.pageBreakBefore = pageBreakBefore;
  }
  const widowControl = parseBooleanToggle(firstChildElement(pPr, WORD_NAMESPACE, "widowControl"));
  if (widowControl !== undefined) {
    formatting.widowControl = widowControl;
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
