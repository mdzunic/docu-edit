// Validates a written `.docx` package against the OPC packaging rules and
// the WordprocessingML (ECMA-376 Transitional) schema constraints for the
// parts this project's writer emits (AC-19, AC-22). A full XSD validator is
// off-limits under the no-new-dependencies constraint, so the relevant
// schema rules are encoded directly: required parts and content types,
// root elements and namespaces, the child sets and *sequence order* of
// w:body/w:p/w:pPr/w:r/w:rPr/w:sectPr (out-of-order children are precisely
// what makes Word show its repair prompt), and attribute value spaces
// (measurements, enums, hex colours, W3CDTF timestamps).
//
// Returns a list of violations — empty means the package is valid. The
// checks are deliberately strict allowlists over what the writer emits: an
// element the writer should never produce is reported rather than ignored,
// so a writer regression cannot slip through as "unknown but tolerated".
import { unzip, ZipError } from "../core/zip";
import { wordAttr, WORD_NAMESPACE } from "./xml";
import { parseXml, XmlParseError } from "../core/xml";

const CONTENT_TYPES_NAMESPACE = "http://schemas.openxmlformats.org/package/2006/content-types";
const RELATIONSHIPS_NAMESPACE = "http://schemas.openxmlformats.org/package/2006/relationships";
const CORE_PROPERTIES_NAMESPACE =
  "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
const DC_NAMESPACE = "http://purl.org/dc/elements/1.1/";
const DCTERMS_NAMESPACE = "http://purl.org/dc/terms/";
const XSI_NAMESPACE = "http://www.w3.org/2001/XMLSchema-instance";

const OFFICE_DOCUMENT_RELATIONSHIP =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
const CORE_PROPERTIES_RELATIONSHIP =
  "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";

const REQUIRED_PARTS = [
  "[Content_Types].xml",
  "_rels/.rels",
  "word/document.xml",
  "word/styles.xml",
  "docProps/core.xml",
];

// ST_Jc (the subset in use plus the strict-mode aliases readers accept).
const JC_VALUES = new Set(["left", "center", "right", "both", "start", "end", "distribute"]);
// ST_LineSpacingRule.
const LINE_RULE_VALUES = new Set(["auto", "exact", "atLeast"]);
// ST_Underline.
const UNDERLINE_VALUES = new Set([
  "single", "words", "double", "thick", "dotted", "dottedHeavy", "dash",
  "dashedHeavy", "dashLong", "dashLongHeavy", "dotDash", "dashDotHeavy",
  "dotDotDash", "dashDotDotHeavy", "wave", "wavyHeavy", "wavyDouble", "none",
]);
// ST_HighlightColor.
const HIGHLIGHT_VALUES = new Set([
  "black", "blue", "cyan", "darkBlue", "darkCyan", "darkGray", "darkGreen",
  "darkMagenta", "darkRed", "darkYellow", "green", "lightGray", "magenta",
  "none", "red", "white", "yellow",
]);

// ST_NumberFormat (the subset the SPEC-018 writer emits).
const NUM_FMT_VALUES = new Set([
  "bullet", "decimal", "lowerLetter", "upperLetter", "lowerRoman", "upperRoman",
]);

// Sequence order of the CT_PPr / CT_RPr / CT_SectPr children the writer
// emits, per the ECMA-376 Transitional schema. Order within each list is
// the schema's required document order.
const PPR_SEQUENCE = [
  "pStyle", "keepNext", "keepLines", "pageBreakBefore", "widowControl",
  "numPr", "tabs", "bidi", "spacing", "ind", "jc", "outlineLvl", "sectPr",
];
// ST_TabJc (the subset the writer emits: custom tab stops, SPEC-032; TOC
// entry dot leaders, SPEC-033).
const TAB_JC_VALUES = new Set(["left", "right", "center", "decimal"]);
// ST_TabTlc.
const TAB_LEADER_CHAR_VALUES = new Set(["none", "dot", "hyphen", "underscore"]);
// ST_FldCharType (SPEC-033's TOC/PAGEREF complex fields).
const FLD_CHAR_TYPES = new Set(["begin", "separate", "end"]);
// CT_RPr's schema order runs "… strike, …, noProof, …, color, spacing, …,
// sz, …, highlight, u, effect, bdr, shd, fitText, vertAlign, rtl, cs, em,
// lang, …" — noProof precedes color, lang follows vertAlign (SPEC-035; the
// rtl/cs/em slots between them are never emitted by this writer).
const RPR_SEQUENCE = [
  "rFonts", "b", "i", "strike", "noProof", "color", "sz", "highlight", "u", "shd", "vertAlign", "lang",
];
// CT_SectPr puts w:type before the page geometry (SPEC-028 AC-31); w:cols
// follows w:pgMar in the schema (SPEC-030).
// CT_SectPr's child sequence per the ECMA-376 Transitional schema. The
// header/footer references (SPEC-029) lead it and each may appear once per
// ST_HdrFtr value (default/first/even), so they are repeatable.
const SECTPR_SEQUENCE = [
  "headerReference", "footerReference", "footnotePr", "endnotePr", "type",
  "pgSz", "pgMar", "paperSrc", "pgBorders", "lnNumType", "pgNumType", "cols",
  "formProt", "vAlign", "noEndnote", "titlePg", "textDirection", "bidi",
  "rtlGutter", "docGrid", "printerSettings", "sectPrChange",
];
const SECTPR_REPEATABLE = new Set(["headerReference", "footerReference"]);
// CT_Style's child sequence per the ECMA-376 Transitional schema. The writer
// itself emits only name/basedOn/next/pPr/rPr, but a same-format save
// PRESERVES the source's styles.xml and patches it in place (see
// patchStylesXml), so any producer's style children land here — w:uiPriority
// and w:qFormat come back from LibreOffice on every round trip.
const STYLE_SEQUENCE = [
  "name", "aliases", "basedOn", "next", "link", "autoRedefine", "hidden",
  "uiPriority", "semiHidden", "unhideWhenUsed", "qFormat", "locked", "personal",
  "personalCompose", "personalReply", "rsid", "pPr", "rPr", "tblPr", "trPr",
  "tcPr", "tblStylePr",
];

// ST_StyleType. The writer emits character styles (SPEC-042's Run.styleId)
// and preserves whatever else the source defined.
const STYLE_TYPES = new Set(["paragraph", "character", "table", "numbering"]);

// EG_BlockLevelElts members the writer emits beyond w:p/w:tbl/w:sectPr,
// which are handled by name above. w:sdt is a preserved content control and
// the bookmark marks are SPEC-025's — all re-emitted verbatim from opaque
// blocks (SPEC-008), so the body legitimately holds them.
const BODY_CHILDREN = new Set(["sdt", "bookmarkStart", "bookmarkEnd", "customXml"]);

// EG_RunInnerContent members the writer emits. w:drawing is every image,
// shape and chart (SPEC-024/051/052); the note/comment reference marks are
// SPEC-031/038; w:sym is SPEC-026's special characters.
const RUN_CHILDREN = new Set([
  "t", "tab", "br", "drawing", "pict", "sym", "softHyphen", "noBreakHyphen",
  "footnoteReference", "endnoteReference", "commentReference",
  "footnoteRef", "endnoteRef", "separator", "continuationSeparator",
  "lastRenderedPageBreak", "delText", "object",
]);

const INTEGER_PATTERN = /^-?\d+$/;
const UNSIGNED_PATTERN = /^\d+$/;
const HEX_COLOR_PATTERN = /^([0-9A-Fa-f]{6}|auto)$/;
const W3CDTF_PATTERN =
  /^\d{4}(-\d{2}(-\d{2}(T\d{2}:\d{2}(:\d{2}(\.\d+)?)?(Z|[+-]\d{2}:\d{2}))?)?)?$/;

function isWordElement(el: Element): boolean {
  return el.namespaceURI === WORD_NAMESPACE;
}

/** Checks `children` against an allowlisted schema sequence: every child
 * must be a WordprocessingML element from `sequence`, appearing in the
 * sequence's order (each element type at most once — matching what the
 * writer emits). */
function checkSequence(
  children: Element[],
  sequence: string[],
  context: string,
  violations: string[],
  /** Elements the schema allows more than once in a row (a section's three
   * header references, one per ST_HdrFtr value). */
  repeatable?: ReadonlySet<string>,
): void {
  let lastIndex = -1;
  for (const child of children) {
    if (!isWordElement(child)) {
      violations.push(`${context}: unexpected element ${child.nodeName}`);
      continue;
    }
    const index = sequence.indexOf(child.localName);
    if (index === -1) {
      violations.push(`${context}: unexpected element w:${child.localName}`);
    } else if (index < lastIndex) {
      violations.push(
        `${context}: w:${child.localName} is out of schema sequence order (must precede w:${sequence[lastIndex]})`,
      );
    } else if (index === lastIndex && !repeatable?.has(child.localName)) {
      violations.push(`${context}: duplicate element w:${child.localName}`);
    } else {
      lastIndex = index;
    }
  }
}

function checkAttr(
  el: Element,
  name: string,
  test: (value: string) => boolean,
  description: string,
  context: string,
  violations: string[],
  required = true,
): void {
  const value = wordAttr(el, name);
  if (value === null) {
    if (required) {
      violations.push(`${context}: w:${el.localName} is missing required attribute w:${name}`);
    }
    return;
  }
  if (!test(value)) {
    violations.push(
      `${context}: w:${el.localName}/@w:${name}="${value}" is not ${description}`,
    );
  }
}

function checkRunProperties(rPr: Element, context: string, violations: string[]): void {
  checkSequence(Array.from(rPr.children), RPR_SEQUENCE, context, violations);
  for (const child of Array.from(rPr.children)) {
    if (!isWordElement(child)) {
      continue;
    }
    switch (child.localName) {
      case "sz":
        checkAttr(child, "val", (v) => UNSIGNED_PATTERN.test(v) && Number(v) > 0,
          "a positive half-point measurement", context, violations);
        break;
      case "color":
        checkAttr(child, "val", (v) => HEX_COLOR_PATTERN.test(v),
          "a six-digit hex colour or 'auto'", context, violations);
        break;
      case "highlight":
        checkAttr(child, "val", (v) => HIGHLIGHT_VALUES.has(v),
          "an ST_HighlightColor value", context, violations);
        break;
      case "u":
        checkAttr(child, "val", (v) => UNDERLINE_VALUES.has(v),
          "an ST_Underline value", context, violations);
        break;
      case "shd":
        // The writer only emits the clear-pattern shading fallback
        // (SPEC-013): a solid fill colour with no pattern overlay.
        checkAttr(child, "val", (v) => v === "clear", 'the shading pattern "clear"',
          context, violations);
        checkAttr(child, "color", (v) => HEX_COLOR_PATTERN.test(v),
          "a six-digit hex colour or 'auto'", context, violations, false);
        checkAttr(child, "fill", (v) => HEX_COLOR_PATTERN.test(v),
          "a six-digit hex colour or 'auto'", context, violations);
        break;
      case "vertAlign":
        checkAttr(child, "val", (v) => ["baseline", "subscript", "superscript"].includes(v),
          "an ST_VerticalAlignRun value", context, violations);
        break;
      case "lang":
        checkAttr(child, "val", (v) => v.length > 0, "a non-empty language tag", context, violations);
        break;
      case "rFonts":
        break;
      default:
        // b/i/strike/noProof: w:val, when present, must be an ST_OnOff value.
        checkAttr(child, "val", (v) => ["0", "1", "true", "false", "on", "off"].includes(v),
          "an ST_OnOff value", context, violations, false);
    }
  }
}

function checkParagraphProperties(pPr: Element, context: string, violations: string[]): void {
  checkSequence(Array.from(pPr.children), PPR_SEQUENCE, context, violations);
  for (const child of Array.from(pPr.children)) {
    if (!isWordElement(child)) {
      continue;
    }
    switch (child.localName) {
      case "pStyle":
        checkAttr(child, "val", (v) => v.length > 0, "a non-empty style id", context, violations);
        break;
      case "keepNext":
      case "keepLines":
      case "pageBreakBefore":
      case "widowControl":
        // The SPEC-028 pagination toggles: w:val, when present, must be an
        // ST_OnOff value (same rule as the rPr b/i/strike toggles).
        checkAttr(child, "val", (v) => ["0", "1", "true", "false", "on", "off"].includes(v),
          "an ST_OnOff value", context, violations, false);
        break;
      case "sectPr":
        // A paragraph-level sectPr ends a non-final section (SPEC-028
        // AC-31) — same shape as the body-level one.
        checkSectionProperties(child, `${context}/w:sectPr`, violations);
        break;
      case "numPr":
        checkSequence(Array.from(child.children), ["ilvl", "numId"], `${context}/w:numPr`, violations);
        for (const numChild of Array.from(child.children)) {
          if (isWordElement(numChild) && numChild.localName === "ilvl") {
            checkAttr(numChild, "val", (v) => UNSIGNED_PATTERN.test(v) && Number(v) <= 8,
              "a list level between 0 and 8", `${context}/w:numPr`, violations);
          } else if (isWordElement(numChild) && numChild.localName === "numId") {
            checkAttr(numChild, "val", (v) => UNSIGNED_PATTERN.test(v),
              "a non-negative numbering id", `${context}/w:numPr`, violations);
          }
        }
        break;
      case "spacing":
        checkAttr(child, "before", (v) => UNSIGNED_PATTERN.test(v),
          "a non-negative twentieths-of-a-point measurement", context, violations, false);
        checkAttr(child, "after", (v) => UNSIGNED_PATTERN.test(v),
          "a non-negative twentieths-of-a-point measurement", context, violations, false);
        checkAttr(child, "line", (v) => INTEGER_PATTERN.test(v),
          "a whole-number line measurement", context, violations, false);
        checkAttr(child, "lineRule", (v) => LINE_RULE_VALUES.has(v),
          "an ST_LineSpacingRule value", context, violations, false);
        break;
      case "ind":
        for (const attr of ["left", "right", "firstLine", "hanging"]) {
          checkAttr(child, attr, (v) => INTEGER_PATTERN.test(v),
            "a whole-number twip measurement", context, violations, false);
        }
        break;
      case "jc":
        checkAttr(child, "val", (v) => JC_VALUES.has(v), "an ST_Jc value", context, violations);
        break;
      case "outlineLvl":
        checkAttr(child, "val", (v) => UNSIGNED_PATTERN.test(v) && Number(v) <= 9,
          "an outline level between 0 and 9", context, violations);
        break;
      case "tabs":
        for (const tab of Array.from(child.children)) {
          if (!isWordElement(tab) || tab.localName !== "tab") {
            continue;
          }
          checkAttr(tab, "val", (v) => TAB_JC_VALUES.has(v), "an ST_TabJc value", context, violations);
          checkAttr(tab, "pos", (v) => INTEGER_PATTERN.test(v),
            "a whole-number twip measurement", context, violations);
          checkAttr(tab, "leader", (v) => TAB_LEADER_CHAR_VALUES.has(v),
            "an ST_TabTlc value", context, violations, false);
        }
        break;
    }
  }
}

function checkRun(run: Element, context: string, violations: string[]): void {
  const children = Array.from(run.children);
  children.forEach((child, index) => {
    if (!isWordElement(child)) {
      violations.push(`${context}: unexpected element ${child.nodeName}`);
      return;
    }
    if (child.localName === "rPr") {
      if (index !== 0) {
        violations.push(`${context}: w:rPr must be the first child of w:r`);
      }
      checkRunProperties(child, `${context}/w:rPr`, violations);
    } else if (child.localName === "fldChar") {
      checkAttr(child, "fldCharType", (v) => FLD_CHAR_TYPES.has(v),
        "an ST_FldCharType value", context, violations);
    } else if (child.localName === "instrText") {
      // Free-form field instruction text (SPEC-033's TOC/PAGEREF fields) —
      // no attribute constraints beyond being a WordprocessingML element.
    } else if (!RUN_CHILDREN.has(child.localName)) {
      violations.push(`${context}: unexpected element w:${child.localName}`);
    }
  });
}

// The r: attribute namespace (w:hyperlink/@r:id) — SPEC-025.
const R_ATTRIBUTE_NAMESPACE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/** A `w:hyperlink` (SPEC-025): must carry exactly one target — an r:id
 * (external relationship) or a w:anchor (internal bookmark/heading) — and
 * contains runs and bookmark elements like the paragraph itself. */
function checkHyperlink(hyperlink: Element, context: string, violations: string[]): void {
  const relId = hyperlink.getAttributeNS(R_ATTRIBUTE_NAMESPACE, "id");
  const anchor = wordAttr(hyperlink, "anchor");
  if ((relId === null || relId.length === 0) === (anchor === null || anchor.length === 0)) {
    violations.push(`${context}: w:hyperlink must carry exactly one of r:id or w:anchor`);
  }
  for (const child of Array.from(hyperlink.children)) {
    if (!isWordElement(child)) {
      violations.push(`${context}: unexpected element ${child.nodeName}`);
    } else if (child.localName === "r") {
      checkRun(child, `${context}/w:r`, violations);
    } else if (child.localName === "bookmarkStart" || child.localName === "bookmarkEnd") {
      checkBookmarkElement(child, context, violations);
    } else {
      violations.push(`${context}: unexpected element w:${child.localName}`);
    }
  }
}

/** A `w:bookmarkStart`/`w:bookmarkEnd` (SPEC-025): both need a numeric
 * w:id; the start additionally needs a non-empty, space-free w:name. */
function checkBookmarkElement(el: Element, context: string, violations: string[]): void {
  checkAttr(el, "id", (v) => UNSIGNED_PATTERN.test(v), "a non-negative bookmark id", context, violations);
  if (el.localName === "bookmarkStart") {
    checkAttr(el, "name", (v) => v.length > 0 && !/\s/.test(v),
      "a non-empty bookmark name without spaces", context, violations);
  }
}

function checkParagraph(paragraph: Element, context: string, violations: string[]): void {
  Array.from(paragraph.children).forEach((child, index) => {
    if (!isWordElement(child)) {
      violations.push(`${context}: unexpected element ${child.nodeName}`);
      return;
    }
    if (child.localName === "pPr") {
      if (index !== 0) {
        violations.push(`${context}: w:pPr must be the first child of w:p`);
      }
      checkParagraphProperties(child, `${context}/w:pPr`, violations);
    } else if (child.localName === "r") {
      checkRun(child, `${context}/w:r`, violations);
    } else if (child.localName === "hyperlink") {
      checkHyperlink(child, `${context}/w:hyperlink`, violations);
    } else if (child.localName === "bookmarkStart" || child.localName === "bookmarkEnd") {
      checkBookmarkElement(child, context, violations);
    } else {
      violations.push(`${context}: unexpected element w:${child.localName}`);
    }
  });
}

function checkSectionProperties(sectPr: Element, context: string, violations: string[]): void {
  checkSequence(Array.from(sectPr.children), SECTPR_SEQUENCE, context, violations, SECTPR_REPEATABLE);
  for (const child of Array.from(sectPr.children)) {
    if (!isWordElement(child)) {
      continue;
    }
    if (child.localName === "type") {
      // ST_SectionMark (SPEC-028): the writer only ever emits "continuous",
      // but the full value space is legal in preserved source sectPrs.
      checkAttr(child, "val",
        (v) => ["continuous", "nextPage", "nextColumn", "evenPage", "oddPage"].includes(v),
        "an ST_SectionMark value", context, violations, false);
    } else if (child.localName === "pgSz") {
      checkAttr(child, "w", (v) => UNSIGNED_PATTERN.test(v) && Number(v) > 0,
        "a positive page width in twips", context, violations);
      checkAttr(child, "h", (v) => UNSIGNED_PATTERN.test(v) && Number(v) > 0,
        "a positive page height in twips", context, violations);
      checkAttr(child, "orient", (v) => v === "portrait" || v === "landscape",
        "an ST_PageOrientation value", context, violations, false);
    } else if (child.localName === "pgMar") {
      // CT_PageMar makes all seven attributes required.
      for (const attr of ["top", "right", "bottom", "left"]) {
        checkAttr(child, attr, (v) => INTEGER_PATTERN.test(v),
          "a whole-number twip measurement", context, violations);
      }
      for (const attr of ["header", "footer", "gutter"]) {
        checkAttr(child, attr, (v) => UNSIGNED_PATTERN.test(v),
          "a non-negative twip measurement", context, violations);
      }
    } else if (child.localName === "cols") {
      // CT_Columns (SPEC-030): w:num a positive column count, w:space a
      // non-negative gutter, w:sep an on/off flag; each w:col child a
      // positive width. All optional in the schema, so absent is legal.
      checkAttr(child, "num", (v) => UNSIGNED_PATTERN.test(v) && Number(v) > 0,
        "a positive column count", context, violations, false);
      checkAttr(child, "space", (v) => UNSIGNED_PATTERN.test(v),
        "a non-negative twip gutter", context, violations, false);
      checkAttr(child, "sep", (v) => ["0", "1", "true", "false", "on", "off"].includes(v),
        "an ST_OnOff value", context, violations, false);
      for (const col of Array.from(child.children)) {
        if (isWordElement(col) && col.localName === "col") {
          checkAttr(col, "w", (v) => UNSIGNED_PATTERN.test(v) && Number(v) > 0,
            "a positive column width in twips", context, violations);
          checkAttr(col, "space", (v) => UNSIGNED_PATTERN.test(v),
            "a non-negative twip gutter", context, violations, false);
        }
      }
    }
  }
}

/** Structural check of a w:tbl (SPEC-022): child order (w:tblPr?,
 * w:tblGrid, w:tr+), a positive-width w:gridCol per column, and a regular
 * grid — every row holds exactly one w:tc per grid column, each with block
 * content. Property bags (w:tblPr/w:trPr/w:tcPr) are not deep-checked: on a
 * same-format save they carry the source document's own XML verbatim. */
function checkTable(table: Element, context: string, violations: string[]): void {
  const children = Array.from(table.children);
  let columnCount = 0;
  let sawGrid = false;
  let sawRow = false;
  children.forEach((child, index) => {
    if (!isWordElement(child)) {
      violations.push(`${context}: unexpected element ${child.nodeName}`);
      return;
    }
    if (child.localName === "tblPr") {
      if (index !== 0) {
        violations.push(`${context}: w:tblPr must be the first child of w:tbl`);
      }
    } else if (child.localName === "tblGrid") {
      if (sawGrid) {
        violations.push(`${context}: duplicate w:tblGrid`);
      }
      if (sawRow) {
        violations.push(`${context}: w:tblGrid must precede the table's rows`);
      }
      sawGrid = true;
      for (const col of Array.from(child.children)) {
        if (!isWordElement(col) || col.localName !== "gridCol") {
          violations.push(`${context}: unexpected element in w:tblGrid`);
          continue;
        }
        columnCount += 1;
        checkAttr(col, "w", (v) => UNSIGNED_PATTERN.test(v) && Number(v) > 0,
          "a positive column width in twips", context, violations);
      }
      if (columnCount === 0) {
        violations.push(`${context}: w:tblGrid must declare at least one w:gridCol`);
      }
    } else if (child.localName === "tr") {
      sawRow = true;
      if (!sawGrid) {
        violations.push(`${context}: w:tr before w:tblGrid`);
      }
      const rowContext = `${context}/w:tr`;
      let cellCount = 0;
      let columnsCovered = 0;
      Array.from(child.children).forEach((rowChild, rowIndex) => {
        if (!isWordElement(rowChild)) {
          violations.push(`${rowContext}: unexpected element ${rowChild.nodeName}`);
          return;
        }
        if (rowChild.localName === "trPr") {
          if (rowIndex !== 0) {
            violations.push(`${rowContext}: w:trPr must be the first child of w:tr`);
          }
        } else if (rowChild.localName === "tc") {
          cellCount += 1;
          const cellContext = `${rowContext}/w:tc[${cellCount}]`;
          let sawBlock = false;
          let gridSpan = 1;
          Array.from(rowChild.children).forEach((cellChild, cellIndex) => {
            if (!isWordElement(cellChild)) {
              violations.push(`${cellContext}: unexpected element ${cellChild.nodeName}`);
              return;
            }
            if (cellChild.localName === "tcPr") {
              if (cellIndex !== 0) {
                violations.push(`${cellContext}: w:tcPr must be the first child of w:tc`);
              }
              // SPEC-023: a horizontally merged cell's w:gridSpan means this
              // one w:tc stands in for that many w:tblGrid columns.
              const gridSpanEl = Array.from(cellChild.children).find(
                (el) => isWordElement(el) && el.localName === "gridSpan",
              );
              if (gridSpanEl) {
                const val = Number(gridSpanEl.getAttributeNS(WORD_NAMESPACE, "val"));
                if (Number.isInteger(val) && val > 0) {
                  gridSpan = val;
                }
              }
            } else if (cellChild.localName === "p") {
              sawBlock = true;
              checkParagraph(cellChild, `${cellContext}/w:p`, violations);
            } else {
              violations.push(`${cellContext}: unexpected element w:${cellChild.localName}`);
            }
          });
          if (!sawBlock) {
            violations.push(`${cellContext}: w:tc must contain at least one w:p`);
          }
          columnsCovered += gridSpan;
        } else {
          violations.push(`${rowContext}: unexpected element w:${rowChild.localName}`);
        }
      });
      if (columnCount > 0 && columnsCovered !== columnCount) {
        violations.push(
          `${rowContext}: cells cover ${columnsCovered} grid columns but w:tblGrid declares ${columnCount}`,
        );
      }
    } else {
      violations.push(`${context}: unexpected element w:${child.localName} in w:tbl`);
    }
  });
  if (!sawGrid) {
    violations.push(`${context}: w:tbl is missing w:tblGrid`);
  }
  if (!sawRow) {
    violations.push(`${context}: w:tbl must contain at least one w:tr`);
  }
}

function checkDocumentPart(doc: Document, violations: string[]): void {
  const root = doc.documentElement;
  if (root.namespaceURI !== WORD_NAMESPACE || root.localName !== "document") {
    violations.push("word/document.xml: root element is not w:document");
    return;
  }
  const children = Array.from(root.children);
  if (children.length !== 1 || !isWordElement(children[0]) || children[0].localName !== "body") {
    violations.push("word/document.xml: w:document must contain exactly one w:body");
    return;
  }
  const bodyChildren = Array.from(children[0].children);
  if (!bodyChildren.some((child) => isWordElement(child) && child.localName === "p")) {
    violations.push("word/document.xml: w:body must contain at least one w:p");
  }
  bodyChildren.forEach((child, index) => {
    if (!isWordElement(child)) {
      violations.push(`word/document.xml: unexpected element ${child.nodeName} in w:body`);
      return;
    }
    if (child.localName === "p") {
      checkParagraph(child, `word/document.xml: w:body/w:p[${index + 1}]`, violations);
    } else if (child.localName === "tbl") {
      checkTable(child, `word/document.xml: w:body/w:tbl[${index + 1}]`, violations);
    } else if (child.localName === "sectPr") {
      if (index !== bodyChildren.length - 1) {
        violations.push("word/document.xml: body-level w:sectPr must be the last child of w:body");
      }
      checkSectionProperties(child, "word/document.xml: w:body/w:sectPr", violations);
    } else if (!BODY_CHILDREN.has(child.localName)) {
      violations.push(`word/document.xml: unexpected element w:${child.localName} in w:body`);
    }
  });
}

function checkStylesPart(doc: Document, violations: string[]): void {
  const root = doc.documentElement;
  if (root.namespaceURI !== WORD_NAMESPACE || root.localName !== "styles") {
    violations.push("word/styles.xml: root element is not w:styles");
    return;
  }
  let sawStyle = false;
  let defaultParagraphStyles = 0;
  for (const child of Array.from(root.children)) {
    if (!isWordElement(child)) {
      violations.push(`word/styles.xml: unexpected element ${child.nodeName}`);
      continue;
    }
    if (child.localName === "docDefaults") {
      if (sawStyle) {
        violations.push("word/styles.xml: w:docDefaults must precede all w:style elements");
      }
    } else if (child.localName === "style") {
      sawStyle = true;
      const context = "word/styles.xml: w:style";
      const styleId = wordAttr(child, "styleId");
      if (!styleId) {
        violations.push(`${context} is missing required attribute w:styleId`);
      }
      const styleType = wordAttr(child, "type");
      if (styleType !== undefined && !STYLE_TYPES.has(styleType)) {
        violations.push(
          `${context}[@w:styleId="${styleId}"] has w:type="${styleType}", not an ST_StyleType value`,
        );
      }
      // One default per FAMILY is legal; only the paragraph family's default
      // is the one WordprocessingML requires exactly one of.
      if (wordAttr(child, "default") === "1" && (styleType ?? "paragraph") === "paragraph") {
        defaultParagraphStyles++;
      }
      checkSequence(
        Array.from(child.children),
        STYLE_SEQUENCE,
        `${context}[@w:styleId="${styleId}"]`,
        violations,
      );
    } else {
      violations.push(`word/styles.xml: unexpected element w:${child.localName}`);
    }
  }
  if (defaultParagraphStyles !== 1) {
    violations.push(
      `word/styles.xml: expected exactly one default paragraph style, found ${defaultParagraphStyles}`,
    );
  }
}

// CT_Lvl children the SPEC-018 writer emits, in schema sequence order.
const LVL_SEQUENCE = ["start", "numFmt", "lvlText", "lvlJc", "pPr"];

/** The numbering part (SPEC-018 AC-18): every abstractNum precedes every
 * num, each num references an abstractNum that exists, and levels carry the
 * value spaces the writer emits. Checked only when the part is present. */
function checkNumberingPart(doc: Document, violations: string[]): void {
  const context = "word/numbering.xml";
  const root = doc.documentElement;
  if (root.namespaceURI !== WORD_NAMESPACE || root.localName !== "numbering") {
    violations.push(`${context}: root element is not w:numbering`);
    return;
  }
  const abstractIds = new Set<string>();
  let sawNum = false;
  for (const child of Array.from(root.children)) {
    if (!isWordElement(child)) {
      violations.push(`${context}: unexpected element ${child.nodeName}`);
      continue;
    }
    if (child.localName === "abstractNum") {
      if (sawNum) {
        violations.push(`${context}: w:abstractNum must precede every w:num`);
      }
      const id = wordAttr(child, "abstractNumId");
      if (id === null) {
        violations.push(`${context}: w:abstractNum is missing w:abstractNumId`);
      } else {
        abstractIds.add(id);
      }
      for (const lvl of Array.from(child.children)) {
        if (!isWordElement(lvl) || lvl.localName !== "lvl") {
          continue;
        }
        const lvlContext = `${context}: w:abstractNum[@w:abstractNumId="${id}"]/w:lvl`;
        checkAttr(lvl, "ilvl", (v) => UNSIGNED_PATTERN.test(v) && Number(v) <= 8,
          "a list level between 0 and 8", lvlContext, violations);
        checkSequence(Array.from(lvl.children), LVL_SEQUENCE, lvlContext, violations);
        for (const lvlChild of Array.from(lvl.children)) {
          if (!isWordElement(lvlChild)) {
            continue;
          }
          if (lvlChild.localName === "start") {
            checkAttr(lvlChild, "val", (v) => UNSIGNED_PATTERN.test(v),
              "a non-negative start value", lvlContext, violations);
          } else if (lvlChild.localName === "numFmt") {
            checkAttr(lvlChild, "val", (v) => NUM_FMT_VALUES.has(v),
              "an ST_NumberFormat value the writer emits", lvlContext, violations);
          } else if (lvlChild.localName === "lvlText") {
            checkAttr(lvlChild, "val", (v) => v.length > 0, "a non-empty level text",
              lvlContext, violations);
          }
        }
      }
    } else if (child.localName === "num") {
      sawNum = true;
      checkAttr(child, "numId", (v) => UNSIGNED_PATTERN.test(v) && Number(v) > 0,
        "a positive numbering id", context, violations);
      const ref = Array.from(child.children).find(
        (el) => isWordElement(el) && el.localName === "abstractNumId",
      );
      if (!ref) {
        violations.push(`${context}: w:num is missing its w:abstractNumId child`);
      } else {
        checkAttr(ref, "val", (v) => abstractIds.has(v),
          "the id of a declared w:abstractNum", context, violations);
      }
    } else if (child.localName !== "numPicBullet") {
      violations.push(`${context}: unexpected element w:${child.localName}`);
    }
  }
}

function checkCorePropertiesPart(doc: Document, violations: string[]): void {
  const root = doc.documentElement;
  if (root.namespaceURI !== CORE_PROPERTIES_NAMESPACE || root.localName !== "coreProperties") {
    violations.push("docProps/core.xml: root element is not cp:coreProperties");
    return;
  }
  const find = (namespace: string, localName: string): Element | undefined =>
    Array.from(root.children).find(
      (child) => child.namespaceURI === namespace && child.localName === localName,
    );
  if (!find(DC_NAMESPACE, "title")) {
    violations.push("docProps/core.xml: missing dc:title");
  }
  const modified = find(DCTERMS_NAMESPACE, "modified");
  if (!modified) {
    violations.push("docProps/core.xml: missing dcterms:modified");
  } else {
    if (modified.getAttributeNS(XSI_NAMESPACE, "type") !== "dcterms:W3CDTF") {
      violations.push('docProps/core.xml: dcterms:modified must carry xsi:type="dcterms:W3CDTF"');
    }
    if (!W3CDTF_PATTERN.test(modified.textContent ?? "")) {
      violations.push(
        `docProps/core.xml: dcterms:modified value "${modified.textContent}" is not a W3CDTF timestamp`,
      );
    }
  }
}

function checkContentTypes(doc: Document, partNames: string[], violations: string[]): void {
  const root = doc.documentElement;
  if (root.namespaceURI !== CONTENT_TYPES_NAMESPACE || root.localName !== "Types") {
    violations.push("[Content_Types].xml: root element is not Types");
    return;
  }
  const defaults = new Map<string, string>();
  const overrides = new Map<string, string>();
  for (const child of Array.from(root.children)) {
    if (child.namespaceURI !== CONTENT_TYPES_NAMESPACE) {
      violations.push(`[Content_Types].xml: unexpected element ${child.nodeName}`);
      continue;
    }
    if (child.localName === "Default") {
      const extension = child.getAttribute("Extension");
      const contentType = child.getAttribute("ContentType");
      if (!extension || !contentType) {
        violations.push("[Content_Types].xml: Default needs Extension and ContentType");
      } else if (defaults.has(extension.toLowerCase())) {
        // OPC (ECMA-376 Part 2 §10.1.2.2): Extension is unique among Defaults.
        violations.push(`[Content_Types].xml: duplicate Default for extension "${extension}"`);
      } else {
        defaults.set(extension.toLowerCase(), contentType);
      }
    } else if (child.localName === "Override") {
      const partName = child.getAttribute("PartName");
      const contentType = child.getAttribute("ContentType");
      if (!partName || !contentType) {
        violations.push("[Content_Types].xml: Override needs PartName and ContentType");
      } else if (overrides.has(partName)) {
        // OPC (ECMA-376 Part 2 §10.1.2.3): PartName is unique among Overrides.
        // A save that re-writes a part the reader consumed used to append a
        // second Override for it and leave the source's in place, so a
        // document with a header went invalid on its second save.
        violations.push(`[Content_Types].xml: duplicate Override for part ${partName}`);
      } else {
        overrides.set(partName, contentType);
      }
    } else {
      violations.push(`[Content_Types].xml: unexpected element ${child.localName}`);
    }
  }

  const contentTypeOf = (partName: string): string | undefined => {
    const override = overrides.get(`/${partName}`);
    if (override) {
      return override;
    }
    const dot = partName.lastIndexOf(".");
    return dot === -1 ? undefined : defaults.get(partName.slice(dot + 1).toLowerCase());
  };
  for (const partName of partNames) {
    if (partName === "[Content_Types].xml") {
      continue;
    }
    if (!contentTypeOf(partName)) {
      violations.push(`[Content_Types].xml: no content type declared for ${partName}`);
    }
  }
  // The inverse: an Override for a part the package does not contain. A save
  // that stops writing a part has to withdraw its Override with it, or the
  // package declares a content type for nothing.
  for (const partName of overrides.keys()) {
    if (!partNames.includes(partName.replace(/^\//, ""))) {
      violations.push(`[Content_Types].xml: Override names missing part ${partName}`);
    }
  }
  const mainType = contentTypeOf("word/document.xml");
  if (
    mainType !== undefined &&
    mainType !== "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
  ) {
    violations.push(
      `[Content_Types].xml: word/document.xml has content type ${mainType}, not the WordprocessingML main part type`,
    );
  }
  if (
    contentTypeOf("docProps/core.xml") !== undefined &&
    contentTypeOf("docProps/core.xml") !== "application/vnd.openxmlformats-package.core-properties+xml"
  ) {
    violations.push("[Content_Types].xml: docProps/core.xml is not typed as core-properties+xml");
  }
}

function checkPackageRels(doc: Document, partNames: string[], violations: string[]): void {
  const root = doc.documentElement;
  if (root.namespaceURI !== RELATIONSHIPS_NAMESPACE || root.localName !== "Relationships") {
    violations.push("_rels/.rels: root element is not Relationships");
    return;
  }
  const byType = new Map<string, string>();
  for (const child of Array.from(root.children)) {
    const type = child.getAttribute("Type");
    const target = child.getAttribute("Target");
    if (child.namespaceURI !== RELATIONSHIPS_NAMESPACE || child.localName !== "Relationship") {
      violations.push(`_rels/.rels: unexpected element ${child.nodeName}`);
    } else if (!child.getAttribute("Id") || !type || !target) {
      violations.push("_rels/.rels: Relationship needs Id, Type and Target");
    } else {
      byType.set(type, target.replace(/^\//, ""));
    }
  }
  for (const [type, label] of [
    [OFFICE_DOCUMENT_RELATIONSHIP, "officeDocument"],
    [CORE_PROPERTIES_RELATIONSHIP, "core-properties"],
  ]) {
    const target = byType.get(type);
    if (!target) {
      violations.push(`_rels/.rels: missing ${label} relationship`);
    } else if (!partNames.includes(target)) {
      violations.push(`_rels/.rels: ${label} relationship targets missing part ${target}`);
    }
  }
}

/** The part a relationship Target names, resolved against the part its
 * `.rels` belongs to. `_rels/.rels` describes the package root, so its own
 * source part is "". */
function resolveRelTarget(sourcePart: string, target: string): string {
  if (target.startsWith("/")) {
    return target.slice(1);
  }
  const slash = sourcePart.lastIndexOf("/");
  const segments = slash === -1 ? [] : sourcePart.slice(0, slash).split("/");
  for (const segment of target.split("/")) {
    if (segment === "" || segment === ".") {
      continue;
    }
    if (segment === "..") {
      segments.pop();
    } else {
      segments.push(segment);
    }
  }
  return segments.join("/");
}

/** The part whose relationships a `.rels` part describes: `word/_rels/
 * document.xml.rels` -> `word/document.xml`, `_rels/.rels` -> "". */
function relsSourcePart(relsName: string): string | null {
  const match = /^(.*?)_rels\/([^/]*)\.rels$/.exec(relsName);
  if (!match) {
    return null;
  }
  return `${match[1]}${match[2]}`;
}

/** Every `.rels` part's relationships, checked for the two invariants an
 * append-only patch can break: Id unique within its part, and Target naming
 * a part the package actually contains. A save that re-writes a part the
 * reader consumed used to append a fresh relationship beside the source's
 * stale one, so both grew without bound across saves. */
function checkRelationshipParts(
  parts: Map<string, Uint8Array>,
  parsePart: (name: string) => Document | undefined,
  violations: string[],
): void {
  for (const name of parts.keys()) {
    const source = relsSourcePart(name);
    if (source === null) {
      continue;
    }
    const doc = parsePart(name);
    if (!doc || doc.documentElement.namespaceURI !== RELATIONSHIPS_NAMESPACE) {
      continue; // Shape/namespace already reported by the caller's own checks.
    }
    const seenIds = new Set<string>();
    for (const child of Array.from(doc.documentElement.children)) {
      if (child.namespaceURI !== RELATIONSHIPS_NAMESPACE || child.localName !== "Relationship") {
        continue;
      }
      const id = child.getAttribute("Id");
      if (id) {
        if (seenIds.has(id)) {
          violations.push(`${name}: duplicate relationship Id "${id}"`);
        }
        seenIds.add(id);
      }
      const target = child.getAttribute("Target");
      if (!target || child.getAttribute("TargetMode") === "External") {
        continue;
      }
      const resolved = resolveRelTarget(source, target);
      if (!parts.has(resolved)) {
        violations.push(`${name}: relationship "${id}" targets missing part ${resolved}`);
      }
    }
  }
}

/** Validates `.docx` bytes produced by the writer. Returns the list of
 * OPC/WordprocessingML violations found; an empty list means valid. */
export async function validateDocx(data: ArrayBuffer | Uint8Array): Promise<string[]> {
  const violations: string[] = [];

  let parts: Map<string, Uint8Array>;
  try {
    parts = await unzip(data);
  } catch (error) {
    if (error instanceof ZipError) {
      return [`package: not a valid ZIP archive (${error.message})`];
    }
    throw error;
  }
  const partNames = Array.from(parts.keys());

  for (const required of REQUIRED_PARTS) {
    if (!parts.has(required)) {
      violations.push(`package: missing required part ${required}`);
    }
  }

  const decoder = new TextDecoder("utf-8");
  const parsePart = (name: string): Document | undefined => {
    const raw = parts.get(name);
    if (!raw) {
      return undefined;
    }
    try {
      return parseXml(decoder.decode(raw));
    } catch (error) {
      if (error instanceof XmlParseError) {
        violations.push(`${name}: not well-formed XML`);
        return undefined;
      }
      throw error;
    }
  };

  const contentTypes = parsePart("[Content_Types].xml");
  if (contentTypes) {
    checkContentTypes(contentTypes, partNames, violations);
  }
  const packageRels = parsePart("_rels/.rels");
  if (packageRels) {
    checkPackageRels(packageRels, partNames, violations);
  }
  checkRelationshipParts(parts, parsePart, violations);
  const documentPart = parsePart("word/document.xml");
  if (documentPart) {
    checkDocumentPart(documentPart, violations);
  }
  const stylesPart = parsePart("word/styles.xml");
  if (stylesPart) {
    checkStylesPart(stylesPart, violations);
  }
  const numberingPart = parsePart("word/numbering.xml");
  if (numberingPart) {
    checkNumberingPart(numberingPart, violations);
  }
  const corePart = parsePart("docProps/core.xml");
  if (corePart) {
    checkCorePropertiesPart(corePart, violations);
  }

  return violations;
}
