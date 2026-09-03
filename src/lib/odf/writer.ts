// Serialises a DocumentModel into an OpenDocument Text package — the
// inverse of reader.ts and the ODF mirror of docx/writer.ts. Every field
// the reader maps into the model is written back out (text, run formatting,
// paragraph formatting, style references, page layout), each as the exact
// inverse of the reader's mapping, so open -> save -> re-open reproduces
// the model.
//
// Emission choices that keep the round trip exact:
//  - ODF has no direct formatting: every run and paragraph property is
//    expressed through an automatic style in content.xml, referenced by
//    text:style-name. A paragraph's automatic style is named after its
//    styleId so the model's style reference survives re-opening; run styles
//    get generated "T<n>" names, which is lossless because the reader never
//    records span style names in the model.
//  - Styles carry the model's *effective* formatting and no parents. The
//    reader resolves style chains into the model, so re-emitting the merged
//    result as a flat style reproduces it without reconstructing the
//    source's style hierarchy.
//  - styles.xml's default-style carries the model's default font family and
//    size — the one place the reader looks for defaultStyles.fontFamily/
//    fontSizePt. The reader merges those defaults into every run it reads;
//    that cannot disturb a round trip because any model that came from an
//    ODF file already carries the merged values on each run.
//  - Lengths are formatted by the exact-round-trip helpers in units.ts, so
//    re-parsing yields bit-identical doubles.
//  - Literal spaces that ODF whitespace collapsing would eat (leading,
//    consecutive, after a tab or line break) are emitted as text:s.
import type {
  Block,
  BreakKind,
  Chart,
  CommentThread,
  DocumentModel,
  DocumentProperties,
  FieldSpec,
  Image,
  ImageWrapMode,
  NamedStyle,
  NoteContent,
  NoteKind,
  NoteNumberFormat,
  NoteSettings,
  PageNumberFormat,
  Paragraph,
  PageBorders,
  ParagraphFormatting,
  Run,
  RunFormatting,
  RefFormat,
  Section,
  SectionColumns,
  Shape,
  Connector,
  Table,
  TableCell,
  TableRow,
  TableOfContents,
  TocEntry,
  TocOptions,
  TableOfFigures,
  TofEntry,
  TrackedChangeInfo,
} from "../documentModel";
import type { ListDefinition } from "../core/lists";
import type { Margins, PageSize } from "../core/pageSetup";
import { computeOwnerGrid, type BorderSide } from "../core/tables";
import { borderShorthand as odfBorderShorthand, DE_BORDER_NAMESPACE } from "./borders";
import { paragraphPaddingPt } from "../paragraphBorders";
import {
  FIELD_PLACEHOLDER_TEXT,
  isEvaluatedField,
  isMergeChipField,
  isTable,
  NOTE_PLACEHOLDER_TEXT,
  resolveColumnWidthsCm,
  sectionMargins,
  sectionPageSize,
} from "../documentModel";
import { chipLabel } from "../mergeChips";
import { dateTimeFormatEntry, odtDateStyleName, type DateTimeFormatEntry } from "../fields";
import { collectNoteRefs, noteLabels, referencedNotes } from "../notes";
import { composeLanguageTag, DEFAULT_LANGUAGE, ODF_NO_LINGUISTIC_CONTENT, splitLanguageTag } from "../core/language";
import { replyParagraphText, visibleCommentThreads } from "../comments";
import { collectChanges } from "../trackChanges";
import {
  collectBarTabDropRecords,
  collectBorderStyleDropRecords,
  collectDropRecords,
  collectEditedChartDropRecords,
  opaquePartBytes,
  type WriteResult,
} from "../opaque";
import { equalStructural } from "../equality";
import { TOC_PLACEHOLDER_TEXT } from "../toc";
import { TOF_PLACEHOLDER_TEXT } from "../tof";
import {
  BULLET_CHARS,
  clampListLevel,
  isBulletStyle,
  LIST_MARKER_WIDTH_CM,
  planListLevels,
  type PlannedListLevel,
} from "../core/lists";
import { definedNumberIndentCm, definedTextIndentCm } from "../core/listDefinitions";
import { zip, type ZipEntryInput } from "../core/zipWriter";
import { escapeXml, firstChildElement, parseXml, serializeElement } from "../core/xml";
import { ODT_MIMETYPE } from "./reader";
import { planAnchoredDrawings } from "../drawingAnchors";
import { drawFrameXml, graphicStyleXml } from "./drawing";
import {
  createOdfDrawingDefs,
  ODF_DRAWING_DEF_LOCAL_NAMES,
  shapeElementXml,
  shapeGraphicStyleXml,
  type OdfDrawingDefs,
  type OdfShapeStyleWriteContext,
} from "./shape";
import { connectorElementXml, connectorGraphicStyleXml, odfDrawId } from "./connector";
import { chartFrameXml } from "./chart";
import { chartObjectDirsInManifest, planOdfChartWrite, type OdfChartWritePlan } from "./chartWritePlan";
import { formulaFrameXml, formulaObjectDirsInManifest } from "./formula";
import { formulaFrameAltText, planOdfFormulaWrite, type OdfFormulaWritePlan } from "./formulaWritePlan";
import { planOdtMediaWrite, type OdtManifestEntry, type OdtMediaWritePlan } from "./media";
import { HIGHLIGHT_HEX } from "../core/colors";
import { catalogToNamedStyles, parseOdfStyleCatalog } from "./styles";
import { tableTemplateXml, tableUseStyleAttrs, templateCellStyleNames } from "./tableTemplate";
import {
  DRAW_NAMESPACE,
  FO_NAMESPACE,
  NUMBER_NAMESPACE,
  OFFICE_NAMESPACE,
  STYLE_NAMESPACE,
  TABLE_NAMESPACE,
  TEXT_NAMESPACE,
  XLINK_NAMESPACE,
  styleAttr,
} from "../core/odfXml";

// Declared locally on each number:date-style/time-style element (the
// TABLE_XMLNS pattern), so content.xml is unchanged for documents with no
// date/time field (SPEC-053).
const NUMBER_XMLNS = ` xmlns:number="${NUMBER_NAMESPACE}"`;
import { inlineWriteLayout, linkTargetOf, type LinkTarget } from "../links";
import { cmToOdfLength, multiplierToOdfPercentage, ptToOdfLength } from "./units";
import { buildMetaXml, type OdfMetaInput } from "../core/odfMeta";
import { stripPersonalInfo } from "../personalInfo";

export { buildMetaXml } from "../core/odfMeta";

export { ODT_MIMETYPE };

const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8"?>';

const MANIFEST_NAMESPACE = "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0";
const DC_NAMESPACE = "http://purl.org/dc/elements/1.1/";

// Declared on every document root; styles.xml needs all four, the others a
// subset, and surplus declarations are harmless.
const DOCUMENT_XMLNS =
  ` xmlns:office="${OFFICE_NAMESPACE}"` +
  ` xmlns:style="${STYLE_NAMESPACE}"` +
  ` xmlns:text="${TEXT_NAMESPACE}"` +
  ` xmlns:fo="${FO_NAMESPACE}"`;

// Declared locally on every table-related element rather than on the
// document root, so content.xml's bytes are unchanged for documents without
// tables (SPEC-022).
const TABLE_XMLNS = ` xmlns:table="${TABLE_NAMESPACE}"`;
// SPEC-076: the banding half of the table model (table:table-template and the
// table:use-*-styles attributes) lives in odf/tableTemplate.ts.

// Declared once on each link-carrying paragraph (SPEC-025) — the TABLE_XMLNS
// pattern, so content.xml is unchanged for documents without hyperlinks —
// rather than on every text:a, so a paragraph with several links still
// carries a single declaration.
const LINK_XMLNS = ` xmlns:xlink="${XLINK_NAMESPACE}"`;

/** The fallback title for a blank/whitespace-only model title — "Document 1"
 * per SPEC-007's decisions (unlike the .docx writer, which falls back to the
 * product's default title). */
export const UNTITLED_ODT_TITLE = "Document 1";

/** The suggested download file name: `<title>.odt` (AC-2). */
export function odtFileName(title: string): string {
  const trimmed = title.trim();
  return `${trimmed.length > 0 ? trimmed : UNTITLED_ODT_TITLE}.odt`;
}

/** A run's highlight as an fo:background-color value: palette names map to
 * their hex colour, raw CSS colours pass through, and the explicit
 * "No colour" (`false`) becomes "transparent" — an explicit off the reader
 * maps back to `false`, so it keeps overriding an inherited background
 * (SPEC-013). An OOXML name outside the palette (which no reader produces)
 * has no ODF representation. */
function highlightToOdfColor(highlight: string | false | undefined): string | undefined {
  if (highlight === undefined) {
    return undefined;
  }
  if (highlight === false) {
    return "transparent";
  }
  return highlight.startsWith("#") ? highlight : HIGHLIGHT_HEX[highlight];
}

/** A style:text-properties element for the run formatting, or "" when there
 * is nothing to say. Explicit false values are written as "normal"/"none" —
 * the model only holds false where the source explicitly switched the
 * property off, and dropping that would change the re-opened model.
 * `fontSizePercent` is the named-style case (SPEC-019): a proportional size
 * a style definition carries when it sets no absolute one. */
function textPropertiesXml(formatting: RunFormatting, fontSizePercent?: number): string {
  const attrs: string[] = [];
  if (formatting.fontFamily !== undefined) {
    attrs.push(`fo:font-family="${escapeXml(formatting.fontFamily)}"`);
  }
  if (formatting.fontSizePt !== undefined) {
    attrs.push(`fo:font-size="${ptToOdfLength(formatting.fontSizePt)}"`);
  } else if (fontSizePercent !== undefined) {
    attrs.push(`fo:font-size="${fontSizePercent}%"`);
  }
  if (formatting.bold !== undefined) {
    attrs.push(`fo:font-weight="${formatting.bold ? "bold" : "normal"}"`);
  }
  if (formatting.italic !== undefined) {
    attrs.push(`fo:font-style="${formatting.italic ? "italic" : "normal"}"`);
  }
  if (formatting.underline !== undefined) {
    attrs.push(`style:text-underline-style="${formatting.underline ? "solid" : "none"}"`);
  }
  if (formatting.strikethrough !== undefined) {
    attrs.push(`style:text-line-through-style="${formatting.strikethrough ? "solid" : "none"}"`);
  }
  if (formatting.color !== undefined) {
    attrs.push(`fo:color="${escapeXml(formatting.color)}"`);
  }
  const background = highlightToOdfColor(formatting.highlight);
  if (background !== undefined) {
    attrs.push(`fo:background-color="${escapeXml(background)}"`);
  }
  // ODF has the one background slot, so a shadingFill preserved beside a
  // highlight cannot be carried — the cross-format save reports it as
  // dropped (see opaque.ts's runShading records) rather than writing it.
  if (formatting.vertAlign !== undefined) {
    // 58% is the conventional sub/superscript font height LibreOffice
    // writes; the reader only reads the sub/super keyword back.
    const position = formatting.vertAlign === "subscript" ? "sub 58%" : "super 58%";
    attrs.push(`style:text-position="${position}"`);
  }
  // SPEC-035: "do not check spelling" maps to ODF's "no linguistic content"
  // language (fo:language="zxx") — there is no separate flag, so this wins
  // over an explicit `lang` on the same run (the run's real language cannot
  // be expressed alongside it in ODF; see lib/language.ts's
  // ODF_NO_LINGUISTIC_CONTENT).
  if (formatting.noProof) {
    attrs.push(`fo:language="${ODF_NO_LINGUISTIC_CONTENT}"`);
  } else if (formatting.lang !== undefined) {
    const { language, country } = splitLanguageTag(formatting.lang);
    attrs.push(`fo:language="${escapeXml(language)}"`);
    if (country !== undefined) {
      attrs.push(`fo:country="${escapeXml(country)}"`);
    }
  }
  return attrs.length > 0 ? `<style:text-properties ${attrs.join(" ")}/>` : "";
}

/** A paragraph's border/shading attributes on `style:paragraph-properties`
 * (SPEC-074 AC-29): `fo:border-*` per edge, `fo:padding-*` for the
 * border-to-text space (the 4pt default written explicitly, so a foreign
 * consumer sees what this product renders), `fo:background-color` for the
 * shading, and `style:join-border` — ODF's own name for the merge behaviour
 * borderRuns implements (AC-12), written `"true"` because this product
 * always merges identical neighbours.
 *
 * The `between` edge has no ODF attribute at all, so it rides in the
 * extension namespace rather than being silently lost — the same convention
 * SPEC-073's `de:shadow-blur-pt` uses. `xmlns:de` is declared on the same
 * element, keeping the declaration local (see odf/shape.ts). */
function paragraphBorderAttrs(formatting: ParagraphFormatting): string[] {
  const attrs: string[] = [];
  const borders = formatting.borders;
  if (borders) {
    let anyEdge = false;
    for (const edge of ["top", "bottom", "left", "right"] as const) {
      const side = borders[edge];
      if (side) {
        attrs.push(`fo:border-${edge}="${odfBorderShorthand(side)}"`);
        attrs.push(`fo:padding-${edge}="${ptToOdfLength(paragraphPaddingPt(borders, edge))}"`);
        anyEdge = true;
      }
    }
    if (borders.between) {
      attrs.push(
        `xmlns:de="${DE_BORDER_NAMESPACE}"`,
        `de:border-between="${escapeXml(odfBorderShorthand(borders.between))}"`,
      );
      anyEdge = true;
    }
    if (anyEdge) {
      attrs.push('style:join-border="true"');
    } else {
      // A borders object that paints nothing is the explicit override of a
      // style's box (SPEC-074 AC-17). ODF spells it with the shorthand:
      // `fo:border="none"` on the paragraph's automatic style beats the
      // named style's edges, where writing no attribute at all would let
      // them through again on the next open.
      attrs.push('fo:border="none"');
    }
  }
  if (formatting.shadingFill !== undefined) {
    attrs.push(`fo:background-color="${escapeXml(formatting.shadingFill)}"`);
  }
  return attrs;
}

/** A style:paragraph-properties element for the paragraph formatting, or ""
 * when there is nothing to say. `forcedBreakBefore` is the manual-break case
 * (SPEC-028 AC-32): ODF has no standalone break element, so a BreakBlock
 * becomes fo:break-before on the FOLLOWING paragraph's automatic style —
 * the same attribute the pageBreakBefore formatting flag emits, ODF's one
 * channel for both (the reader canonicalises it back to a BreakBlock). */
function paragraphPropertiesXml(
  formatting: ParagraphFormatting,
  forcedBreakBefore?: BreakKind,
  pageNumberStart?: number,
): string {
  const attrs: string[] = [];
  if (formatting.alignment !== undefined) {
    // The model's alignment values are all legal fo:text-align keywords.
    attrs.push(`fo:text-align="${formatting.alignment}"`);
  }
  // The model's indents are logical start/end; ODF's fo:margin-left/right
  // are physical, so an RTL paragraph writes start on the right edge and
  // end on the left — the exact inverse of the reader's post-resolution
  // swap (SPEC-014). fo:text-indent is start-relative and never swaps.
  const rtl = formatting.direction === "rtl";
  const marginLeftCm = rtl ? formatting.indentEndCm : formatting.indentStartCm;
  const marginRightCm = rtl ? formatting.indentStartCm : formatting.indentEndCm;
  if (marginLeftCm !== undefined) {
    attrs.push(`fo:margin-left="${cmToOdfLength(marginLeftCm)}"`);
  }
  if (marginRightCm !== undefined) {
    attrs.push(`fo:margin-right="${cmToOdfLength(marginRightCm)}"`);
  }
  if (formatting.indentFirstLineCm !== undefined) {
    attrs.push(`fo:text-indent="${cmToOdfLength(formatting.indentFirstLineCm)}"`);
  }
  if (formatting.spaceBeforePt !== undefined) {
    attrs.push(`fo:margin-top="${ptToOdfLength(formatting.spaceBeforePt)}"`);
  }
  if (formatting.spaceAfterPt !== undefined) {
    attrs.push(`fo:margin-bottom="${ptToOdfLength(formatting.spaceAfterPt)}"`);
  }
  // lineSpacingPt wins over lineSpacing when both are set, matching the
  // model's documented precedence and the .docx writer. The absent-rule
  // default is "at least" (style:line-height-at-least); an explicit "exact"
  // is ODF's fixed fo:line-height (SPEC-014).
  if (formatting.lineSpacingPt !== undefined) {
    const length = ptToOdfLength(formatting.lineSpacingPt);
    attrs.push(
      formatting.lineSpacingRule === "exact"
        ? `fo:line-height="${length}"`
        : `style:line-height-at-least="${length}"`,
    );
  } else if (formatting.lineSpacing !== undefined) {
    attrs.push(`fo:line-height="${multiplierToOdfPercentage(formatting.lineSpacing)}"`);
  }
  // Pagination flags (SPEC-028 AC-22..29) — only the non-default states are
  // written, the exact inverse of parseParagraphProperties. A manual break
  // forces fo:break-before even though this paragraph's own formatting
  // lacks pageBreakBefore (AC-32, see forcedBreakBefore above); a "column"
  // break is written for real ODF consumers but dropped on re-open until
  // columns exist (SPEC-030).
  const breakBefore = forcedBreakBefore ?? (formatting.pageBreakBefore ? "page" : undefined);
  if (breakBefore !== undefined) {
    attrs.push(`fo:break-before="${breakBefore}"`);
  }
  if (formatting.keepNext) {
    attrs.push('fo:keep-with-next="always"');
  }
  if (formatting.keepLines) {
    attrs.push('fo:keep-together="always"');
  }
  if (formatting.widowControl === false) {
    attrs.push('fo:widows="0"', 'fo:orphans="0"');
  }
  if (rtl) {
    attrs.push('style:writing-mode="rl-tb"');
  }
  // Paragraph borders and shading (SPEC-074 AC-29).
  attrs.push(...paragraphBorderAttrs(formatting));
  // "Start numbering at" (SPEC-029 AC-22) — the section-starting
  // paragraph's own restart value; absent means "continue from the
  // previous section", which needs no attribute at all.
  if (pageNumberStart !== undefined) {
    attrs.push(`style:page-number="${pageNumberStart}"`);
  }
  // Custom tab stops (SPEC-032 AC-19..25): a style:tab-stops child. The
  // leader is written BOTH as style:leader-style and style:leader-text so
  // either form recovers it on read — including foreign files that carry
  // only the style keyword. Bar stops (SPEC-054) have no ODF equivalent and
  // are dropped here — the DropRecord for each is reported by opaque.ts's
  // collectPackageDropRecords (AC-38).
  let tabStops = "";
  const writableStops = (formatting.tabStops ?? []).filter((stop) => stop.alignment !== "bar");
  if (writableStops.length > 0) {
    tabStops = writableStops
      .map((stop) => {
        let stopAttrs = `style:position="${cmToOdfLength(stop.positionCm)}"`;
        if (stop.alignment === "decimal") {
          stopAttrs += ' style:type="char" style:char="."';
        } else if (stop.alignment !== "left") {
          stopAttrs += ` style:type="${stop.alignment}"`;
        }
        if (stop.leader === "dot") {
          stopAttrs += ' style:leader-style="dotted" style:leader-text="."';
        } else if (stop.leader === "hyphen") {
          stopAttrs += ' style:leader-style="dash" style:leader-text="-"';
        } else if (stop.leader === "underscore") {
          stopAttrs += ' style:leader-style="solid" style:leader-text="_"';
        }
        return `<style:tab-stop ${stopAttrs}/>`;
      })
      .join("");
    tabStops = `<style:tab-stops>${tabStops}</style:tab-stops>`;
  }
  if (attrs.length === 0 && tabStops === "") {
    return "";
  }
  const attrText = attrs.length > 0 ? ` ${attrs.join(" ")}` : "";
  if (tabStops === "") {
    return `<style:paragraph-properties${attrText}/>`;
  }
  return `<style:paragraph-properties${attrText}>${tabStops}</style:paragraph-properties>`;
}

// style:num-format values for the number marker styles (ODF 19.500) — the
// inverse of the reader catalog's NUMBER_STYLE_BY_FORMAT.
const NUM_FORMAT_BY_STYLE: Record<string, string> = {
  decimal: "1",
  lowerAlpha: "a",
  upperAlpha: "A",
  lowerRoman: "i",
  upperRoman: "I",
};

/** One list level of a generated text:list-style (SPEC-018 AC-19). The
 * indent is written in the label-alignment form, whose fo:margin-left
 * carries the level's text indent as a single exact length — the older
 * space-before + min-label-width split would re-import with float noise. */
function listLevelXml(level: number, def: PlannedListLevel): string {
  const props =
    '<style:list-level-properties text:list-level-position-and-space-mode="label-alignment">' +
    '<style:list-level-label-alignment text:label-followed-by="listtab"' +
    ` fo:margin-left="${cmToOdfLength(def.indentCm)}"` +
    ` fo:text-indent="-${cmToOdfLength(LIST_MARKER_WIDTH_CM)}"/>` +
    "</style:list-level-properties>";
  if (def.type === "bullet") {
    return (
      `<text:list-level-style-bullet text:level="${level}" text:bullet-char="${BULLET_CHARS[def.style]}">` +
      `${props}</text:list-level-style-bullet>`
    );
  }
  const start = def.startAt !== 1 ? ` text:start-value="${def.startAt}"` : "";
  return (
    `<text:list-level-style-number text:level="${level}"` +
    ` style:num-format="${NUM_FORMAT_BY_STYLE[def.style]}" style:num-suffix="."${start}>` +
    `${props}</text:list-level-style-number>`
  );
}

/** One level of a SPEC-032 definition's text:list-style (AC-27). Unlike the
 * SPEC-018 form above, `text:display-levels` is ALWAYS written (even at its
 * default of 1): its presence is the reader's marker that the style is a
 * full multilevel definition, so a re-opened definition keeps its identity
 * instead of degrading into an adjacency list. */
function definedListLevelXml(def: ListDefinition, level: number): string {
  const levelDef = def.levels[level - 1] ?? { style: "decimal" };
  const textIndentCm = definedTextIndentCm(def, level);
  const numberIndentCm = definedNumberIndentCm(def, level);
  const alignment = levelDef.alignment ?? "left";
  const props =
    '<style:list-level-properties text:list-level-position-and-space-mode="label-alignment"' +
    (alignment !== "left" ? ` fo:text-align="${alignment}"` : "") +
    ">" +
    '<style:list-level-label-alignment text:label-followed-by="listtab"' +
    ` fo:margin-left="${cmToOdfLength(textIndentCm)}"` +
    ` fo:text-indent="-${cmToOdfLength(Math.max(0, textIndentCm - numberIndentCm))}"/>` +
    "</style:list-level-properties>";
  const prefix =
    levelDef.prefix !== undefined && levelDef.prefix !== ""
      ? ` style:num-prefix="${escapeXml(levelDef.prefix)}"`
      : "";
  const suffix =
    levelDef.suffix !== undefined && levelDef.suffix !== ""
      ? ` style:num-suffix="${escapeXml(levelDef.suffix)}"`
      : "";
  if (isBulletStyle(levelDef.style)) {
    return (
      `<text:list-level-style-bullet text:level="${level}"${prefix}${suffix}` +
      ` text:bullet-char="${BULLET_CHARS[levelDef.style]}">` +
      `${props}</text:list-level-style-bullet>`
    );
  }
  const include = Math.min(Math.max(Math.round(levelDef.includeLevels ?? 1), 1), level);
  const start = (levelDef.startAt ?? 1) !== 1 ? ` text:start-value="${levelDef.startAt}"` : "";
  return (
    `<text:list-level-style-number text:level="${level}"${prefix}${suffix}` +
    ` style:num-format="${NUM_FORMAT_BY_STYLE[levelDef.style]}"` +
    ` text:display-levels="${include}"${start}>` +
    `${props}</text:list-level-style-number>`
  );
}

/** Assigns and serialises the automatic styles of content.xml. Paragraph
 * styles are named after the paragraph's styleId (all styleIds are reserved
 * up front, so a generated name can never shadow one); styleless formatted
 * paragraphs get "P<n>" and runs share "T<n>" styles per distinct
 * formatting; list runs share "L<n>" list styles per distinct level set
 * (SPEC-018). */
/** A cell's `BorderSide` as ODF's `fo:border`-family shorthand ("0.5pt
 * solid #ff0000", or "none"). SPEC-023 floors the width at a quarter point
 * so a hairline cell rule still renders; the shared mapping (odf/borders.ts,
 * SPEC-074) supplies the keyword. */
function borderShorthand(side: BorderSide): string {
  if (side.style === "none") {
    return "none";
  }
  return odfBorderShorthand({ ...side, widthPt: Math.max(0.25, side.widthPt) });
}

/** A cell's `style:table-cell-properties` attributes (SPEC-023 AC-10..12,
 * AC-15..17, AC-29, AC-30): borders, background shading, padding, vertical
 * alignment. Empty when the cell sets none of these — such a cell gets no
 * generated style at all (cellStyleName below), rather than a synthesised
 * default: unlike .docx's table-level `w:tblBorders` fallback (invisible to
 * the per-cell model), any per-cell default this writer invented would be
 * indistinguishable from real user/source data on the next parse and would
 * make an untouched save-and-reopen see borders that were never set. */
function cellStylePropertiesXml(cell: TableCell): string {
  const parts: string[] = [];
  const borders = cell.borders;
  if (borders?.top) {
    parts.push(`fo:border-top="${borderShorthand(borders.top)}"`);
  }
  if (borders?.bottom) {
    parts.push(`fo:border-bottom="${borderShorthand(borders.bottom)}"`);
  }
  if (borders?.left) {
    parts.push(`fo:border-left="${borderShorthand(borders.left)}"`);
  }
  if (borders?.right) {
    parts.push(`fo:border-right="${borderShorthand(borders.right)}"`);
  }
  if (cell.shadingFill !== undefined) {
    parts.push(`fo:background-color="${cell.shadingFill}"`);
  }
  if (cell.paddingCm !== undefined) {
    parts.push(`fo:padding="${cmToOdfLength(cell.paddingCm)}"`);
  }
  if (cell.verticalAlign !== undefined) {
    const value = cell.verticalAlign === "center" ? "middle" : cell.verticalAlign;
    parts.push(`style:vertical-align="${value}"`);
  }
  return parts.join(" ");
}

/** The names a serialised table references: its `table:name`, its
 * table-family style, one table-column style per column, and per-cell/
 * per-row styles for whichever cells/rows need one. `cellStyleName`/
 * `rowStyleName` are lazy and dedupe by the properties they'd emit, so two
 * cells with the same formatting (including "none", SPEC-023's default
 * look) share one definition. */
interface TableNames {
  name: string;
  styleName: string;
  columnStyleNames: string[];
  cellStyleName(cell: TableCell): string | undefined;
  rowStyleName(row: TableRow): string | undefined;
}

/** Position-dependent extras a paragraph's automatic style carries
 * (SPEC-028): a manual BreakBlock immediately before the paragraph (AC-32,
 * forced fo:break-before) and/or the master page the paragraph switches to
 * because it starts a section (AC-33, style:master-page-name — an attribute
 * of the style:style element itself, not of paragraph-properties). Both are
 * part of the style-dedup key: a section-starting paragraph must never
 * share its style with an identical paragraph elsewhere. */
interface ParagraphStyleOptions {
  breakBefore?: BreakKind;
  masterPageName?: string;
  /** "Start numbering at" (SPEC-029 AC-22) — style:page-number on
   * paragraph-properties. */
  pageNumberStart?: number;
}

interface AutomaticStyles {
  /** Serialised style:style elements, in first-use order. */
  defs: string[];
  paragraphNameFor(paragraph: Paragraph, options?: ParagraphStyleOptions): string | undefined;
  textNameFor(formatting: RunFormatting, styleId?: string): string;
  listNameFor(levels: PlannedListLevel[]): string;
  /** A SPEC-032 definition's own text:list-style (AC-27), one per defId. */
  definedListNameFor(def: ListDefinition): string;
  /** Resolves a paragraph's defId against the model's definitions. */
  listDefinitionFor(defId: string): ListDefinition | undefined;
  /** A fresh xml:id for one written text:list segment — document-unique, so
   * text:continue-list references stay unambiguous across cells/sections. */
  nextListXmlId(): string;
  tableNamesFor(table: Table, preserveNames: boolean): TableNames;
  /** A floating image's graphic-family style (SPEC-024 AC-13, AC-14) —
   * shared by every image with the same wrap mode; undefined for "inline",
   * which needs no positioning style of its own (see drawFrameXml). */
  graphicStyleNameFor(wrap: ImageWrapMode): string | undefined;
  /** A shape's own graphic-family style (SPEC-051): unlike an image, never
   * shared across shapes (fill/outline/wrap all vary per shape), so this
   * always mints a fresh name. */
  shapeStyleNameFor(shape: Shape): string;
  /** A connector's own graphic-family style (SPEC-070) — like a shape's, never
   * shared across connectors (stroke/markers/wrap vary), so it mints fresh. */
  connectorStyleNameFor(connector: Connector): string;
  /** The `number:date-style`/`number:time-style` name for a supported
   * SPEC-053 date/time format, minting and registering it on first use so a
   * `text:date`/`text:time` field can reference it (`style:data-style-name`).
   * Deduped by format. */
  dateStyleNameFor(entry: DateTimeFormatEntry): string;
  /** Records that a `text:user-field-get` (a SPEC-053 `docProperty` field)
   * references the named user field, so buildContentXml can emit a
   * `text:user-field-decls` block declaring it. */
  registerUserField(name: string): void;
  /** The user-field names collected via `registerUserField`, in first-use
   * order. */
  userFieldNames(): string[];
}

function createAutomaticStyles(model: DocumentModel, shapeStyleCtx?: OdfShapeStyleWriteContext): AutomaticStyles {
  const takenNames = new Set<string>();
  // SPEC-053: date/time number styles (deduped by format) and the user-field
  // names a `docProperty` field references.
  const dateStyleNames = new Map<string, string>();
  const userFields = new Set<string>();
  const takeTableNames = (table: Table) => {
    if (table.odtName) {
      takenNames.add(table.odtName);
    }
    if (table.odtStyleName) {
      takenNames.add(table.odtStyleName);
    }
    for (const columnName of table.odtColumnStyleNames ?? []) {
      if (columnName) {
        takenNames.add(columnName);
      }
    }
    for (const row of table.rows) {
      if (row.odtStyleName) {
        takenNames.add(row.odtStyleName);
      }
      for (const cell of row.cells) {
        if (cell.odtStyleName) {
          takenNames.add(cell.odtStyleName);
        }
      }
    }
  };
  for (const section of model.sections) {
    for (const block of section.blocks) {
      if (block.type === "table") {
        takeTableNames(block);
        for (const row of block.rows) {
          for (const cell of row.cells) {
            for (const paragraph of cell.blocks) {
              if (paragraph.styleId) {
                takenNames.add(paragraph.styleId);
              }
            }
          }
        }
        continue;
      }
      if (block.type === "image" || block.type === "break" || block.type === "toc" || block.type === "tof" || block.type === "chart" || block.type === "connector") {
        continue;
      }
      if (block.type === "shape") {
        for (const shapeBlock of block.text?.blocks ?? []) {
          if (shapeBlock.type === "paragraph" && shapeBlock.styleId) {
            takenNames.add(shapeBlock.styleId);
          }
        }
        continue;
      }
      if (block.styleId) {
        takenNames.add(block.styleId);
      }
    }
  }
  // Preserved automatic styles (graphic, table, … families) are re-emitted
  // verbatim beside the generated defs; their names must never be claimed
  // for a generated style.
  for (const preserved of model.sourceXml?.odtAutomaticStyles ?? []) {
    for (const match of preserved.matchAll(/style:name="([^"]*)"/g)) {
      takenNames.add(match[1]);
    }
  }

  // Paragraph and text styles share one defs list, serialised in first-use
  // order — deterministic, so a re-saved model reproduces identical bytes.
  const defs: string[] = [];
  // The serialised properties element doubles as the formatting signature:
  // emission order is fixed, so equal formatting yields equal XML.
  const paragraphNames = new Map<string, string>();
  const textNames = new Map<string, string>();
  const listNames = new Map<string, string>();
  const definedListNames = new Map<string, string>();
  const graphicNames = new Map<ImageWrapMode, string>();
  const boundStyleIds = new Set<string>();
  let listXmlIdCounter = 0;

  const claimName = (base: string, separator: string): string => {
    for (let n = 1; ; n++) {
      const candidate = `${base}${separator}${n}`;
      if (!takenNames.has(candidate)) {
        takenNames.add(candidate);
        return candidate;
      }
    }
  };

  return {
    defs,

    paragraphNameFor(paragraph: Paragraph, options?: ParagraphStyleOptions): string | undefined {
      const propertiesXml = paragraphPropertiesXml(
        paragraph.formatting ?? {},
        options?.breakBefore,
        options?.pageNumberStart,
      );
      const masterAttr =
        options?.masterPageName !== undefined
          ? ` style:master-page-name="${escapeXml(options.masterPageName)}"`
          : "";
      // Nothing to reference and nothing to say: emit no style at all, so
      // the paragraph re-opens without a styleId, matching the model.
      if (!paragraph.styleId && propertiesXml === "" && masterAttr === "") {
        return undefined;
      }
      const key = `${paragraph.styleId ?? ""} ${masterAttr} ${propertiesXml}`;
      const existing = paragraphNames.get(key);
      if (existing !== undefined) {
        return existing;
      }
      let name: string;
      if (paragraph.styleId && !boundStyleIds.has(paragraph.styleId)) {
        // The first formatting seen under a styleId gets the id itself, so
        // the reference round-trips.
        boundStyleIds.add(paragraph.styleId);
        name = paragraph.styleId;
      } else {
        // A styleless paragraph, or (possible only for models that came
        // from .docx, where direct formatting exists) a styleId already
        // bound to different formatting.
        name = paragraph.styleId ? claimName(paragraph.styleId, "-") : claimName("P", "");
      }
      paragraphNames.set(key, name);
      defs.push(
        propertiesXml === ""
          ? `<style:style style:name="${escapeXml(name)}" style:family="paragraph"${masterAttr}/>`
          : `<style:style style:name="${escapeXml(name)}" style:family="paragraph"${masterAttr}>${propertiesXml}</style:style>`,
      );
      return name;
    },

    dateStyleNameFor(entry: DateTimeFormatEntry): string {
      const existing = dateStyleNames.get(entry.format);
      if (existing !== undefined) {
        return existing;
      }
      const name = odtDateStyleName(entry);
      dateStyleNames.set(entry.format, name);
      const tag = entry.odtStyleTag === "time" ? "number:time-style" : "number:date-style";
      defs.push(`<${tag}${NUMBER_XMLNS} style:name="${escapeXml(name)}">${entry.odtStyleBody}</${tag}>`);
      return name;
    },

    registerUserField(name: string): void {
      userFields.add(name);
    },

    userFieldNames(): string[] {
      return [...userFields];
    },

    textNameFor(formatting: RunFormatting, styleId?: string): string {
      const propertiesXml = textPropertiesXml(formatting);
      // SPEC-042: a character-styled run's automatic style is named after
      // its `styleId` the first time (shadowing the named style of the same
      // name in office:styles, the same trick `paragraphNameFor` already
      // plays for paragraph styles) so the reference round-trips; a run
      // with no character style keeps the plain "T1"/"T2"/… counter.
      const key = `${styleId ?? ""} ${propertiesXml}`;
      const existing = textNames.get(key);
      if (existing !== undefined) {
        return existing;
      }
      let name: string;
      if (styleId && !boundStyleIds.has(styleId)) {
        boundStyleIds.add(styleId);
        name = styleId;
      } else if (styleId) {
        name = claimName(styleId, "-");
      } else {
        name = `T${textNames.size + 1}`;
      }
      textNames.set(key, name);
      defs.push(
        propertiesXml === ""
          ? `<style:style style:name="${escapeXml(name)}" style:family="text"/>`
          : `<style:style style:name="${escapeXml(name)}" style:family="text">${propertiesXml}</style:style>`,
      );
      return name;
    },

    listNameFor(levels: PlannedListLevel[]): string {
      const levelsXml = levels.map((def, index) => listLevelXml(index + 1, def)).join("");
      const existing = listNames.get(levelsXml);
      if (existing !== undefined) {
        return existing;
      }
      const name = claimName("L", "");
      listNames.set(levelsXml, name);
      defs.push(`<text:list-style style:name="${name}">${levelsXml}</text:list-style>`);
      return name;
    },

    definedListNameFor(def: ListDefinition): string {
      const existing = definedListNames.get(def.defId);
      if (existing !== undefined) {
        return existing;
      }
      const name = claimName("LD", "");
      definedListNames.set(def.defId, name);
      let levelsXml = "";
      for (let level = 1; level <= 9; level++) {
        levelsXml += definedListLevelXml(def, level);
      }
      defs.push(`<text:list-style style:name="${name}">${levelsXml}</text:list-style>`);
      return name;
    },

    listDefinitionFor(defId: string): ListDefinition | undefined {
      return model.listDefinitions?.find((entry) => entry.defId === defId);
    },

    nextListXmlId(): string {
      listXmlIdCounter += 1;
      return `list${listXmlIdCounter}`;
    },

    tableNamesFor(table: Table, preserveNames: boolean): TableNames {
      // With preserveNames (a same-format save), the source's own style
      // names are re-emitted — their definitions travel in the preserved
      // automatic styles — and only nameless parts (added or edited in the
      // editor) get generated styles. A cross-format save generates
      // everything.
      const totalCm = table.columnWidthsCm.reduce((sum, cm) => sum + cm, 0);
      let styleName = preserveNames ? table.odtStyleName : undefined;
      let base: string;
      if (styleName === undefined) {
        base = claimName("Table", "");
        styleName = base;
        // style:rel-width (SPEC-023 AC-23, AC-31) alongside the absolute
        // width: a percentage-of-text-area table still needs a fallback
        // absolute size for a consumer that ignores rel-width.
        const relWidth =
          table.widthMode === "percent" ? ` style:rel-width="${Math.round(table.widthPercent ?? 100)}%"` : "";
        defs.push(
          `<style:style style:name="${escapeXml(base)}" style:family="table">` +
            `<style:table-properties style:width="${cmToOdfLength(totalCm)}"${relWidth}/>` +
            "</style:style>",
        );
      } else {
        base = styleName;
      }
      const columnStyleNames = table.columnWidthsCm.map((widthCm, index) => {
        const preserved = preserveNames ? table.odtColumnStyleNames?.[index] : undefined;
        if (preserved !== undefined) {
          return preserved;
        }
        const columnName = claimName(base, ".");
        defs.push(
          `<style:style style:name="${escapeXml(columnName)}" style:family="table-column">` +
            `<style:table-column-properties style:column-width="${cmToOdfLength(widthCm)}"/>` +
            "</style:style>",
        );
        return columnName;
      });
      // Cell/row styles are generated lazily, one per distinct set of
      // properties, so two cells/rows with the same formatting share a
      // definition and a cell/row with nothing to say never claims one.
      const cellStyleNames = new Map<string, string>();
      const rowStyleNames = new Map<number, string>();
      return {
        name: table.odtName ?? base,
        styleName,
        columnStyleNames,
        cellStyleName(cell: TableCell): string | undefined {
          const propsXml = cellStylePropertiesXml(cell);
          if (propsXml === "") {
            return undefined;
          }
          const existing = cellStyleNames.get(propsXml);
          if (existing !== undefined) {
            return existing;
          }
          const name = claimName(base, ".Cell");
          cellStyleNames.set(propsXml, name);
          defs.push(
            `<style:style style:name="${escapeXml(name)}" style:family="table-cell">` +
              `<style:table-cell-properties ${propsXml}/>` +
              "</style:style>",
          );
          return name;
        },
        rowStyleName(row: TableRow): string | undefined {
          if (row.minHeightCm === undefined) {
            return undefined;
          }
          const existing = rowStyleNames.get(row.minHeightCm);
          if (existing !== undefined) {
            return existing;
          }
          const name = claimName(base, ".Row");
          rowStyleNames.set(row.minHeightCm, name);
          defs.push(
            `<style:style style:name="${escapeXml(name)}" style:family="table-row">` +
              `<style:table-row-properties style:min-row-height="${cmToOdfLength(row.minHeightCm)}"/>` +
              "</style:style>",
          );
          return name;
        },
      };
    },

    graphicStyleNameFor(wrap: ImageWrapMode): string | undefined {
      if (wrap === "inline") {
        return undefined;
      }
      const existing = graphicNames.get(wrap);
      if (existing !== undefined) {
        return existing;
      }
      const name = claimName("fr", "");
      graphicNames.set(wrap, name);
      defs.push(graphicStyleXml(name, wrap));
      return name;
    },

    shapeStyleNameFor(shape: Shape): string {
      const name = claimName("shfr", "");
      defs.push(shapeGraphicStyleXml(name, shape, shapeStyleCtx));
      return name;
    },

    connectorStyleNameFor(connector: Connector): string {
      const name = claimName("cnfr", "");
      defs.push(connectorGraphicStyleXml(name, connector, shapeStyleCtx));
      return name;
    },
  };
}

/** A contiguous run of list paragraphs as nested text:list markup (AC-19):
 * one outer text:list carrying the generated list style, a text:list-item
 * per item, and a nested style-less text:list (inheriting the outer style,
 * per ODF) opened inside the current item for each level of depth. Separate
 * runs get separate text:list elements, which is what restarts numbering
 * after an interruption (AC-11). */
function listRunXml(
  paragraphs: Paragraph[],
  styles: AutomaticStyles,
  includeOpaque: boolean,
  styleOptions?: ParagraphStyleOptions,
): string {
  const styleName = styles.listNameFor(planListLevels(paragraphs.map((p) => p.list!)));
  let xml = "";
  let depth = 0;
  let first = true;
  for (const paragraph of paragraphs) {
    const level = clampListLevel(paragraph.list!.level);
    if (depth === 0) {
      xml += `<text:list text:style-name="${escapeXml(styleName)}"><text:list-item>`;
      depth = 1;
    } else if (level <= depth) {
      while (depth > level) {
        xml += "</text:list-item></text:list>";
        depth--;
      }
      xml += "</text:list-item><text:list-item>";
    }
    while (depth < level) {
      xml += "<text:list><text:list-item>";
      depth++;
    }
    // A pending break/section start (SPEC-028) binds to the run's FIRST
    // paragraph — the one the source break/boundary sat before.
    xml += paragraphXml(paragraph, styles, includeOpaque, first ? styleOptions : undefined);
    first = false;
  }
  while (depth > 0) {
    xml += "</text:list-item></text:list>";
    depth--;
  }
  return xml;
}

/** One segment of a DEFINED list instance (SPEC-032 AC-27, AC-28) as a
 * text:list: same nested structure as listRunXml, but the outer element
 * carries the definition's list style, a document-unique xml:id, and — when
 * an earlier segment of the same instance was already written —
 * text:continue-list/text:continue-numbering, so ODF's own numbering
 * continues the interrupted sequence. A segment WITHOUT the continue
 * attributes restarts, which is exactly how a restarted instance
 * serialises. An item's explicit start value ("Set numbering value…")
 * rides on its own text:list-item as text:start-value. */
function definedListRunXml(
  paragraphs: Paragraph[],
  def: ListDefinition,
  styles: AutomaticStyles,
  includeOpaque: boolean,
  styleOptions: ParagraphStyleOptions | undefined,
  segment: { xmlId: string; continueFrom?: string },
): string {
  const styleName = styles.definedListNameFor(def);
  const continueAttrs =
    segment.continueFrom !== undefined
      ? ` text:continue-list="${escapeXml(segment.continueFrom)}" text:continue-numbering="true"`
      : "";
  const listOpen =
    `<text:list text:style-name="${escapeXml(styleName)}"` +
    ` xml:id="${escapeXml(segment.xmlId)}"${continueAttrs}>`;
  let xml = "";
  let depth = 0;
  let first = true;
  for (const paragraph of paragraphs) {
    const list = paragraph.list!;
    const level = clampListLevel(list.level);
    const startAttr =
      list.startAt !== undefined ? ` text:start-value="${list.startAt}"` : "";
    // The start value belongs on the item at the paragraph's OWN level, not
    // on the ancestor items opened on the way down.
    const itemAttr = (atLevel: number) => (atLevel === level ? startAttr : "");
    if (depth === 0) {
      xml += `${listOpen}<text:list-item${itemAttr(1)}>`;
      depth = 1;
    } else if (level <= depth) {
      while (depth > level) {
        xml += "</text:list-item></text:list>";
        depth--;
      }
      xml += `</text:list-item><text:list-item${itemAttr(level)}>`;
    }
    while (depth < level) {
      depth++;
      xml += `<text:list><text:list-item${itemAttr(depth)}>`;
    }
    xml += paragraphXml(paragraph, styles, includeOpaque, first ? styleOptions : undefined);
    first = false;
  }
  while (depth > 0) {
    xml += "</text:list-item></text:list>";
    depth--;
  }
  return xml;
}

/** A defined paragraph's instance signature, or null for a SPEC-018 item
 * (or a dangling defId, which degrades to the ad-hoc path — mirroring the
 * renderer's fallback). */
function instanceSignatureOf(paragraph: Paragraph, styles: AutomaticStyles): string | null {
  const list = paragraph.list;
  if (list?.defId === undefined || styles.listDefinitionFor(list.defId) === undefined) {
    return null;
  }
  return `${list.defId}\u0000${list.listId ?? ""}`;
}

/** A contiguous run of list paragraphs, split into per-instance segments
 * (SPEC-032): consecutive items of the SAME (defId, listId) share one
 * text:list; an instance change mid-run — a restart, "Set numbering
 * value…", or two independent lists that happen to touch — starts a new
 * one. This split is what the previous attempt lost restarts to: grouping
 * a whole run under the first paragraph's instance let ODF renumber a
 * restarted item as a continuation. `continuity` maps instance signatures
 * to the xml:id of their last written segment, one map per numbering
 * context (the body; each table cell). */
function listRunsXml(
  paragraphs: Paragraph[],
  styles: AutomaticStyles,
  includeOpaque: boolean,
  styleOptions: ParagraphStyleOptions | undefined,
  continuity: Map<string, string>,
): string {
  let xml = "";
  let i = 0;
  let first = true;
  while (i < paragraphs.length) {
    const signature = instanceSignatureOf(paragraphs[i], styles);
    let end = i + 1;
    while (end < paragraphs.length && instanceSignatureOf(paragraphs[end], styles) === signature) {
      end++;
    }
    const segment = paragraphs.slice(i, end);
    const options = first ? styleOptions : undefined;
    if (signature === null) {
      xml += listRunXml(segment, styles, includeOpaque, options);
    } else {
      const def = styles.listDefinitionFor(segment[0].list!.defId!)!;
      const xmlId = styles.nextListXmlId();
      const continueFrom = continuity.get(signature);
      continuity.set(signature, xmlId);
      xml += definedListRunXml(segment, def, styles, includeOpaque, options, {
        xmlId,
        continueFrom,
      });
    }
    first = false;
    i = end;
  }
  return xml;
}

/** Whitespace context threaded through a paragraph's runs: the last
 * character emitted so far — the writer-side mirror of the reader's
 * InlineAccumulator.lastChar, deciding which spaces need text:s. */
interface WhitespaceState {
  lastChar: string;
}

/** A run's text as literal text, text:s, text:tab and text:line-break —
 * the inverse of the reader's whitespace handling. A literal space survives
 * re-opening only after a non-whitespace character; every other space
 * (leading, consecutive, after a tab or break) must be an explicit text:s. */
function runContentXml(text: string, state: WhitespaceState): string {
  let xml = "";
  let buffer = "";
  const flush = () => {
    if (buffer.length > 0) {
      xml += escapeXml(buffer);
      buffer = "";
    }
  };
  for (let i = 0; i < text.length; ) {
    const char = text[i];
    if (char === "\t") {
      flush();
      xml += "<text:tab/>";
      state.lastChar = "\t";
      i++;
    } else if (char === "\n") {
      flush();
      xml += "<text:line-break/>";
      state.lastChar = "\n";
      i++;
    } else if (char === " ") {
      let count = 1;
      while (i + count < text.length && text[i + count] === " ") {
        count++;
      }
      let explicit = count;
      if (state.lastChar !== "" && !" \t\n".includes(state.lastChar)) {
        buffer += " ";
        explicit--;
      }
      flush();
      if (explicit > 0) {
        xml += explicit > 1 ? `<text:s text:c="${explicit}"/>` : "<text:s/>";
      }
      state.lastChar = " ";
      i += count;
    } else {
      buffer += char;
      state.lastChar = char;
      i++;
    }
  }
  flush();
  return xml;
}

// style:num-format for a page-number/page-count field (SPEC-029 AC-16..18)
// — absent (decimal) needs none, ODF's own default.
const FIELD_NUM_FORMAT: Record<PageNumberFormat, string> = {
  decimal: "",
  lowerRoman: ' style:num-format="i"',
  lowerAlpha: ' style:num-format="a"',
};

/** A page-number/page-count field (SPEC-029 AC-25, AC-26): the placeholder
 * text is the field's cached display value, exactly the shape a real ODF
 * producer's own field elements carry (their content is the last-rendered
 * result, recomputed live by the consuming application). */
// The ODF `text:reference-format` value for a cross-reference's show mode
// (SPEC-053): plain text, the target's page, or its above/below position;
// SPEC-055 adds the caption formats — `category-and-value` (label + number),
// `value` (number only) and `caption` (the text past the label/number).
const REF_FORMAT_ATTR: Record<RefFormat, string> = {
  text: "text",
  page: "page",
  position: "direction",
  full: "text",
  labelAndNumber: "category-and-value",
  numberOnly: "value",
  textOnly: "caption",
};

/** The `text:user-field-decls` block declaring every `docProperty` field's
 * user field (SPEC-053), each value read from the matching SPEC-041 custom
 * property (or "" when the property is gone). */
function userFieldDeclsXml(model: DocumentModel, names: string[]): string {
  if (names.length === 0) {
    return "";
  }
  const decls = names
    .map((name) => {
      const value = model.properties?.customProperties?.find((property) => property.name === name)?.value ?? "";
      return `<text:user-field-decl office:value-type="string" office:string-value="${escapeXml(value)}" text:name="${escapeXml(name)}"/>`;
    })
    .join("");
  return `<text:user-field-decls>${decls}</text:user-field-decls>`;
}

function fieldContentXml(field: FieldSpec, styles: AutomaticStyles): string {
  if (isMergeChipField(field)) {
    // A placeholder carries its own text, needs no declarations, and
    // LibreOffice renders it as a grey placeholder — visually a chip
    // (SPEC-046 "Chip serialisation").
    return `<text:placeholder text:placeholder-type="text">${escapeXml(chipLabel(field))}</text:placeholder>`;
  }
  const cached = isEvaluatedField(field) ? escapeXml(field.cached) : FIELD_PLACEHOLDER_TEXT;
  switch (field.kind) {
    case "page":
      return `<text:page-number${FIELD_NUM_FORMAT[field.format]}>${FIELD_PLACEHOLDER_TEXT}</text:page-number>`;
    case "numpages":
      return `<text:page-count${FIELD_NUM_FORMAT[field.format]}>${FIELD_PLACEHOLDER_TEXT}</text:page-count>`;
    case "date": {
      const entry = dateTimeFormatEntry(field.format);
      const styleAttr = entry ? ` style:data-style-name="${escapeXml(styles.dateStyleNameFor(entry))}"` : "";
      // A locked field is "fixed" in ODF (SPEC-053 rule 4): it is not updated.
      const fixedAttr = field.locked ? ' text:fixed="true"' : "";
      return `<text:date${styleAttr}${fixedAttr}>${cached}</text:date>`;
    }
    case "time": {
      const entry = dateTimeFormatEntry(field.format);
      const styleAttr = entry ? ` style:data-style-name="${escapeXml(styles.dateStyleNameFor(entry))}"` : "";
      const fixedAttr = field.locked ? ' text:fixed="true"' : "";
      return `<text:time${styleAttr}${fixedAttr}>${cached}</text:time>`;
    }
    case "author":
      return `<text:author-name>${cached}</text:author-name>`;
    case "docTitle":
      return `<text:title>${cached}</text:title>`;
    case "docSubject":
      return `<text:subject>${cached}</text:subject>`;
    case "docKeywords":
      return `<text:keywords>${cached}</text:keywords>`;
    case "wordCount":
      return `<text:word-count>${cached}</text:word-count>`;
    case "fileName":
      return `<text:file-name text:display="${field.path ? "full" : "name-and-extension"}">${cached}</text:file-name>`;
    case "docProperty":
      styles.registerUserField(field.name);
      return `<text:user-field-get text:name="${escapeXml(field.name)}">${cached}</text:user-field-get>`;
    case "ref":
      // A caption reference is a `text:sequence-ref` with the matching
      // reference-format (SPEC-055); a heading/bookmark reference stays a
      // `text:bookmark-ref` (SPEC-053).
      if (field.show === "full" || field.show === "labelAndNumber" || field.show === "numberOnly" || field.show === "textOnly") {
        return `<text:sequence-ref text:reference-format="${REF_FORMAT_ATTR[field.show]}" text:ref-name="${escapeXml(field.target)}">${cached}</text:sequence-ref>`;
      }
      return `<text:bookmark-ref text:reference-format="${REF_FORMAT_ATTR[field.show]}" text:ref-name="${escapeXml(field.target)}">${cached}</text:bookmark-ref>`;
    case "pageRef":
      return `<text:bookmark-ref text:reference-format="page" text:ref-name="${escapeXml(field.target)}">${cached}</text:bookmark-ref>`;
    case "seq":
      return `<text:sequence text:name="${escapeXml(field.label)}"${FIELD_NUM_FORMAT[field.format]} text:formula="ooow:${escapeXml(field.label)}+1">${cached}</text:sequence>`;
    case "unsupported":
      // An unsupported field read from an `.odt` keeps its whole source
      // element as `code`, re-emitted verbatim (SPEC-053 byte-for-byte). One
      // read from a `.docx` (code is an instruction string) has no ODF
      // element, so it degrades to a placeholder carrying its saved value;
      // the cross-format save records a DropRecord for it.
      return field.code.startsWith("<")
        ? field.code
        : `<text:placeholder text:placeholder-type="text">${cached}</text:placeholder>`;
    default:
      return cached;
  }
}

// SPEC-031: the notes of the content part being built — set by
// buildContentXml around its body walk so a note reference run anywhere in
// the flow (list items and table cells included) can emit its text:note
// without threading a parameter through every serialiser. Null outside a
// build (and while a note's own body serialises, so a malformed nested
// reference degrades to nothing rather than recursing).
interface OdtNotesWriteContext {
  contentById: Map<string, NoteContent>;
  /** Derived citation text per noteId — the number consumers re-derive. */
  citations: Map<string, string>;
  includeOpaque: boolean;
}
let odtNotesContext: OdtNotesWriteContext | null = null;

function odtNotesWriteContext(model: DocumentModel): OdtNotesWriteContext {
  const contentById = new Map<string, NoteContent>();
  const citations = new Map<string, string>();
  for (const kind of ["footnote", "endnote"] as NoteKind[]) {
    const refs = collectNoteRefs(model, kind);
    const labels = noteLabels(model, kind, refs);
    for (const note of referencedNotes(model, kind, refs)) {
      contentById.set(note.noteId, note);
      citations.set(note.noteId, labels.get(note.noteId) ?? "");
    }
  }
  return { contentById, citations, includeOpaque: model.format === "odt" };
}

/** The model noteId as an XML-safe text:id — model ids (note-UUIDs, docx
 * fn/en ids, ODT source ids) already fit; anything else is sanitised. */
function odtNoteXmlId(noteId: string): string {
  return noteId.replace(/[^A-Za-z0-9._-]/g, "-");
}

// SPEC-038: the comment threads of the content part being built — set by
// buildContentXml around its body walk, mirroring odtNotesContext, so a
// comment-anchor boundary run anywhere in the flow can emit its
// office:annotation/office:annotation-end without threading a parameter
// through every serialiser.
let odtCommentsContext: Map<string, CommentThread> | null = null;

/** One `text:changed-region`'s worth of identity (SPEC-039): enough to emit
 * both the definitions block and every body marker referencing it. Text-level
 * changes reuse lib/trackChanges.ts's `collectChanges` (kind "insert"/
 * "delete"/"format"); structural changes (paragraph mark, table row/cell —
 * AC-16) are walked directly off the model instead, since `collectChanges`'
 * review-pane-shaped `ChangeEntry` collapses a markInsert/markDelete pair (or
 * a row/cell insert/delete pair) into one "paragraphMark"/"row"/"cell" kind
 * with no insert-vs-delete distinction of its own. */
interface OdtRegionDef {
  changeId: string;
  kind: "insert" | "delete" | "format";
  author: string;
  date: string;
}

// SPEC-039: every tracked change the .odt format can represent — text-level
// (insert/delete/format) plus structural (paragraph mark, table row/cell
// insert/delete, AC-16) — in the model being written, keyed by changeId; set
// by buildContentXml around its body walk, mirroring odtCommentsContext, so a
// marked run/paragraph/row/cell anywhere in the flow can emit its
// text:change-start/-end/text:change through this context without threading
// a parameter through every serialiser. The definitions block itself is
// built from the same context. Table/row/cell *property* changes (a merge,
// a border/shading/height/width edit) are not included — ODF has no
// standard slot for a property change's "previous" state at all (the same
// documented gap text-level formatChange already has), so those are
// reported as a SPEC-008 DropRecord instead (AC-19) rather than represented
// here; see collectOdtUnrepresentableTrackedChanges. */
let odtChangedRegionsContext: Map<string, OdtRegionDef> | null = null;

// SPEC-056: the formula write plan for the whole body — set by buildContentXml
// around its body walk, mirroring odtNotesContext, so an equation run anywhere
// in the flow (inline, or the sole run of a display paragraph) emits its
// draw:frame/draw:object referencing the right embedded formula sub-document.
let odtFormulaContext: OdfFormulaWritePlan | null = null;

/** Walks paragraph marks and table row/cell insert/delete throughout
 * `blocks` (recursing into table cells), reporting each one found — the
 * structural half of odtChangedRegionsWriteContext (see its doc comment for
 * why this cannot reuse lib/trackChanges.ts's collectChanges). */
function walkOdtStructuralChanges(
  blocks: Block[],
  note: (info: TrackedChangeInfo | undefined, kind: "insert" | "delete") => void,
): void {
  for (const block of blocks) {
    if (block.type === "paragraph") {
      note(block.changes?.markInsert, "insert");
      note(block.changes?.markDelete, "delete");
      continue;
    }
    if (isTable(block)) {
      for (const row of block.rows) {
        note(row.changes?.insert, "insert");
        note(row.changes?.delete, "delete");
        for (const cell of row.cells) {
          note(cell.changes?.insert, "insert");
          note(cell.changes?.delete, "delete");
          walkOdtStructuralChanges(cell.blocks, note);
        }
      }
    }
  }
}

function odtChangedRegionsWriteContext(model: DocumentModel): Map<string, OdtRegionDef> {
  const byId = new Map<string, OdtRegionDef>();
  for (const entry of collectChanges(model)) {
    if (entry.kind === "insert" || entry.kind === "delete" || entry.kind === "format") {
      byId.set(entry.changeId, { changeId: entry.changeId, kind: entry.kind, author: entry.author, date: entry.date });
    }
  }
  const note = (info: TrackedChangeInfo | undefined, kind: "insert" | "delete") => {
    if (info) {
      byId.set(info.changeId, { changeId: info.changeId, kind, author: info.author, date: info.date });
    }
  };
  for (const section of model.sections) {
    walkOdtStructuralChanges(section.blocks, note);
  }
  return byId;
}

/** The model changeId as an XML-safe text:id — same sanitisation as
 * odtNoteXmlId/odtCommentXmlName. */
function odtChangeXmlId(changeId: string): string {
  return changeId.replace(/[^A-Za-z0-9._-]/g, "-");
}

/** The document's `text:tracked-changes` definitions block (SPEC-039
 * AC-13): one `text:changed-region` per distinct tracked change, each
 * naming its kind and `office:change-info` (author/date) — ODF has no
 * standard slot for a format change's *previous* properties (unlike
 * `.docx`'s `w:rPrChange`), so rejecting an .odt-sourced formatting change
 * falls back to clearing the marker only (documented gap). Empty (and so
 * omitted) for a document with no tracked changes. */
function trackedChangesDefinitionsXml(ctx: Map<string, OdtRegionDef>): string {
  if (ctx.size === 0) {
    return "";
  }
  const regions = Array.from(ctx.values())
    .map((entry) => {
      const tag = entry.kind === "insert" ? "insertion" : entry.kind === "delete" ? "deletion" : "format-change";
      return (
        `<text:changed-region text:id="${odtChangeXmlId(entry.changeId)}">` +
        `<text:${tag}><office:change-info>` +
        `<dc:creator>${escapeXml(entry.author)}</dc:creator><dc:date>${escapeXml(entry.date)}</dc:date>` +
        `</office:change-info></text:${tag}></text:changed-region>`
      );
    })
    .join("");
  return `<text:tracked-changes xmlns:dc="${DC_NAMESPACE}">${regions}</text:tracked-changes>`;
}

/** A run's tracked-change marks (SPEC-039), wrapping its content in
 * `text:change-start`/`-end` pairs — one per insert/delete/formatChange it
 * carries, so a nested change (AC "editing over an existing tracked
 * insertion... produces a nested change") wraps with both. These are point
 * markers, not containers, so nesting order is cosmetic; insert is written
 * outermost to match the .docx writer's w:ins-outside-w:del convention. */
function wrapTrackedChangeMarks(run: Run, inner: string): string {
  const changes = run.changes;
  if (!changes || !odtChangedRegionsContext) {
    return inner;
  }
  let xml = inner;
  const wrap = (mark: TrackedChangeInfo | undefined) => {
    if (!mark || !odtChangedRegionsContext?.has(mark.changeId)) {
      return;
    }
    const id = odtChangeXmlId(mark.changeId);
    xml = `<text:change-start text:change-id="${id}"/>${xml}<text:change-end text:change-id="${id}"/>`;
  };
  wrap(changes.formatChange);
  wrap(changes.delete);
  wrap(changes.insert);
  return xml;
}

function odtCommentsWriteContext(model: DocumentModel): Map<string, CommentThread> {
  const byThreadId = new Map<string, CommentThread>();
  for (const entry of visibleCommentThreads(model)) {
    byThreadId.set(entry.thread.threadId, entry.thread);
  }
  return byThreadId;
}

/** The model threadId as an XML-safe office:name — same sanitisation as
 * odtNoteXmlId. */
function odtCommentXmlName(threadId: string): string {
  return threadId.replace(/[^A-Za-z0-9._-]/g, "-");
}

/** One plain-text message as a bare `<text:p>` (no style reference — ODF
 * defaults an unstyled text:p to "Standard") — the comment body's own
 * lightweight paragraph shape; unlike document body paragraphs there is no
 * automatic style, run formatting or opaque content to carry (comments are
 * plain text only). */
function commentBodyParagraphXml(text: string): string {
  const state: WhitespaceState = { lastChar: "" };
  return `<text:p>${runContentXml(text, state)}</text:p>`;
}

/** A comment thread's `office:annotation` (SPEC-038): `dc:creator`/`dc:date`
 * from the opening message, body one `<text:p>` per message — the opening
 * message's text verbatim, then each reply via lib/comments.ts's
 * `replyParagraphText` — the whole comment inline at its range's start
 * (ODF's own shape, unlike `.docx`'s separate comments.xml part). Emits
 * nothing for a dangling thread id (not in ctx — see visibleCommentThreads)
 * or outside a body build. */
function annotationStartXml(threadId: string): string {
  const thread = odtCommentsContext?.get(threadId);
  if (!thread) {
    return "";
  }
  const opening = thread.messages[0];
  const body = [opening.text, ...thread.messages.slice(1).map(replyParagraphText)]
    .map(commentBodyParagraphXml)
    .join("");
  // xmlns:dc declared locally (the TABLE_XMLNS/LINK_XMLNS convention) so
  // content.xml is unchanged for documents without comments.
  return (
    `<office:annotation xmlns:dc="${DC_NAMESPACE}" office:name="${escapeXml(odtCommentXmlName(threadId))}">` +
    `<dc:creator>${escapeXml(opening.author)}</dc:creator>` +
    `<dc:date>${escapeXml(opening.createdAt)}</dc:date>` +
    body +
    `</office:annotation>`
  );
}

/** A comment thread's `office:annotation-end` (SPEC-038) — emits nothing
 * for a dangling thread id, the same convention as annotationStartXml. */
function annotationEndXml(threadId: string): string {
  if (!odtCommentsContext?.has(threadId)) {
    return "";
  }
  return `<office:annotation-end office:name="${escapeXml(odtCommentXmlName(threadId))}"/>`;
}

/** A note reference run as its `text:note` element (SPEC-031 AC-23/24):
 * citation plus the note's whole body inline at the reference, the ODF
 * shape. Emits nothing for a dangling reference or outside a body build. */
function noteXml(run: Run, styles: AutomaticStyles, state: WhitespaceState): string {
  const ctx = odtNotesContext;
  const note = run.note ? ctx?.contentById.get(run.note.noteId) : undefined;
  if (!ctx || !run.note || !note) {
    return "";
  }
  const citation = ctx.citations.get(run.note.noteId) ?? "";
  odtNotesContext = null;
  const body = note.blocks
    .map((block) =>
      block.type === "paragraph" ? paragraphXml(block, styles, ctx.includeOpaque) : "",
    )
    .join("");
  odtNotesContext = ctx;
  state.lastChar = NOTE_PLACEHOLDER_TEXT;
  return (
    `<text:note text:id="${escapeXml(odtNoteXmlId(run.note.noteId))}" text:note-class="${run.note.kind}">` +
    `<text:note-citation>${escapeXml(citation)}</text:note-citation>` +
    `<text:note-body>${body}</text:note-body></text:note>`
  );
}

function runXml(run: Run, styles: AutomaticStyles, state: WhitespaceState): string {
  if (run.note) {
    return noteXml(run, styles, state);
  }
  const styleName = styles.textNameFor(run.formatting, run.styleId);
  let span: string;
  if (run.equation) {
    // SPEC-056: an inline/display equation is an embedded formula object
    // referenced by an as-char draw:frame (AC-31). The frame is wrapped in a
    // styled span so the run's own formatting still applies; the placeholder
    // character never reaches the output.
    state.lastChar = FIELD_PLACEHOLDER_TEXT;
    const href = odtFormulaContext?.hrefFor(run);
    const frame = href ? formulaFrameXml(run.equation, href, formulaFrameAltText(run)) : "";
    span = `<text:span text:style-name="${styleName}">${frame}</text:span>`;
  } else if (run.field) {
    state.lastChar = FIELD_PLACEHOLDER_TEXT;
    span = `<text:span text:style-name="${styleName}">${fieldContentXml(run.field, styles)}</text:span>`;
  } else {
    span = `<text:span text:style-name="${styleName}">${runContentXml(run.text, state)}</text:span>`;
  }
  return wrapTrackedChangeMarks(run, span);
}

/** The paragraph's inline content: its runs, with any preserved opaque
 * inline nodes re-emitted verbatim at their anchored positions (SPEC-008),
 * consecutive same-target link runs grouped under one `text:a` (SPEC-025),
 * and bookmarks emitted as point `text:bookmark` elements at their
 * character offsets. The whitespace state passes through opaque nodes,
 * link boundaries and bookmarks unchanged — elements are transparent to
 * ODF whitespace collapsing, so the reader's view matches. Links and
 * bookmarks are modelled content, emitted on cross-format saves too. */
function inlineContentXml(
  paragraph: Paragraph,
  styles: AutomaticStyles,
  includeOpaque: boolean,
): string {
  const state: WhitespaceState = { lastChar: "" };
  const opaque = includeOpaque ? (paragraph.opaqueInlines ?? []) : [];
  const hasLinks = paragraph.runs.some((run) => linkTargetOf(run.formatting) !== null);
  const hasBookmarks = (paragraph.bookmarks ?? []).length > 0;
  const hasComments = (paragraph.commentAnchors ?? []).length > 0;
  if (opaque.length === 0 && !hasLinks && !hasBookmarks && !hasComments) {
    return paragraph.runs.map((run) => runXml(run, styles, state)).join("");
  }
  const { runs, bookmarksAtBoundary, commentStartsAtBoundary, commentEndsAtBoundary, boundaryFor } =
    inlineWriteLayout(paragraph);
  const opaqueAtBoundary = new Map<number, string[]>();
  for (const node of opaque) {
    const boundary = boundaryFor(node.runIndex);
    const list = opaqueAtBoundary.get(boundary) ?? [];
    list.push(node.xml);
    opaqueAtBoundary.set(boundary, list);
  }
  let content = "";
  let openTarget: LinkTarget | null = null;
  const closeLink = () => {
    if (openTarget) {
      content += "</text:a>";
      openTarget = null;
    }
  };
  for (let i = 0; i <= runs.length; i++) {
    const run = i < runs.length ? runs[i] : null;
    const target = run ? linkTargetOf(run.formatting) : null;
    const opaqueHere = opaqueAtBoundary.get(i) ?? [];
    // An opaque node's XML is arbitrary — never emit it inside a text:a;
    // the link is split around it instead (two same-target text:a elements
    // are equivalent).
    if (opaqueHere.length > 0) {
      closeLink();
    }
    if (
      openTarget &&
      (!target || target.href !== openTarget.href || target.anchor !== openTarget.anchor)
    ) {
      closeLink();
    }
    content += opaqueHere.join("");
    for (const name of bookmarksAtBoundary.get(i) ?? []) {
      content += `<text:bookmark text:name="${escapeXml(name)}"/>`;
    }
    // Comment ranges (SPEC-038): office:annotation-end must precede
    // annotation starts landing at the very same boundary (closing a
    // finished range before opening the next keeps well-formed nesting),
    // so ends are emitted first.
    for (const threadId of commentEndsAtBoundary.get(i) ?? []) {
      content += annotationEndXml(threadId);
    }
    for (const threadId of commentStartsAtBoundary.get(i) ?? []) {
      content += annotationStartXml(threadId);
    }
    if (!run) {
      break;
    }
    if (target && !openTarget) {
      const href = target.href !== undefined ? target.href : `#${target.anchor ?? ""}`;
      content += `<text:a xlink:type="simple" xlink:href="${escapeXml(href)}">`;
      openTarget = target;
    }
    content += runXml(run, styles, state);
  }
  closeLink();
  return content;
}

// --- table of contents (SPEC-033) -------------------------------------------

function tocEntryTemplateXml(level: number): string {
  return (
    `<text:table-of-content-entry-template text:outline-level="${level}" text:style-name="Contents_20_${level}">` +
    "<text:index-entry-link-start/>" +
    "<text:index-entry-text/>" +
    '<text:index-entry-tab-stop style:type="right" style:leader-char="."/>' +
    "<text:index-entry-page-number/>" +
    "<text:index-entry-link-end/>" +
    "</text:table-of-content-entry-template>"
  );
}

/** The `text:table-of-content-source` config (SPEC-033 AC-5, AC-18):
 * LibreOffice/Word read this to know which outline levels and entry shape
 * "Update index" should regenerate — this app's own reader ignores it
 * (entries round-trip from the index body itself, see reader.ts's
 * parseTableOfContent), but it is what makes the file's own "regenerate"
 * command available in a real word processor too. */
function tocSourceXml(options: TocOptions): string {
  const levels = options.levels.length > 0 ? options.levels : [1];
  const maxLevel = Math.max(...levels);
  const templates = levels.map((level) => tocEntryTemplateXml(level)).join("");
  return (
    `<text:table-of-content-source text:outline-level="${maxLevel}" text:use-outline-level="true">` +
    '<text:index-title-template text:style-name="Contents_20_Heading"/>' +
    templates +
    "</text:table-of-content-source>"
  );
}

/** One TOC entry as an index-body paragraph: its heading text as an
 * internal hyperlink (AC-8, AC-9 — every entry is clickable regardless of
 * options), plus, unless `hyperlinksOnly`, a tab and the cached page number
 * (AC-2..4, AC-7). */
function tocEntryParagraphXml(entry: TocEntry, options: TocOptions): string {
  const level = Math.max(1, Math.min(6, entry.level));
  const trailer =
    options.showPageNumbers && !options.hyperlinksOnly
      ? `<text:tab/>${escapeXml(String(entry.pageNumber ?? ""))}`
      : "";
  return (
    `<text:p text:style-name="Contents_20_${level}"${LINK_XMLNS}>` +
    `<text:a xlink:type="simple" xlink:href="#${escapeXml(entry.anchor)}">${escapeXml(entry.text)}${trailer}</text:a>` +
    "</text:p>"
  );
}

/** A TOC block as `text:table-of-content` (SPEC-033 AC-18): a source config
 * plus an index body holding the cached entries (or the empty-state
 * placeholder) as plain paragraphs. `text:protected="true"` mirrors the
 * block's own "not directly editable" contract in the format itself — a
 * real ODF consumer refuses to type into a protected index the same way
 * this app's own canvas refuses to (SPEC-033 AC-14/15/24). */
function tocXml(toc: TableOfContents): string {
  const body =
    toc.entries.length === 0
      ? `<text:p text:style-name="Contents_20_1">${escapeXml(TOC_PLACEHOLDER_TEXT)}</text:p>`
      : toc.entries.map((entry) => tocEntryParagraphXml(entry, toc.options)).join("");
  return (
    '<text:table-of-content text:protected="true" text:name="TableOfContents1">' +
    tocSourceXml(toc.options) +
    `<text:index-body>${body}</text:index-body>` +
    "</text:table-of-content>"
  );
}

// --- table of figures (SPEC-055) --------------------------------------------

/** The ODF index element name a label maps to: a "Table" label indexes tables
 * (`text:table-index`), every other label indexes illustrations
 * (`text:illustration-index`) — Writer's own split for figure vs table lists. */
function tofIndexTag(label: string): string {
  return label.toLowerCase() === "table" ? "text:table-index" : "text:illustration-index";
}

/** The paragraph style ODF entries carry, per index kind. */
function tofEntryStyle(label: string): string {
  return label.toLowerCase() === "table" ? "Table_20_index_20_1" : "Illustration_20_Index_20_1";
}

/** One ToF entry as an index-body paragraph: the caption text as an internal
 * hyperlink, plus (unless hyperlinks/page-numbers are off) a tab and the
 * cached page number. Mirrors `tocEntryParagraphXml`. */
function tofEntryParagraphXml(entry: TofEntry, tof: TableOfFigures): string {
  const style = tofEntryStyle(tof.label);
  const trailer =
    tof.options.pageNumbers && tof.options.hyperlinks
      ? `<text:tab/>${escapeXml(String(entry.page ?? ""))}`
      : "";
  return (
    `<text:p text:style-name="${style}"${LINK_XMLNS}>` +
    `<text:a xlink:type="simple" xlink:href="#${escapeXml(entry.targetBookmark)}">${escapeXml(entry.text)}${trailer}</text:a>` +
    "</text:p>"
  );
}

/** A table of figures as `text:illustration-index` / `text:table-index`
 * (SPEC-055), mirroring `tocXml`: a source config carrying the label as its
 * caption sequence plus an index body holding the cached entries. */
function tofXml(tof: TableOfFigures): string {
  const tag = tofIndexTag(tof.label);
  const sourceTag = `${tag}-source`;
  const body =
    tof.entries.length === 0
      ? `<text:p text:style-name="${tofEntryStyle(tof.label)}">${escapeXml(TOF_PLACEHOLDER_TEXT)}</text:p>`
      : tof.entries.map((entry) => tofEntryParagraphXml(entry, tof)).join("");
  return (
    `<${tag} text:protected="true" text:name="${escapeXml(tof.label)}Index1">` +
    `<${sourceTag} text:caption-sequence-name="${escapeXml(tof.label)}"/>` +
    `<text:index-body>${body}</text:index-body>` +
    `</${tag}>`
  );
}

/** A pending paragraph-mark change's own tracked-change marker (SPEC-039
 * AC-16): a point `text:change` appended after the paragraph's own inline
 * content. ODF has no `w:pPr/w:rPr/w:ins`-equivalent slot for "this
 * paragraph's own trailing mark is pending", so the marker rides at the end
 * of the paragraph's own content instead — unambiguous against a row/cell
 * marker (see rowChangeMarkerXml/cellChangeMarkerXml), which always LEADS a
 * paragraph's content, never trails one. */
function paragraphMarkChangeXml(paragraph: Paragraph): string {
  if (!odtChangedRegionsContext) {
    return "";
  }
  const parts: string[] = [];
  const markInsert = paragraph.changes?.markInsert;
  const markDelete = paragraph.changes?.markDelete;
  if (markInsert && odtChangedRegionsContext.has(markInsert.changeId)) {
    parts.push(`<text:change text:change-id="${odtChangeXmlId(markInsert.changeId)}"/>`);
  }
  if (markDelete && odtChangedRegionsContext.has(markDelete.changeId)) {
    parts.push(`<text:change text:change-id="${odtChangeXmlId(markDelete.changeId)}"/>`);
  }
  return parts.join("");
}

/** A paragraph as text:p — or text:h with text:outline-level when the model
 * carries a heading level, so the level survives re-opening regardless of
 * the style's definition. `leadingXml` is a table row/cell's own tracked-
 * change marker (SPEC-039 AC-16), prepended ahead of the paragraph's own
 * content when this is a cell's first paragraph (see cellContentXml). */
function paragraphXml(
  paragraph: Paragraph,
  styles: AutomaticStyles,
  includeOpaque: boolean,
  styleOptions?: ParagraphStyleOptions,
  leadingXml = "",
  trailingXml = "",
): string {
  // `trailingXml` (SPEC-078): draw elements of floating shapes/charts
  // absorbed into this, their anchor paragraph — after the paragraph's own
  // content when the drawing follows its anchor in the block list (the
  // leading side rides in via `leadingXml`), so the reader's split
  // reproduces the model's block order.
  const styleName = styles.paragraphNameFor(paragraph, styleOptions);
  const tag = paragraph.headingLevel !== undefined ? "text:h" : "text:p";
  let attrs = styleName !== undefined ? ` text:style-name="${escapeXml(styleName)}"` : "";
  if (paragraph.headingLevel !== undefined) {
    attrs += ` text:outline-level="${paragraph.headingLevel}"`;
  }
  if (paragraph.runs.some((run) => linkTargetOf(run.formatting) !== null)) {
    attrs += LINK_XMLNS;
  }
  const content =
    leadingXml + inlineContentXml(paragraph, styles, includeOpaque) + paragraphMarkChangeXml(paragraph) + trailingXml;
  return content.length > 0 ? `<${tag}${attrs}>${content}</${tag}>` : `<${tag}${attrs}/>`;
}

function isListParagraph(block: Block): boolean {
  return block.type === "paragraph" && block.list !== undefined;
}

/** An empty paragraph the sole purpose of which is to carry a SPEC-028
 * style: a trailing manual break (nothing follows to attach fo:break-before
 * to) or a break/section boundary immediately before a table (table:table
 * takes no paragraph style). Re-opens as the break/section boundary plus an
 * empty paragraph — a documented convergence; the second cycle is a fixed
 * point. */
function carrierParagraphXml(styles: AutomaticStyles, styleOptions: ParagraphStyleOptions): string {
  const name = styles.paragraphNameFor({ type: "paragraph", runs: [] }, styleOptions);
  return name !== undefined ? `<text:p text:style-name="${escapeXml(name)}"/>` : "<text:p/>";
}

/** An Image block (SPEC-024): ODF has no body-level picture element either,
 * so it is wrapped the way real producers do — one `text:p` holding the
 * `draw:frame` (see drawing.ts's drawFrameXml). Returns "" when `mediaPlan`
 * has no href for this image's media (defensive; every referenced media id
 * should always have one — ODF, unlike .docx, has no format it can't
 * store). A pending break/section start (SPEC-028) rides on the wrapper
 * paragraph's style, so it round-trips without an extra carrier paragraph
 * (the wrapper contributes no paragraph of its own on re-open). */
function imageParagraphXml(
  image: Image,
  styles: AutomaticStyles,
  mediaPlan: OdtMediaWritePlan,
  styleOptions?: ParagraphStyleOptions,
): string {
  const href = mediaPlan.hrefFor(image.mediaId);
  if (!href) {
    // Keep the break/section boundary alive even without the image.
    return styleOptions ? carrierParagraphXml(styles, styleOptions) : "";
  }
  const wrapperName = styleOptions
    ? styles.paragraphNameFor({ type: "paragraph", runs: [] }, styleOptions)
    : undefined;
  const wrapperAttr = wrapperName !== undefined ? ` text:style-name="${escapeXml(wrapperName)}"` : "";
  const graphicStyleName = styles.graphicStyleNameFor(image.wrap);
  return `<text:p${wrapperAttr}>${drawFrameXml(image, href, graphicStyleName)}</text:p>`;
}

/** A Shape block (SPEC-051): the shape-flavoured sibling of
 * imageParagraphXml — one `text:p` holding the `draw:custom-shape`/
 * `draw:line` (see shape.ts's shapeElementXml). Its text frame's paragraphs
 * are serialised through the same paragraphXml every body paragraph uses. */
function shapeParagraphXml(
  shape: Shape,
  styles: AutomaticStyles,
  includeOpaque: boolean,
  styleOptions?: ParagraphStyleOptions,
  drawId?: string,
): string {
  const wrapperName = styleOptions
    ? styles.paragraphNameFor({ type: "paragraph", runs: [] }, styleOptions)
    : undefined;
  const wrapperAttr = wrapperName !== undefined ? ` text:style-name="${escapeXml(wrapperName)}"` : "";
  const styleName = styles.shapeStyleNameFor(shape);
  const element = shapeElementXml(
    shape,
    styleName,
    { paragraphXml: (paragraph) => paragraphXml(paragraph, styles, includeOpaque) },
    drawId,
  );
  return `<text:p${wrapperAttr}>${element}</text:p>`;
}

/** A Chart block (SPEC-052): the chart-flavoured sibling of
 * imageParagraphXml/shapeParagraphXml — one `text:p` holding the
 * `draw:frame`/`draw:object` (see chart.ts's chartFrameXml). Returns just
 * the carrier paragraph (or "") when `chartPlan` has no entry for this
 * location (defensive; every Chart block always gets one from
 * planOdfChartWrite). */
function chartParagraphXml(
  chart: Chart,
  styles: AutomaticStyles,
  chartPlan: OdfChartWritePlan,
  section: number,
  block: number,
  styleOptions?: ParagraphStyleOptions,
): string {
  const wrapperName = styleOptions
    ? styles.paragraphNameFor({ type: "paragraph", runs: [] }, styleOptions)
    : undefined;
  const wrapperAttr = wrapperName !== undefined ? ` text:style-name="${escapeXml(wrapperName)}"` : "";
  const entry = chartPlan.planFor(section, block);
  if (!entry) {
    return styleOptions ? carrierParagraphXml(styles, styleOptions) : "";
  }
  return `<text:p${wrapperAttr}>${chartFrameXml(chart, `./${entry.objectDir}`)}</text:p>`;
}

/** A cell's block content (SPEC-022): paragraphs, with consecutive list
 * items grouped into text:list runs exactly like body content — a cell is
 * its own list context, so runs never join one outside the table.
 * `leadingMarkerXml` is a pending row/cell tracked-change marker (SPEC-039
 * AC-16), prepended into the cell's first paragraph — dropped if that
 * paragraph turns out to be a list item (text:list has no slot for it; a
 * documented gap, the same shape as the cell/row-property DropRecord one). */
function cellContentXml(blocks: Paragraph[], styles: AutomaticStyles, includeOpaque: boolean, leadingMarkerXml = ""): string {
  // The cell is its own numbering context (SPEC-022), so instance
  // continuity never crosses the cell boundary — a fresh map per cell.
  const continuity = new Map<string, string>();
  let xml = "";
  let i = 0;
  while (i < blocks.length) {
    if (!blocks[i].list) {
      xml += paragraphXml(blocks[i], styles, includeOpaque, undefined, i === 0 ? leadingMarkerXml : "");
      i++;
      continue;
    }
    let end = i + 1;
    while (end < blocks.length && blocks[end].list) {
      end++;
    }
    xml += listRunsXml(blocks.slice(i, end), styles, includeOpaque, undefined, continuity);
    i = end;
  }
  if (xml.length > 0) {
    return xml;
  }
  return leadingMarkerXml ? `<text:p>${leadingMarkerXml}</text:p>` : "<text:p/>";
}

/** A table row's own pending insert/delete tracked-change marker (SPEC-039
 * AC-16): a point `text:change` leading its first (non-covered) cell's
 * first paragraph. `table:change-scope="row"` distinguishes it from a
 * cell-level marker that could otherwise land at the exact same position
 * (that same first cell, if it independently carries its own cell-level
 * change too — see cellChangeMarkerXml, which never sets this attribute). */
function rowChangeMarkerXml(row: TableRow): string {
  if (!odtChangedRegionsContext) {
    return "";
  }
  const mark = row.changes?.insert ?? row.changes?.delete;
  if (!mark || !odtChangedRegionsContext.has(mark.changeId)) {
    return "";
  }
  return `<text:change text:change-id="${odtChangeXmlId(mark.changeId)}" table:change-scope="row"/>`;
}

/** A table cell's own pending insert/delete tracked-change marker (SPEC-039
 * AC-16: a column insert/delete), leading its own first paragraph. */
function cellChangeMarkerXml(cell: TableCell): string {
  if (!odtChangedRegionsContext) {
    return "";
  }
  const mark = cell.changes?.insert ?? cell.changes?.delete;
  if (!mark || !odtChangedRegionsContext.has(mark.changeId)) {
    return "";
  }
  return `<text:change text:change-id="${odtChangeXmlId(mark.changeId)}"/>`;
}

/** One table block as table:table markup (SPEC-022; formatting and merged
 * cells SPEC-023): column definitions referencing width-carrying styles,
 * then rows of cells. The table: namespace is declared on the element
 * itself — see TABLE_XMLNS.
 *
 * A grid slot a merged cell covers (SPEC-023 AC-28) emits
 * `<table:covered-table-cell/>` instead of `<table:table-cell>`; the real
 * cell carries `table:number-columns-spanned`/`table:number-rows-spanned`. */
function tableXml(table: Table, styles: AutomaticStyles, includeOpaque: boolean): string {
  const names = styles.tableNamesFor(table, includeOpaque);
  const alignAttr = table.alignment !== undefined ? ` table:align="${table.alignment}"` : "";
  // SPEC-076 AC-68: the banding style the table points at, and which of its
  // regions this table uses. `table:style-name` already carries the
  // width-bearing automatic style, so the NAMED style rides on
  // table:template-name — which is also what makes it come back on read.
  const templateAttr =
    table.styleId !== undefined ? ` table:template-name="${escapeXml(table.styleId)}"` : "";
  let xml =
    `<table:table${TABLE_XMLNS} table:name="${escapeXml(names.name)}"` +
    ` table:style-name="${escapeXml(names.styleName)}"${alignAttr}${templateAttr}` +
    `${tableUseStyleAttrs(table.look)}>`;
  for (const columnName of names.columnStyleNames) {
    xml += `<table:table-column table:style-name="${escapeXml(columnName)}"/>`;
  }
  const ownerGrid = computeOwnerGrid(table);
  // The leading contiguous run of isHeader rows (SPEC-043) becomes ODF's
  // table:table-header-rows wrapper — a structural container, not a per-row
  // attribute, so only a leading run is representable; an isHeader row after
  // a non-header one has nothing to wrap it and is written as a plain row.
  let headerRunLength = 0;
  while (headerRunLength < table.rows.length && table.rows[headerRunLength].isHeader) {
    headerRunLength += 1;
  }
  if (headerRunLength > 0) {
    xml += "<table:table-header-rows>";
  }
  table.rows.forEach((row, rowIndex) => {
    const rowStyleName =
      includeOpaque && row.odtStyleName !== undefined ? row.odtStyleName : names.rowStyleName(row);
    const rowStyleAttr = rowStyleName !== undefined ? ` table:style-name="${escapeXml(rowStyleName)}"` : "";
    xml += `<table:table-row${rowStyleAttr}>`;
    let firstRealCellInRow = true;
    for (let colIndex = 0; colIndex < table.columnWidthsCm.length; colIndex++) {
      const owner = ownerGrid[rowIndex][colIndex];
      if (owner.originRow !== rowIndex || owner.originCol !== colIndex) {
        xml += "<table:covered-table-cell/>";
        continue;
      }
      const cell = row.cells[colIndex];
      // A preserved cell keeps exactly the style reference it came with —
      // including none at all; anything else gets a style generated from
      // its own SPEC-023 properties, or no style at all if it sets none.
      const cellStyleName =
        includeOpaque && cell.odtStyleName !== undefined ? cell.odtStyleName : names.cellStyleName(cell);
      const cellStyleAttr = cellStyleName !== undefined ? ` table:style-name="${escapeXml(cellStyleName)}"` : "";
      const spanAttrs =
        (owner.colSpan > 1 ? ` table:number-columns-spanned="${owner.colSpan}"` : "") +
        (owner.rowSpan > 1 ? ` table:number-rows-spanned="${owner.rowSpan}"` : "");
      // SPEC-039 AC-16: a pending row change marks its first real cell; a
      // pending cell change marks its own — both can coexist on the row's
      // own first cell (table:change-scope tells them apart on read).
      const marker = (firstRealCellInRow ? rowChangeMarkerXml(row) : "") + cellChangeMarkerXml(cell);
      firstRealCellInRow = false;
      xml +=
        `<table:table-cell${cellStyleAttr}${spanAttrs} office:value-type="string">` +
        cellContentXml(cell.blocks, styles, includeOpaque, marker) +
        "</table:table-cell>";
    }
    xml += "</table:table-row>";
    if (rowIndex === headerRunLength - 1) {
      xml += "</table:table-header-rows>";
    }
  });
  return `${xml}</table:table>`;
}

/** The content part: automatic styles plus the document body. All sections'
 * blocks are written in order (page setup lives on defaultStyles, not per
 * section); a model with no blocks still gets one empty text:p, mirroring
 * the .docx writer.
 *
 * SPEC-008: when the model came from a .odt, preserved opaque blocks and
 * inlines are re-emitted verbatim at their anchored positions, along with
 * the original font-face declarations and the non-paragraph/text automatic
 * styles those nodes may reference. Cross-format models carry foreign
 * (WordprocessingML) XML in their opaque nodes, so those are excluded — the
 * writer reports them as dropped instead. */
export function buildContentXml(
  model: DocumentModel,
  mediaPlan: OdtMediaWritePlan = planOdtMediaWrite(model),
  masterNames?: (string | null)[],
  chartPlan: OdfChartWritePlan = planOdfChartWrite(model, new Set()),
  formulaPlan: OdfFormulaWritePlan = planOdfFormulaWrite(model),
  drawingDefs: OdfDrawingDefs = createOdfDrawingDefs(),
): string {
  const includeOpaque = model.format === "odt";
  // SPEC-073: shape/connector graphic styles reference shared drawing
  // definitions (gradients, hatches, fill images, markers, dash patterns)
  // through this registry; the caller (writeOdt) reads the collected
  // definitions back out for styles.xml once the content is built. Preserved
  // unmodelled styling is only re-emitted on a same-format save (AC-35) — a
  // cross-format model's unsupportedStyleXml is the other format's
  // vocabulary, reported centrally as dropped instead.
  const shapeStyleCtx: OdfShapeStyleWriteContext = {
    defs: drawingDefs,
    hrefForMedia: (mediaId) => mediaPlan.hrefFor(mediaId) || undefined,
    includePreserved: includeOpaque,
  };
  const styles = createAutomaticStyles(model, shapeStyleCtx);
  const blockXml: string[] = [];
  // Footnotes/endnotes (SPEC-031): note reference runs anywhere in the body
  // emit their text:note through this context; cleared before returning.
  odtNotesContext = odtNotesWriteContext(model);
  // Comments (SPEC-038): comment-anchor boundary runs anywhere in the body
  // emit their office:annotation/-end through this context; cleared with
  // odtNotesContext below.
  odtCommentsContext = odtCommentsWriteContext(model);
  // Tracked changes (SPEC-039): a marked run anywhere in the body emits its
  // text:change-start/-end through this context; cleared with the others
  // below. The definitions block itself is built from the same context.
  odtChangedRegionsContext = odtChangedRegionsWriteContext(model);
  // Equations (SPEC-056): an equation run anywhere in the body emits its
  // draw:frame/draw:object through this plan; cleared with the others below.
  odtFormulaContext = formulaPlan;
  // Defined-list instance continuity for the whole body (SPEC-032): one map
  // across every section, so an instance interrupted by a section break
  // still writes text:continue-list on its later segments.
  const bodyListContinuity = new Map<string, string>();
  // SPEC-028 AC-32: ODF has no standalone break element — a BreakBlock
  // becomes fo:break-before on the NEXT emitted paragraph's automatic
  // style, so the break is held pending until that paragraph is reached.
  let pendingBreak: BreakKind | undefined;
  model.sections.forEach((section, sectionIndex) => {
    // SPEC-070: block indices any connector in this section attaches to — those
    // shapes are written with a draw:id so the connector's draw:start-shape/
    // draw:end-shape reference resolves on reopen.
    const connectorTargets = new Set<number>();
    for (const block of section.blocks) {
      if (block.type === "connector") {
        if (block.startAttachment) {
          connectorTargets.add(block.startAttachment.target.block);
        }
        if (block.endAttachment) {
          connectorTargets.add(block.endAttachment.target.block);
        }
      }
    }
    const drawIdFor = (blockIndex: number): string | undefined =>
      connectorTargets.has(blockIndex) ? odfDrawId(sectionIndex, blockIndex) : undefined;
    // SPEC-028 AC-33: every section after the first switches to its own
    // master page via style:master-page-name on its first paragraph's
    // automatic style — in ODF that reference IS the section break (page
    // break + layout switch). Emitted for every later section regardless of
    // `start`: ODF has no continuous-section concept, so a "continuous"
    // section re-opens as nextPage — a documented fidelity loss that keeps
    // the section boundary itself alive.
    // The plan is authoritative when given: a null entry means the section's
    // generated master collided with a foreign definition and was skipped by
    // the patch, so content.xml must not reference it either.
    let pendingMaster =
      sectionIndex > 0
        ? masterNames
          ? (masterNames[sectionIndex] ?? undefined)
          : sectionMasterName(sectionIndex)
        : undefined;
    // SPEC-029 AC-22: a section's own restart value rides on the same
    // first-paragraph carrier as its master-page switch — for section 0,
    // which never switches master, it is the ONLY reason that paragraph's
    // automatic style would carry anything section-related at all.
    let pendingPageNumberStart = section.pageNumberStart;
    const takeStyleOptions = (): ParagraphStyleOptions | undefined => {
      if (pendingBreak === undefined && pendingMaster === undefined && pendingPageNumberStart === undefined) {
        return undefined;
      }
      const options: ParagraphStyleOptions = {};
      if (pendingBreak !== undefined) {
        options.breakBefore = pendingBreak;
      }
      if (pendingMaster !== undefined) {
        options.masterPageName = pendingMaster;
      }
      if (pendingPageNumberStart !== undefined) {
        options.pageNumberStart = pendingPageNumberStart;
      }
      pendingBreak = undefined;
      pendingMaster = undefined;
      pendingPageNumberStart = undefined;
      return options;
    };
    const opaqueBlocks = includeOpaque
      ? [...(section.opaqueBlocks ?? [])].sort((a, b) => a.blockIndex - b.blockIndex)
      : [];
    const blocks = section.blocks;
    // Containment (SPEC-078): floating shapes/charts with an explicit,
    // resolvable anchor pairing are emitted inside their anchor paragraph's
    // text:p — how text:anchor-type="paragraph" natively expresses the
    // anchor — and skipped at their own block slot; same plan as the .docx
    // writer. ODT's carve-out: a drawing anchored to a LIST paragraph keeps
    // the legacy own-paragraph form, since its paragraph is emitted inside
    // a text:list run with no seam for a foreign element (SPEC-078's out of
    // scope; the pairing is re-derived by adjacency on open).
    const anchoredPlan = planAnchoredDrawings(blocks);
    for (const [paragraphIndex, entry] of anchoredPlan.byParagraph) {
      const anchorBlock = blocks[paragraphIndex];
      if (anchorBlock.type === "paragraph" && anchorBlock.list) {
        for (const drawingIndex of [...entry.leading, ...entry.trailing]) {
          anchoredPlan.absorbed.delete(drawingIndex);
        }
        anchoredPlan.byParagraph.delete(paragraphIndex);
      }
    }
    // ODF's draw:start-shape/end-shape can only glue to a shape (the ODF
    // writer mints a draw:id for shapes, not images/charts); a connector
    // attached to a non-shape drawing writes no ref and renders free — the
    // lossy case collectPackageDropRecords reports (SPEC-070 AC-32).
    const shapeTargetId = (blockIndex: number | undefined): string | undefined => {
      if (blockIndex === undefined) {
        return undefined;
      }
      return section.blocks[blockIndex]?.type === "shape" ? odfDrawId(sectionIndex, blockIndex) : undefined;
    };
    const connectorElementFor = (connector: Connector, drawingIndex: number): string =>
      connectorElementXml(connector, styles.connectorStyleNameFor(connector), {
        self: drawIdFor(drawingIndex),
        start: connector.startAttachment ? shapeTargetId(connector.startAttachment.target.block) : undefined,
        end: connector.endAttachment ? shapeTargetId(connector.endAttachment.target.block) : undefined,
      });
    const anchoredElementXml = (drawingIndex: number): string => {
      const drawing = blocks[drawingIndex];
      if (drawing.type === "shape") {
        return shapeElementXml(
          drawing,
          styles.shapeStyleNameFor(drawing),
          { paragraphXml: (paragraph) => paragraphXml(paragraph, styles, includeOpaque) },
          drawIdFor(drawingIndex),
        );
      }
      if (drawing.type === "connector") {
        return connectorElementFor(drawing, drawingIndex);
      }
      if (drawing.type === "chart") {
        const chartEntry = chartPlan.planFor(sectionIndex, drawingIndex);
        return chartEntry ? chartFrameXml(drawing, `./${chartEntry.objectDir}`) : "";
      }
      if (drawing.type === "image") {
        // A floating picture absorbs into its anchor paragraph like every
        // other floating drawing — it joined them when Image gained real
        // placement. Missing this case dropped the picture entirely: the
        // plan skipped its own block slot and nothing emitted it here.
        const href = mediaPlan.hrefFor(drawing.mediaId);
        return href ? drawFrameXml(drawing, href, styles.graphicStyleNameFor(drawing.wrap)) : "";
      }
      return "";
    };
    const anchorsAt = (index: number): typeof opaqueBlocks =>
      opaqueBlocks.filter((node) => Math.min(node.blockIndex, blocks.length) === index);
    let i = 0;
    while (i <= blocks.length) {
      for (const node of anchorsAt(i)) {
        blockXml.push(node.xml);
      }
      if (i >= blocks.length) {
        break;
      }
      // Serialised before the styles are read: paragraphXml/listRunXml
      // register each automatic style on first use.
      const block = blocks[i];
      if (block.type === "table") {
        // table:table takes no paragraph style, so a pending break/section
        // boundary needs an empty carrier paragraph before it (SPEC-028).
        const options = takeStyleOptions();
        if (options) {
          blockXml.push(carrierParagraphXml(styles, options));
        }
        blockXml.push(tableXml(block, styles, includeOpaque));
        i++;
        continue;
      }
      if (block.type === "image") {
        blockXml.push(imageParagraphXml(block, styles, mediaPlan, takeStyleOptions()));
        i++;
        continue;
      }
      if (block.type === "shape") {
        // An absorbed drawing (SPEC-078) is emitted inside its anchor
        // paragraph via anchoredElementXml; a pending break/master rides on
        // to the next emitted paragraph untaken.
        if (!anchoredPlan.absorbed.has(i)) {
          blockXml.push(shapeParagraphXml(block, styles, includeOpaque, takeStyleOptions(), drawIdFor(i)));
        }
        i++;
        continue;
      }
      if (block.type === "connector") {
        if (!anchoredPlan.absorbed.has(i)) {
          const options = takeStyleOptions();
          const wrapperName = options
            ? styles.paragraphNameFor({ type: "paragraph", runs: [] }, options)
            : undefined;
          const wrapperAttr = wrapperName !== undefined ? ` text:style-name="${escapeXml(wrapperName)}"` : "";
          blockXml.push(`<text:p${wrapperAttr}>${connectorElementFor(block, i)}</text:p>`);
        }
        i++;
        continue;
      }
      if (block.type === "chart") {
        if (!anchoredPlan.absorbed.has(i)) {
          blockXml.push(chartParagraphXml(block, styles, chartPlan, sectionIndex, i, takeStyleOptions()));
        }
        i++;
        continue;
      }
      if (block.type === "break") {
        // SPEC-028 AC-32: held until the next paragraph (see pendingBreak);
        // the break block itself emits nothing here.
        pendingBreak = block.kind;
        i++;
        continue;
      }
      if (block.type === "toc") {
        // text:table-of-content takes no paragraph style either, same as a
        // table (SPEC-028's pending break/section boundary needs a carrier).
        const options = takeStyleOptions();
        if (options) {
          blockXml.push(carrierParagraphXml(styles, options));
        }
        blockXml.push(tocXml(block));
        i++;
        continue;
      }
      if (block.type === "tof") {
        const options = takeStyleOptions();
        if (options) {
          blockXml.push(carrierParagraphXml(styles, options));
        }
        blockXml.push(tofXml(block));
        i++;
        continue;
      }
      if (!block.list) {
        const anchoredEntry = anchoredPlan.byParagraph.get(i);
        blockXml.push(
          paragraphXml(
            block,
            styles,
            includeOpaque,
            takeStyleOptions(),
            anchoredEntry ? anchoredEntry.leading.map(anchoredElementXml).join("") : "",
            anchoredEntry ? anchoredEntry.trailing.map(anchoredElementXml).join("") : "",
          ),
        );
        i++;
        continue;
      }
      // A run of consecutive list paragraphs becomes one nested text:list
      // (SPEC-018 AC-19). An opaque block anchored inside the run ends it —
      // the preserved node belongs at the body level, and the following
      // items start a fresh list. A table block ends the run the same way.
      let end = i + 1;
      while (end < blocks.length && isListParagraph(blocks[end]) && anchorsAt(end).length === 0) {
        end++;
      }
      blockXml.push(
        listRunsXml(
          blocks.slice(i, end) as Paragraph[],
          styles,
          includeOpaque,
          takeStyleOptions(),
          bodyListContinuity,
        ),
      );
      i = end;
    }
    if (pendingMaster !== undefined || pendingPageNumberStart !== undefined) {
      // A section with no paragraph to carry its master-page reference or
      // restart value (empty): an empty carrier paragraph keeps it alive
      // (SPEC-028 AC-33, SPEC-029 AC-22).
      blockXml.push(carrierParagraphXml(styles, takeStyleOptions()!));
    }
  });
  if (pendingBreak !== undefined) {
    // A manual break as the LAST block: nothing follows to carry
    // fo:break-before, so an empty carrier paragraph holds it (SPEC-028
    // AC-32 — re-opens as the break plus an empty paragraph; the second
    // cycle is a fixed point).
    blockXml.push(carrierParagraphXml(styles, { breakBefore: pendingBreak }));
  }
  const trackedChanges = trackedChangesDefinitionsXml(odtChangedRegionsContext);
  odtNotesContext = null;
  odtCommentsContext = null;
  odtChangedRegionsContext = null;
  odtFormulaContext = null;
  const body = blockXml.length > 0 ? blockXml.join("") : "<text:p/>";
  // SPEC-053: declare every user field a `docProperty` field references, its
  // value taken from the matching custom property (SPEC-041). LibreOffice
  // needs the declaration before `text:user-field-get` can resolve.
  const userFieldDecls = userFieldDeclsXml(model, styles.userFieldNames());
  const fontFaceDecls = includeOpaque ? (model.sourceXml?.odtFontFaceDecls ?? "") : "";
  const preservedStyles = includeOpaque ? (model.sourceXml?.odtAutomaticStyles ?? []) : [];
  return (
    `${XML_DECLARATION}<office:document-content${DOCUMENT_XMLNS} office:version="1.2">` +
    fontFaceDecls +
    `<office:automatic-styles>${styles.defs.join("")}${preservedStyles.join("")}</office:automatic-styles>` +
    `<office:body><office:text>${userFieldDecls}${trackedChanges}${body}</office:text></office:body>` +
    "</office:document-content>"
  );
}

/** One named paragraph style:style definition (SPEC-019) — the id as
 * style:name (with the display name alongside when it differs), the parent
 * link, the outline level that makes it a heading style, and the style's
 * own paragraph/text properties. */
/** An `unmodeled` style's original XML (SPEC-042 AC-18), re-emitted
 * verbatim except for the user-visible name — mirrors docx/writer.ts's
 * `unmodeledStyleXml`. */
function unmodeledOdfStyleXml(definition: NamedStyle): string {
  if (!definition.rawXml) {
    const displayAttr =
      definition.name !== undefined ? ` style:display-name="${escapeXml(definition.name)}"` : "";
    return `<style:style style:name="${escapeXml(definition.id)}"${displayAttr} style:family="table"/>`;
  }
  try {
    const wrapper = parseXml(
      `<style:root xmlns:style="${STYLE_NAMESPACE}" xmlns:office="${OFFICE_NAMESPACE}">${definition.rawXml}</style:root>`,
    );
    const el = wrapper.documentElement.firstElementChild;
    if (!el) {
      return definition.rawXml;
    }
    if (definition.name !== undefined) {
      el.setAttributeNS(STYLE_NAMESPACE, "style:display-name", definition.name);
    }
    return serializeElement(el);
  } catch {
    return definition.rawXml;
  }
}

function namedStyleFamily(type: NamedStyle["type"]): "paragraph" | "text" | "table" {
  if (type === "character") {
    return "text";
  }
  if (type === "table") {
    return "table";
  }
  return "paragraph";
}

function namedStyleDefinitionXml(definition: NamedStyle): string {
  if (definition.unmodeled) {
    return unmodeledOdfStyleXml(definition);
  }
  const family = namedStyleFamily(definition.type);
  let attrs = `style:name="${escapeXml(definition.id)}"`;
  if (definition.name !== undefined && definition.name !== definition.id) {
    attrs += ` style:display-name="${escapeXml(definition.name)}"`;
  }
  attrs += ` style:family="${family}"`;
  if (definition.basedOn !== undefined) {
    attrs += ` style:parent-style-name="${escapeXml(definition.basedOn)}"`;
  }
  if (definition.next !== undefined && family === "paragraph") {
    attrs += ` style:next-style-name="${escapeXml(definition.next)}"`;
  }
  if (definition.headingLevel !== undefined) {
    attrs += ` style:default-outline-level="${definition.headingLevel}"`;
  }
  if (family === "table") {
    // ODF has no per-table-style border concept (borders are always
    // per-cell) — only background shading round-trips here; see
    // NamedStyle.cellBorders's doc comment for the docx side, which does
    // support it. Table style banding is out of scope either way.
    const hex = definition.cellShadingFill;
    const properties = hex !== undefined ? `<style:table-properties fo:background-color="${escapeXml(hex)}"/>` : "";
    // SPEC-076 AC-68: the style's banding becomes a table:table-template plus
    // the generated table-cell styles its regions reference — siblings of the
    // style itself under office:styles.
    const banding = tableTemplateXml(definition);
    const self =
      properties.length > 0
        ? `<style:style ${attrs}>${properties}</style:style>`
        : `<style:style ${attrs}/>`;
    return self + banding.cellStyles + banding.template;
  }
  const properties =
    (family === "paragraph" ? paragraphPropertiesXml(definition.paragraphFormatting ?? {}) : "") +
    textPropertiesXml(definition.runFormatting ?? {}, definition.fontSizePercent);
  return properties.length > 0
    ? `<style:style ${attrs}>${properties}</style:style>`
    : `<style:style ${attrs}/>`;
}

// --- section master pages (SPEC-028 AC-33) ----------------------------------

/** The master-page name section `index` (0-based) is referenced by: the
 * first section keeps ODF's conventional "Standard"; every later one gets a
 * deterministic generated name, shared between content.xml (the reference)
 * and styles.xml (the definition). */
function sectionMasterName(index: number): string {
  return index === 0 ? "Standard" : `SectionMaster${index + 1}`;
}

function sectionLayoutName(index: number): string {
  return `pm${index + 1}`;
}

interface OdtStyleNamesByName {
  layouts: Map<string, Element>;
  masters: Map<string, Element>;
}

/** Collects every page-layout and master-page the preserved styles part
 * already defines, keyed by style:name — the uniqueness ground a same-format
 * save checks before deciding which generated names it may use. */
function collectOdtStyleNames(root: Element): OdtStyleNamesByName {
  const layouts = new Map<string, Element>();
  const masters = new Map<string, Element>();
  const collect = (parentName: string, localName: string, into: Map<string, Element>) => {
    const parent = firstChildElement(root, OFFICE_NAMESPACE, parentName);
    if (!parent) {
      return;
    }
    for (const el of Array.from(parent.children)) {
      if (el.namespaceURI === STYLE_NAMESPACE && el.localName === localName) {
        const name = el.getAttributeNS(STYLE_NAMESPACE, "name");
        if (name) {
          into.set(name, el);
        }
      }
    }
  };
  collect("automatic-styles", "page-layout", layouts);
  collect("master-styles", "master-page", masters);
  return { layouts, masters };
}

/** The page-layout name the writer uses for the section at `index` (>=1) on
 * a same-format save. The conventional `pm${index+1}` (e.g. pm2) is
 * OpenOffice/older-LibreOffice's own running numbering for automatic
 * page-layouts, so a foreign doc's second section is literally "pm2" —
 * colliding with the factory's generated layout. To honour SPEC-008 (never
 * displace a definition the source owns) the writer scans for a name that
 * either does not exist in the preserved styles at all, or already exists
 * AND is referenced by the writer's OWN master page for that section — the
 * marker that it came from a previous factory save. Anything else is a
 * foreign definition and is skipped, never overwritten. Returns null when
 * every candidate is taken by someone else's styles. */
function usableLayoutName(
  index: number,
  names: OdtStyleNamesByName,
  preferredMasterName: string,
): string | null {
  const preferred = sectionLayoutName(index);
  for (let n = 0; n < 64; n++) {
    const candidate = n === 0 ? preferred : `${preferred}-${n + 1}`;
    if (!names.layouts.has(candidate)) {
      return candidate; // Free — safe to add.
    }
    const master = names.masters.get(preferredMasterName);
    if (master && styleAttr(master, "page-layout-name") === candidate) {
      return candidate; // Ours from a previous factory save — safe to reuse.
    }
  }
  return null; // Every candidate collides with a foreign definition.
}

/** The master-page name the writer uses for the section at `index` (>=1),
 * given the page-layout name that section's generated master page will
 * reference. The conventional `SectionMaster${index+1}` is factory-specific,
 * but a stored definition under that name that references a DIFFERENT layout
 * belongs to the source — never displace it. Returns null when every
 * candidate is someone else's. */
function usableMasterName(
  index: number,
  names: OdtStyleNamesByName,
  layoutName: string,
): string | null {
  const preferred = sectionMasterName(index);
  for (let n = 0; n < 64; n++) {
    const candidate = n === 0 ? preferred : `SectionMaster${index + 1}-${n + 1}`;
    if (!names.masters.has(candidate)) {
      return candidate;
    }
    const master = names.masters.get(candidate);
    if (master && styleAttr(master, "page-layout-name") === layoutName) {
      return candidate; // Ours from a previous factory save.
    }
  }
  return null;
}

/** The per-section master-page/page-layout names a save will emit, aligned
 * with `model.sections` (index 0 keeps the conventional Standard/pm1 pair
 * that every single-section document uses). When a preserved styles.xml is
 * given, names are resolved collision-free against it (see usableLayoutName
 * / usableMasterName); otherwise — fresh documents — the conventional names
 * are used, which buildStylesXml also emits. content.xml's master-page
 * references and the styles patch are built from the SAME plan, so the
 * referenced and defined names always agree. */
function planOdtSectionNames(
  model: DocumentModel,
  originalStylesXml: string | undefined,
): { masters: (string | null)[]; layouts: (string | null)[] } {
  const n = model.sections.length;
  const masters: (string | null)[] = Array.from({ length: n }, (_, i) => sectionMasterName(i));
  const layouts: (string | null)[] = Array.from({ length: n }, (_, i) => sectionLayoutName(i));
  if (n <= 1 || !originalStylesXml) {
    return { masters, layouts };
  }
  let stylesDoc: Document;
  try {
    stylesDoc = parseXml(originalStylesXml);
  } catch {
    return { masters, layouts };
  }
  const names = collectOdtStyleNames(stylesDoc.documentElement);
  for (let index = 1; index < n; index++) {
    const layoutName = usableLayoutName(index, names, sectionMasterName(index));
    if (layoutName === null) {
      // Every candidate is foreign — mark the pair unusable so the patch
      // skips this section rather than clobber.
      masters[index] = null;
      layouts[index] = null;
      continue;
    }
    const masterName = usableMasterName(index, names, layoutName);
    if (masterName === null) {
      masters[index] = null;
      layouts[index] = null;
      continue;
    }
    layouts[index] = layoutName;
    masters[index] = masterName;
    // Register both chosen names so later sections never re-pick them.
    names.layouts.set(layoutName, stylesDoc.documentElement);
    names.masters.set(masterName, stylesDoc.documentElement);
  }
  return { masters, layouts };
}

// The `style:columns` element of a page layout (SPEC-030 AC-17):
// `fo:column-count`/`fo:column-gap`, a `style:column-sep` when "line between"
// is on, and — for unequal columns — a `style:column` per column carrying its
// proportional `style:rel-width` (a relative weight, `*`-suffixed).
function columnsElementXml(columns: SectionColumns, contentWidthCm: number): string {
  const gap = ` fo:column-gap="${cmToOdfLength(Math.max(0, columns.gapCm))}"`;
  const sep = columns.separator
    ? '<style:column-sep style:width="0.002cm" style:color="#000000"/>'
    : "";
  if (columns.equalWidth || !columns.widthsCm || columns.widthsCm.length !== columns.count) {
    return `<style:columns fo:column-count="${columns.count}"${gap}>${sep}</style:columns>`;
  }
  const widths = resolveColumnWidthsCm(columns, contentWidthCm);
  // rel-width is a relative weight; twip-scale keeps integers stable and their
  // ratios exact.
  const cols = widths
    .map((width) => `<style:column style:rel-width="${Math.max(1, Math.round(width * 1000))}*"/>`)
    .join("");
  return `<style:columns fo:column-count="${columns.count}"${gap}>${cols}${sep}</style:columns>`;
}

// --- headers, footers and page numbers (SPEC-029) ---------------------------

/** A minimal AutomaticStyles for header/footer content: mints simply-named,
 * always-unique styles (`${prefix}P1`/`${prefix}T1`, ...) rather than
 * reusing createAutomaticStyles' whole-model dedup index — header/footer
 * content is small enough that dedup buys nothing, and a per-part prefix is
 * what keeps several header/footer parts' styles from colliding once they
 * all land in the same styles.xml automatic-styles block. Tables, images
 * and lists are out of scope for header/footer content (SPEC-029) — a
 * paragraph carrying list formatting (possible only via a .docx source) is
 * rendered as plain text, and those three methods are never called. */
function createHeaderFooterAutomaticStyles(prefix: string): AutomaticStyles {
  const defs: string[] = [];
  const paragraphNames = new Map<string, string>();
  const textNames = new Map<string, string>();
  // SPEC-053: a `date`/`time` field in a header/footer (AC-30, AC-31) needs a
  // number style, registered into this same styles.xml automatic-styles block.
  const dateStyleNames = new Map<string, string>();
  const userFields = new Set<string>();
  return {
    defs,
    dateStyleNameFor(entry: DateTimeFormatEntry): string {
      const existing = dateStyleNames.get(entry.format);
      if (existing !== undefined) {
        return existing;
      }
      const name = odtDateStyleName(entry);
      dateStyleNames.set(entry.format, name);
      const tag = entry.odtStyleTag === "time" ? "number:time-style" : "number:date-style";
      defs.push(`<${tag}${NUMBER_XMLNS} style:name="${escapeXml(name)}">${entry.odtStyleBody}</${tag}>`);
      return name;
    },
    registerUserField(name: string): void {
      userFields.add(name);
    },
    userFieldNames(): string[] {
      return [...userFields];
    },
    paragraphNameFor(paragraph: Paragraph): string | undefined {
      const propertiesXml = paragraphPropertiesXml(paragraph.formatting ?? {});
      if (propertiesXml === "") {
        return undefined;
      }
      const existing = paragraphNames.get(propertiesXml);
      if (existing !== undefined) {
        return existing;
      }
      const name = `${prefix}P${paragraphNames.size + 1}`;
      paragraphNames.set(propertiesXml, name);
      defs.push(`<style:style style:name="${name}" style:family="paragraph">${propertiesXml}</style:style>`);
      return name;
    },
    textNameFor(formatting: RunFormatting): string {
      const propertiesXml = textPropertiesXml(formatting);
      const existing = textNames.get(propertiesXml);
      if (existing !== undefined) {
        return existing;
      }
      const name = `${prefix}T${textNames.size + 1}`;
      textNames.set(propertiesXml, name);
      defs.push(
        propertiesXml === ""
          ? `<style:style style:name="${name}" style:family="text"/>`
          : `<style:style style:name="${name}" style:family="text">${propertiesXml}</style:style>`,
      );
      return name;
    },
    listNameFor(): string {
      throw new Error("lists are not supported in header/footer content (SPEC-029)");
    },
    definedListNameFor(): string {
      throw new Error("lists are not supported in header/footer content (SPEC-029)");
    },
    listDefinitionFor(): ListDefinition | undefined {
      // Header/footer paragraphs never render as defined list items; their
      // writer strips list formatting before it reaches a list run.
      return undefined;
    },
    nextListXmlId(): string {
      throw new Error("lists are not supported in header/footer content (SPEC-029)");
    },
    tableNamesFor(): TableNames {
      throw new Error("tables are not supported in header/footer content (SPEC-029)");
    },
    graphicStyleNameFor(): string | undefined {
      throw new Error("images are not supported in header/footer content (SPEC-029)");
    },
    shapeStyleNameFor(): string {
      throw new Error("shapes are not supported in header/footer content (SPEC-029)");
    },
    connectorStyleNameFor(): string {
      throw new Error("connectors are not supported in header/footer content (SPEC-029)");
    },
  };
}

interface HeaderFooterMasterContent {
  headerXml: string;
  headerLeftXml: string;
  footerXml: string;
  footerLeftXml: string;
  /** style:style defs the content above referenced — added to styles.xml's
   * own office:automatic-styles, a separate scope from content.xml's. */
  defsXml: string;
}

const EMPTY_HEADER_FOOTER_MASTER_CONTENT: HeaderFooterMasterContent = {
  headerXml: "",
  headerLeftXml: "",
  footerXml: "",
  footerLeftXml: "",
  defsXml: "",
};

/** A section's header/footer content as `<style:master-page>` children
 * (SPEC-029 AC-9, AC-24, AC-26): `style:header`/`style:footer` for the
 * default kind, `style:header-left`/`style:footer-left` for "even" —
 * emitted only when the section actually wants different odd/even pages
 * (AC-12), matching the reader's "the -left variant's mere presence turns
 * it on" contract. Paragraphs only (no tables/images — see
 * createHeaderFooterAutomaticStyles); `namePrefix` keeps this section's
 * generated style names from colliding with another section's in the same
 * styles.xml. */
function headerFooterMasterContent(
  section: Section,
  namePrefix: string,
  includeOpaque: boolean,
): HeaderFooterMasterContent {
  const map = { header: section.headers, footer: section.footers };
  if (!map.header && !map.footer) {
    return EMPTY_HEADER_FOOTER_MASTER_CONTENT;
  }
  const styles = createHeaderFooterAutomaticStyles(namePrefix);
  const build = (content: { blocks: Block[] } | undefined, tag: string): string => {
    if (!content) {
      return "";
    }
    const body = content.blocks
      .map((block) => (block.type === "paragraph" ? paragraphXml(block, styles, includeOpaque) : ""))
      .join("");
    return `<style:${tag}>${body.length > 0 ? body : "<text:p/>"}</style:${tag}>`;
  };
  const evenAndOdd = section.evenAndOddHeaders === true;
  return {
    headerXml: build(map.header?.default, "header"),
    headerLeftXml: evenAndOdd ? build(map.header?.even, "header-left") : "",
    footerXml: build(map.footer?.default, "footer"),
    footerLeftXml: evenAndOdd ? build(map.footer?.even, "footer-left") : "",
    defsXml: styles.defs.join(""),
  };
}

/** One style:page-layout carrying a page geometry — the shape the reader's
 * parsePageLayoutProperties parses back. */
/** A section's page border and background on `style:page-layout-properties`
 * (SPEC-074 AC-31): `fo:border-*` per edge with `fo:padding-*` carrying the
 * offset, and `fo:background-color` for the page background — the master
 * page's own properties, which is where ODF puts both. ODF has no attribute
 * for *which* edge the offset measures from, nor for restricting the border
 * to (or from) the first page, so those two ride in the extension namespace
 * like the paragraph `between` edge does. */
function pageLayoutBorderAttrs(pageBorders?: PageBorders, pageBackground?: string): string {
  const attrs: string[] = [];
  if (pageBorders) {
    const offsetPt = pageBorders.offsetPt ?? 24;
    let anyEdge = false;
    for (const edge of ["top", "bottom", "left", "right"] as const) {
      const side = pageBorders[edge];
      if (side) {
        attrs.push(`fo:border-${edge}="${odfBorderShorthand(side)}"`);
        attrs.push(`fo:padding-${edge}="${ptToOdfLength(offsetPt)}"`);
        anyEdge = true;
      }
    }
    if (anyEdge) {
      attrs.push(`xmlns:de="${DE_BORDER_NAMESPACE}"`);
      attrs.push(`de:page-border-measure-from="${pageBorders.measureFrom ?? "page"}"`);
      if (pageBorders.appliesTo !== undefined && pageBorders.appliesTo !== "allPages") {
        attrs.push(`de:page-border-applies-to="${pageBorders.appliesTo}"`);
      }
    }
  }
  if (pageBackground !== undefined) {
    attrs.push(`fo:background-color="${escapeXml(pageBackground)}"`);
  }
  return attrs.length > 0 ? ` ${attrs.join(" ")}` : "";
}

function pageLayoutXml(
  name: string,
  pageSize: PageSize,
  margins: Margins,
  columns?: SectionColumns,
  pageBorders?: PageBorders,
  pageBackground?: string,
): string {
  // ODF stores landscape pages with width/height already swapped; the
  // orientation attribute just marks them, mirroring the .docx writer's
  // w:orient.
  const orientation =
    pageSize.widthCm > pageSize.heightCm ? ' style:print-orientation="landscape"' : "";
  const props =
    `<style:page-layout-properties fo:page-width="${cmToOdfLength(pageSize.widthCm)}"` +
    ` fo:page-height="${cmToOdfLength(pageSize.heightCm)}"${orientation}` +
    ` fo:margin-top="${cmToOdfLength(margins.topCm)}" fo:margin-right="${cmToOdfLength(margins.rightCm)}"` +
    ` fo:margin-bottom="${cmToOdfLength(margins.bottomCm)}" fo:margin-left="${cmToOdfLength(margins.leftCm)}"` +
    pageLayoutBorderAttrs(pageBorders, pageBackground);
  const columnsXml =
    columns && columns.count > 1
      ? columnsElementXml(columns, Math.max(0.1, pageSize.widthCm - margins.leftCm - margins.rightCm))
      : "";
  const propsElement = columnsXml ? `${props}>${columnsXml}</style:page-layout-properties>` : `${props}/>`;
  return `<style:page-layout style:name="${escapeXml(name)}">${propsElement}</style:page-layout>`;
}

function masterPageXml(name: string, layoutName: string, headerFooter?: HeaderFooterMasterContent): string {
  const children =
    (headerFooter?.headerXml ?? "") +
    (headerFooter?.headerLeftXml ?? "") +
    (headerFooter?.footerXml ?? "") +
    (headerFooter?.footerLeftXml ?? "");
  const attrs = `style:name="${escapeXml(name)}" style:page-layout-name="${escapeXml(layoutName)}"`;
  return children.length > 0
    ? `<style:master-page ${attrs}>${children}</style:master-page>`
    : `<style:master-page ${attrs}/>`;
}

/** The styles part: a paragraph default-style carrying the model's default
 * font (the reader's source for defaultStyles.fontFamily/fontSizePt), the
 * model's named paragraph styles (SPEC-019), and one page layout + master
 * page per section (SPEC-007 AC-15; SPEC-028 AC-33) — the first section's
 * pair keeps the "pm1"/"Standard" names every single-section document
 * always used, so those stay byte-identical. */
// style:num-format for a note numbering format (SPEC-031) — ODF has no
// symbol cycle format, so "symbol" degrades to decimal on an .odt save (a
// documented fidelity limit; the in-app rendering keeps the symbols).
const ODT_NOTE_NUM_FORMAT: Record<NoteNumberFormat, string> = {
  decimal: "1",
  lowerRoman: "i",
  lowerAlpha: "a",
  symbol: "1",
};

/** The notes-configuration attributes the model's settings map to (SPEC-031
 * AC-17..20), as [namespaceURI, qualifiedName, value] triples — shared by
 * the from-scratch styles.xml and the preserved-part patch so both write
 * the same thing. `text:start-value` is the counter's initial value, one
 * below the first displayed number (LibreOffice's own convention). */
function notesConfigurationAttrs(settings: NoteSettings): Array<[string, string, string]> {
  const attrs: Array<[string, string, string]> = [];
  if (settings.numberFormat !== undefined) {
    attrs.push([STYLE_NAMESPACE, "style:num-format", ODT_NOTE_NUM_FORMAT[settings.numberFormat]]);
  }
  if (settings.startAt !== undefined) {
    attrs.push([TEXT_NAMESPACE, "text:start-value", String(Math.max(0, settings.startAt - 1))]);
  }
  if (settings.restart === "page") {
    attrs.push([TEXT_NAMESPACE, "text:start-numbering-at", "page"]);
  } else if (settings.restart === "section") {
    attrs.push([TEXT_NAMESPACE, "text:start-numbering-at", "chapter"]);
  }
  return attrs;
}

function notesConfigurationXml(cls: NoteKind, settings: NoteSettings | undefined): string {
  if (!settings) {
    return "";
  }
  const attrs = notesConfigurationAttrs(settings);
  if (attrs.length === 0) {
    return "";
  }
  const rendered = attrs.map(([, qname, value]) => ` ${qname}="${escapeXml(value)}"`).join("");
  return `<text:notes-configuration text:note-class="${cls}"${rendered}/>`;
}

/** A preserved styles.xml with the model's note numbering settings merged
 * onto its `text:notes-configuration` elements (SPEC-031), or null when
 * nothing needs changing. Existing configuration elements keep every
 * attribute this model does not set (citation styles, positions, …) — the
 * patch sets only the modelled attributes, the patchOdtStylesXml
 * philosophy. */
export function patchOdtNotesConfiguration(
  model: DocumentModel,
  originalStylesXml: string,
): string | null {
  const wanted: Array<[NoteKind, NoteSettings]> = [];
  if (model.footnoteSettings && notesConfigurationAttrs(model.footnoteSettings).length > 0) {
    wanted.push(["footnote", model.footnoteSettings]);
  }
  if (model.endnoteSettings && notesConfigurationAttrs(model.endnoteSettings).length > 0) {
    wanted.push(["endnote", model.endnoteSettings]);
  }
  if (wanted.length === 0) {
    return null;
  }
  let stylesDoc: Document;
  try {
    stylesDoc = parseXml(originalStylesXml);
  } catch {
    return null;
  }
  const root = stylesDoc.documentElement;
  let officeStyles = firstChildElement(root, OFFICE_NAMESPACE, "styles");
  if (!officeStyles) {
    officeStyles = stylesDoc.createElementNS(OFFICE_NAMESPACE, "office:styles");
    root.insertBefore(
      officeStyles,
      firstChildElement(root, OFFICE_NAMESPACE, "automatic-styles") ?? null,
    );
  }
  let changed = false;
  for (const [cls, settings] of wanted) {
    const attrs = notesConfigurationAttrs(settings);
    const existing = Array.from(officeStyles.children).find(
      (el) =>
        el.namespaceURI === TEXT_NAMESPACE &&
        el.localName === "notes-configuration" &&
        (el.getAttributeNS(TEXT_NAMESPACE, "note-class") ?? "footnote") === cls,
    );
    if (existing) {
      for (const [ns, qname, value] of attrs) {
        const local = qname.split(":")[1];
        if (existing.getAttributeNS(ns, local) !== value) {
          existing.setAttributeNS(ns, qname, value);
          changed = true;
        }
      }
    } else {
      const wrapper = parseXml(
        `<office:document-styles${DOCUMENT_XMLNS}>${notesConfigurationXml(cls, settings)}</office:document-styles>`,
      );
      officeStyles.appendChild(stylesDoc.importNode(wrapper.documentElement.firstElementChild!, true));
      changed = true;
    }
  }
  return changed ? `${XML_DECLARATION}${serializeElement(root)}` : null;
}

/** The document-default language declared by a styles.xml part's
 * `style:default-style[family=paragraph]`, or undefined when it declares
 * none — the raw value, with no DEFAULT_LANGUAGE fallback applied (that
 * fallback is the reader/patchOdtDefaultLanguage's job, not this parse). */
function defaultLanguageOf(stylesDoc: Document): string | undefined {
  const officeStyles = firstChildElement(stylesDoc.documentElement, OFFICE_NAMESPACE, "styles");
  const defaultStyle = officeStyles
    ? Array.from(officeStyles.children).find(
        (el) =>
          el.namespaceURI === STYLE_NAMESPACE &&
          el.localName === "default-style" &&
          el.getAttributeNS(STYLE_NAMESPACE, "family") === "paragraph",
      )
    : undefined;
  const textProps = defaultStyle ? firstChildElement(defaultStyle, STYLE_NAMESPACE, "text-properties") : undefined;
  const language = textProps?.getAttributeNS(FO_NAMESPACE, "language");
  if (!language) {
    return undefined;
  }
  return composeLanguageTag(language, textProps?.getAttributeNS(FO_NAMESPACE, "country"));
}

/** A preserved styles.xml with the model's default language (SPEC-035)
 * written into `style:default-style`'s `fo:language`/`fo:country`, or null
 * when the model's language already matches what re-opening the *original*
 * part would produce — the same "no patch means byte-identical" contract as
 * patchOdtNotesConfiguration, needed here because
 * DocumentModel.defaultStyles.language is always populated (falls back to
 * DEFAULT_LANGUAGE the same way the reader does), so a naive "is it set"
 * check would rewrite every untouched document's styles.xml on every save. */
export function patchOdtDefaultLanguage(model: DocumentModel, originalStylesXml: string): string | null {
  let stylesDoc: Document;
  try {
    stylesDoc = parseXml(originalStylesXml);
  } catch {
    return null;
  }
  const wanted = model.defaultStyles.language ?? DEFAULT_LANGUAGE;
  const original = defaultLanguageOf(stylesDoc) ?? DEFAULT_LANGUAGE;
  if (wanted === original) {
    return null;
  }
  const root = stylesDoc.documentElement;
  let officeStyles = firstChildElement(root, OFFICE_NAMESPACE, "styles");
  if (!officeStyles) {
    officeStyles = stylesDoc.createElementNS(OFFICE_NAMESPACE, "office:styles");
    root.insertBefore(officeStyles, firstChildElement(root, OFFICE_NAMESPACE, "automatic-styles") ?? null);
  }
  let defaultStyle = Array.from(officeStyles.children).find(
    (el) =>
      el.namespaceURI === STYLE_NAMESPACE &&
      el.localName === "default-style" &&
      el.getAttributeNS(STYLE_NAMESPACE, "family") === "paragraph",
  );
  if (!defaultStyle) {
    const wrapper = parseXml(
      `<office:document-styles${DOCUMENT_XMLNS}><style:default-style style:family="paragraph"/></office:document-styles>`,
    );
    defaultStyle = stylesDoc.importNode(wrapper.documentElement.firstElementChild!, true);
    officeStyles.insertBefore(defaultStyle, officeStyles.firstElementChild ?? null);
  }
  let textProps = firstChildElement(defaultStyle, STYLE_NAMESPACE, "text-properties");
  if (!textProps) {
    textProps = stylesDoc.createElementNS(STYLE_NAMESPACE, "style:text-properties");
    defaultStyle.appendChild(textProps);
  }
  const { language, country } = splitLanguageTag(wanted);
  textProps.setAttributeNS(FO_NAMESPACE, "fo:language", language);
  if (country !== undefined) {
    textProps.setAttributeNS(FO_NAMESPACE, "fo:country", country);
  } else {
    textProps.removeAttributeNS(FO_NAMESPACE, "country");
  }
  return `${XML_DECLARATION}${serializeElement(root)}`;
}

/** A preserved styles.xml with the save's shared drawing definitions
 * (SPEC-073) managed inside `<office:styles>`, or null when nothing needs
 * changing. Every existing definition matching this editor's minting scheme —
 * "de" + an UpperCamelCase tail (deGrad1, deHatch2, deArrowTriangleMed,
 * deDashLongDash) — is a stale definition a previous save of THIS editor
 * wrote: removed and replaced by the fresh set collected while content.xml
 * was built. The uppercase requirement keeps a foreign definition whose name
 * merely begins with lowercase "de" (e.g. "decorative") untouched — deleting
 * it would strand any preserved shape that references it. A preserved
 * definition riding in some block's unsupportedStyleXml re-enters through
 * the defs registry but is skipped here when the part still defines that
 * name (no duplicates). */
const MINTED_DEF_NAME = /^de[A-Z]/;

export function patchOdtShapeDefs(originalStylesXml: string, drawingDefsXml: string): string | null {
  let stylesDoc: Document;
  try {
    stylesDoc = parseXml(originalStylesXml);
  } catch {
    return null;
  }
  const root = stylesDoc.documentElement;
  let officeStyles = firstChildElement(root, OFFICE_NAMESPACE, "styles");
  let changed = false;
  if (officeStyles) {
    for (const child of Array.from(officeStyles.children)) {
      if (
        child.namespaceURI === DRAW_NAMESPACE &&
        ODF_DRAWING_DEF_LOCAL_NAMES.has(child.localName) &&
        MINTED_DEF_NAME.test(child.getAttributeNS(DRAW_NAMESPACE, "name") ?? "")
      ) {
        officeStyles.removeChild(child);
        changed = true;
      }
    }
  }
  if (drawingDefsXml !== "") {
    if (!officeStyles) {
      officeStyles = stylesDoc.createElementNS(OFFICE_NAMESPACE, "office:styles");
      root.insertBefore(
        officeStyles,
        firstChildElement(root, OFFICE_NAMESPACE, "automatic-styles") ?? null,
      );
    }
    // Each definition element declares its own namespaces (see shape.ts's
    // DEF_*_XMLNS), so the plain DOCUMENT_XMLNS wrapper suffices to parse.
    let wrapper: Document;
    try {
      wrapper = parseXml(`<office:document-styles${DOCUMENT_XMLNS}>${drawingDefsXml}</office:document-styles>`);
    } catch {
      return changed ? `${XML_DECLARATION}${serializeElement(root)}` : null;
    }
    for (const def of Array.from(wrapper.documentElement.children)) {
      const name = def.getAttributeNS(DRAW_NAMESPACE, "name");
      const exists =
        name !== null &&
        Array.from(officeStyles.children).some(
          (el) =>
            el.namespaceURI === DRAW_NAMESPACE &&
            el.localName === def.localName &&
            el.getAttributeNS(DRAW_NAMESPACE, "name") === name,
        );
      if (exists) {
        continue;
      }
      officeStyles.appendChild(stylesDoc.importNode(def, true));
      changed = true;
    }
  }
  return changed ? `${XML_DECLARATION}${serializeElement(root)}` : null;
}

export function buildStylesXml(model: DocumentModel, drawingDefsXml = ""): string {
  const { fontFamily, fontSizePt, language, defaultTabStopCm } = model.defaultStyles;
  // The document's default tab interval (SPEC-054): style:tab-stop-distance on
  // the default paragraph style. Omitted when the model carries no explicit
  // value.
  const defaultTabProps =
    defaultTabStopCm !== undefined
      ? `<style:paragraph-properties style:tab-stop-distance="${cmToOdfLength(defaultTabStopCm)}"/>`
      : "";
  const languageAttrs = (() => {
    if (language === undefined) {
      return "";
    }
    const { language: lang, country } = splitLanguageTag(language);
    return ` fo:language="${escapeXml(lang)}"${country !== undefined ? ` fo:country="${escapeXml(country)}"` : ""}`;
  })();
  const includeOpaque = model.format === "odt";
  const namedStyles = (model.styles ?? [])
    .filter((style) => style.name !== undefined)
    .map(namedStyleDefinitionXml)
    .join("");
  const pageLayouts = model.sections
    .map((section, index) =>
      pageLayoutXml(
        sectionLayoutName(index),
        sectionPageSize(model, index),
        sectionMargins(model, index),
        section.columns,
        section.pageBorders,
        section.pageBackground,
      ),
    )
    .join("");
  // Header/footer content (SPEC-029): each section's own automatic styles
  // (namePrefix keyed by index, so no two sections' generated style names
  // collide) alongside its master page.
  const headerFooterContents = model.sections.map((section, index) =>
    headerFooterMasterContent(section, `HF${index + 1}`, includeOpaque),
  );
  const masterPages = model.sections
    .map((_, index) => masterPageXml(sectionMasterName(index), sectionLayoutName(index), headerFooterContents[index]))
    .join("");
  const headerFooterDefs = headerFooterContents.map((content) => content.defsXml).join("");
  return (
    `${XML_DECLARATION}<office:document-styles${DOCUMENT_XMLNS} office:version="1.2">` +
    "<office:styles>" +
    '<style:default-style style:family="paragraph">' +
    defaultTabProps +
    `<style:text-properties fo:font-family="${escapeXml(fontFamily)}" fo:font-size="${ptToOdfLength(fontSizePt)}"${languageAttrs}/>` +
    "</style:default-style>" +
    namedStyles +
    notesConfigurationXml("footnote", model.footnoteSettings) +
    notesConfigurationXml("endnote", model.endnoteSettings) +
    // SPEC-073: the shared drawing definitions (draw:gradient/draw:hatch/
    // draw:fill-image/draw:marker/draw:stroke-dash) the content's graphic
    // styles reference by name — collected while content.xml was built.
    drawingDefsXml +
    "</office:styles>" +
    "<office:automatic-styles>" +
    pageLayouts +
    headerFooterDefs +
    "</office:automatic-styles>" +
    "<office:master-styles>" +
    masterPages +
    "</office:master-styles>" +
    "</office:document-styles>"
  );
}

/** A preserved styles.xml with the model's named-style changes applied
 * (SPEC-019), or null when nothing needs changing — the caller then keeps
 * the original part byte-identical (SPEC-008). A named style is rewritten
 * only when the model's catalogue entry no longer matches what the part
 * defines ("Update <style> to match"); one the part doesn't define at all
 * (a built-in materialized by applying it) is added. `contentXml` is the
 * save's rebuilt content part — the same font-face declarations the reader
 * will resolve style font names against. */
export function patchOdtStylesXml(
  model: DocumentModel,
  originalStylesXml: string,
  contentXml: string,
): string | null {
  if (!model.styles) {
    return null;
  }
  let stylesDoc: Document;
  let contentDoc: Document;
  try {
    stylesDoc = parseXml(originalStylesXml);
    contentDoc = parseXml(contentXml);
  } catch {
    return null;
  }
  const original = new Map(
    catalogToNamedStyles(parseOdfStyleCatalog(contentDoc, stylesDoc)).map((style) => [
      style.id,
      style,
    ]),
  );
  const upserts = (model.styles ?? []).filter((entry) => {
    if (entry.name === undefined) {
      return false;
    }
    const existing = original.get(entry.id);
    return !existing || !equalStructural(existing, entry);
  });
  if (upserts.length === 0) {
    return null;
  }

  const root = stylesDoc.documentElement;
  let officeStyles = firstChildElement(root, OFFICE_NAMESPACE, "styles");
  if (!officeStyles) {
    officeStyles = stylesDoc.createElementNS(OFFICE_NAMESPACE, "office:styles");
    root.insertBefore(
      officeStyles,
      firstChildElement(root, OFFICE_NAMESPACE, "automatic-styles") ?? null,
    );
  }
  for (const definition of upserts) {
    const wrapper = parseXml(
      `<office:document-styles${DOCUMENT_XMLNS}>${namedStyleDefinitionXml(definition)}</office:document-styles>`,
    );
    // A banded table style is several siblings — the style itself, its
    // generated table-cell styles and the table:table-template (SPEC-076
    // AC-68) — so import every element, not just the first.
    const imported = Array.from(wrapper.documentElement.children).map(
      (el) => stylesDoc.importNode(el, true) as Element,
    );
    if (definition.type === "table" && !definition.unmodeled) {
      // Clear the artefacts a previous save generated for this style before
      // re-emitting: a since-removed region's cell style must not linger,
      // and a cleared banding must not resurrect from a stale template.
      const generatedNames = new Set(templateCellStyleNames(definition.id));
      for (const el of Array.from(officeStyles.children)) {
        const staleTemplate =
          el.namespaceURI === TABLE_NAMESPACE &&
          el.localName === "table-template" &&
          el.getAttributeNS(TABLE_NAMESPACE, "name") === definition.id;
        const staleCellStyle =
          el.namespaceURI === STYLE_NAMESPACE &&
          el.localName === "style" &&
          el.getAttributeNS(STYLE_NAMESPACE, "family") === "table-cell" &&
          generatedNames.has(el.getAttributeNS(STYLE_NAMESPACE, "name") ?? "");
        if (staleTemplate || staleCellStyle) {
          officeStyles.removeChild(el);
        }
      }
    }
    // Matched by name alone (SPEC-042): a style's family no longer maps 1:1
    // onto "paragraph", and a style:name is unique across families in one
    // styles.xml the same way a .docx styleId is.
    const existing = Array.from(officeStyles.children).find(
      (el) =>
        el.namespaceURI === STYLE_NAMESPACE &&
        el.localName === "style" &&
        el.getAttributeNS(STYLE_NAMESPACE, "name") === definition.id,
    );
    const [self, ...siblings] = imported;
    if (existing) {
      officeStyles.replaceChild(self, existing);
    } else {
      officeStyles.appendChild(self);
    }
    let anchor = self;
    for (const sibling of siblings) {
      officeStyles.insertBefore(sibling, anchor.nextSibling);
      anchor = sibling;
    }
  }
  return `${XML_DECLARATION}${serializeElement(root)}`;
}

/** A preserved styles.xml with the per-section master pages and page
 * layouts content.xml references upserted (SPEC-028 AC-33), or null when it
 * already defines them all exactly — a same-format save of an unedited
 * multi-section document then keeps the part byte-identical (the
 * patchOdtStylesXml contract). The source's own master pages are left
 * untouched; only the resolved generated pairs from `sectionNames` (the
 * plan planOdtSectionNames computed against the preserved part) are managed,
 * so a foreign definition that shares a conventional "pmN" name is never
 * displaced (SPEC-008) while a definition a previous factory save wrote is
 * replaced in place when its geometry was edited. */
export function patchOdtSectionMasterPages(
  model: DocumentModel,
  originalStylesXml: string,
  sectionNames?: { masters: (string | null)[]; layouts: (string | null)[] },
): string | null {
  if (model.sections.length <= 1) {
    return null;
  }
  let stylesDoc: Document;
  try {
    stylesDoc = parseXml(originalStylesXml);
  } catch {
    return null;
  }
  const root = stylesDoc.documentElement;
  const names = sectionNames ?? planOdtSectionNames(model, originalStylesXml);
  let changed = false;
  const upsert = (parentName: string, localName: string, styleName: string, xml: string) => {
    let parent = firstChildElement(root, OFFICE_NAMESPACE, parentName);
    if (!parent) {
      // Degenerate packages only — any real producer writes both
      // containers. office:master-styles must stay last per the schema.
      parent = stylesDoc.createElementNS(OFFICE_NAMESPACE, `office:${parentName}`);
      const masterStyles =
        parentName === "master-styles" ? null : firstChildElement(root, OFFICE_NAMESPACE, "master-styles");
      root.insertBefore(parent, masterStyles ?? null);
    }
    const wrapper = parseXml(`<office:document-styles${DOCUMENT_XMLNS}>${xml}</office:document-styles>`);
    const imported = stylesDoc.importNode(wrapper.documentElement.firstElementChild!, true);
    const existing = Array.from(parent.children).find(
      (el) =>
        el.namespaceURI === STYLE_NAMESPACE &&
        el.localName === localName &&
        el.getAttributeNS(STYLE_NAMESPACE, "name") === styleName,
    );
    if (existing) {
      if (serializeElement(existing) === serializeElement(imported)) {
        return; // Already up to date — nothing to change.
      }
      parent.replaceChild(imported, existing);
    } else {
      parent.appendChild(imported);
    }
    changed = true;
  };
  for (let index = 1; index < model.sections.length; index++) {
    const masterName = names.masters[index];
    const layoutName = names.layouts[index];
    if (layoutName === null || masterName === null) {
      // Every generated name collided with a foreign definition — the master
      // this save would define belongs to someone else's styles. Refuse to
      // clobber; the reader will fall back to the document default layout.
      continue;
    }
    upsert(
      "automatic-styles",
      "page-layout",
      layoutName,
      pageLayoutXml(
        layoutName,
        sectionPageSize(model, index),
        sectionMargins(model, index),
        model.sections[index]?.columns,
        model.sections[index]?.pageBorders,
        model.sections[index]?.pageBackground,
      ),
    );
    upsert(
      "master-styles",
      "master-page",
      masterName,
      masterPageXml(masterName, layoutName),
    );
  }
  return changed ? `${XML_DECLARATION}${serializeElement(root)}` : null;
}

/** A preserved styles.xml with every section's header/footer content
 * upserted into its master page (SPEC-029), or null when nothing changed.
 * Unlike patchOdtSectionMasterPages (page geometry, sections 1+ only —
 * section 0 keeps whatever "Standard" the source already defines), this
 * runs for EVERY section including 0: header/footer content is new,
 * user-editable content with nowhere else to live, so section 0's existing
 * "Standard" master page is itself augmented in place rather than left
 * alone. A section whose generated master name collided with a foreign
 * definition (per `sectionNames`) has no master-page element resolved by
 * the plan; a section with content that has no such element to upsert
 * INTO at all (e.g. "Standard" is missing from a degenerate package) is
 * silently skipped rather than fabricating a master page — this function
 * only ever augments existing master pages. */
export function patchOdtSectionHeaderFooters(
  model: DocumentModel,
  originalStylesXml: string,
  sectionNames?: { masters: (string | null)[]; layouts: (string | null)[] },
): string | null {
  const includeOpaque = model.format === "odt";
  let stylesDoc: Document;
  try {
    stylesDoc = parseXml(originalStylesXml);
  } catch {
    return null;
  }
  const root = stylesDoc.documentElement;
  const masterStyles = firstChildElement(root, OFFICE_NAMESPACE, "master-styles");
  if (!masterStyles) {
    return null;
  }
  const names = model.sections.length > 1 ? (sectionNames ?? planOdtSectionNames(model, originalStylesXml)) : null;
  let changed = false;

  const findMasterPage = (name: string): Element | undefined =>
    Array.from(masterStyles.children).find(
      (el) =>
        el.namespaceURI === STYLE_NAMESPACE &&
        el.localName === "master-page" &&
        el.getAttributeNS(STYLE_NAMESPACE, "name") === name,
    );

  const upsertDefs = (prefix: string, defsXml: string) => {
    let automaticStyles = firstChildElement(root, OFFICE_NAMESPACE, "automatic-styles");
    if (!automaticStyles && defsXml === "") {
      return;
    }
    if (!automaticStyles) {
      automaticStyles = stylesDoc.createElementNS(OFFICE_NAMESPACE, "office:automatic-styles");
      root.insertBefore(automaticStyles, masterStyles);
    }
    // Deterministic prefixed names (see buildStylesXml) — drop every style
    // this section minted on a PREVIOUS save before adding the current set,
    // so removed/changed header-footer content never leaves orphaned defs.
    for (const el of Array.from(automaticStyles.children)) {
      if (
        el.namespaceURI === STYLE_NAMESPACE &&
        el.localName === "style" &&
        el.getAttributeNS(STYLE_NAMESPACE, "name")?.startsWith(prefix)
      ) {
        automaticStyles.removeChild(el);
        changed = true;
      }
    }
    if (defsXml === "") {
      return;
    }
    const wrapper = parseXml(`<office:document-styles${DOCUMENT_XMLNS}><office:automatic-styles>${defsXml}</office:automatic-styles></office:document-styles>`);
    for (const def of Array.from(wrapper.documentElement.firstElementChild!.children)) {
      automaticStyles.appendChild(stylesDoc.importNode(def, true));
    }
    changed = true;
  };

  const upsertHeaderFooterChild = (masterPage: Element, localName: string, xml: string) => {
    const existing = Array.from(masterPage.children).find(
      (el) => el.namespaceURI === STYLE_NAMESPACE && el.localName === localName,
    );
    if (xml === "") {
      if (existing) {
        masterPage.removeChild(existing);
        changed = true;
      }
      return;
    }
    const wrapper = parseXml(`<office:document-styles${DOCUMENT_XMLNS}>${xml}</office:document-styles>`);
    const imported = stylesDoc.importNode(wrapper.documentElement.firstElementChild!, true);
    if (existing) {
      if (serializeElement(existing) === serializeElement(imported)) {
        return;
      }
      masterPage.replaceChild(imported, existing);
    } else {
      masterPage.appendChild(imported);
    }
    changed = true;
  };

  model.sections.forEach((section, index) => {
    const masterName = index === 0 ? "Standard" : (names?.masters[index] ?? null);
    if (masterName === null) {
      return;
    }
    const masterPage = findMasterPage(masterName);
    if (!masterPage) {
      return;
    }
    const prefix = `HF${index + 1}`;
    const content = headerFooterMasterContent(section, prefix, includeOpaque);
    upsertHeaderFooterChild(masterPage, "header", content.headerXml);
    upsertHeaderFooterChild(masterPage, "header-left", content.headerLeftXml);
    upsertHeaderFooterChild(masterPage, "footer", content.footerXml);
    upsertHeaderFooterChild(masterPage, "footer-left", content.footerLeftXml);
    upsertDefs(prefix, content.defsXml);
  });

  return changed ? `${XML_DECLARATION}${serializeElement(root)}` : null;
}

/** Resolves the model's `properties` into `buildMetaXml`'s input — the
 * `.odt` mirror of docx/writer.ts's `coreInputFrom`. */
function metaInputFrom(properties: DocumentProperties | undefined, fallbackTitle: string): OdfMetaInput {
  const title = properties?.title?.trim() ? properties.title : fallbackTitle;
  return {
    title,
    subject: properties?.subject,
    author: properties?.author,
    lastModifiedBy: properties?.lastModifiedBy,
    keywords: properties?.keywords,
    comments: properties?.comments,
    created: properties?.created,
    extraFields: properties?.extraFields,
    customProperties: properties?.customProperties,
    unsupportedCustomProperties: properties?.unsupportedCustomProperties,
  };
}

/** settings.xml (AC-7). The app has no settings to persist; office:settings
 * is optional in the schema, and omitting it entirely is valid while an
 * empty office:settings element is not. */
export function buildSettingsXml(): string {
  return `${XML_DECLARATION}<office:document-settings xmlns:office="${OFFICE_NAMESPACE}" office:version="1.2"/>`;
}

/** META-INF/manifest.xml listing every part (AC-8), plus a file-entry per
 * image media part the model references (SPEC-024 AC-26). The mimetype
 * entry is not listed — per ODF its media type belongs to the "/" entry. */
export function buildManifestXml(mediaEntries: OdtManifestEntry[] = []): string {
  const part = (path: string, mediaType: string) =>
    `<manifest:file-entry manifest:full-path="${path}" manifest:media-type="${mediaType}"/>`;
  return (
    `${XML_DECLARATION}<manifest:manifest xmlns:manifest="${MANIFEST_NAMESPACE}" manifest:version="1.2">` +
    `<manifest:file-entry manifest:full-path="/" manifest:version="1.2" manifest:media-type="${ODT_MIMETYPE}"/>` +
    part("content.xml", "text/xml") +
    part("styles.xml", "text/xml") +
    part("meta.xml", "text/xml") +
    part("settings.xml", "text/xml") +
    mediaEntries.map((entry) => part(entry.path, entry.mediaType)).join("") +
    "</manifest:manifest>"
  );
}

export interface WriteOdtOptions {
  /** The dc:date modification timestamp; defaults to now. Injectable so
   * tests can assert an exact meta.xml. */
  modified?: Date;
}

/** Ensures a preserved manifest lists the parts this save (re)builds —
 * needed when the source package lacked meta.xml or settings.xml, or gains
 * image media parts (SPEC-024). The patch is textual and no-ops gracefully
 * on producers that bound the manifest namespace to another prefix (their
 * packages carry the parts already). */
function withManifestEntries(manifestXml: string, entries: OdtManifestEntry[]): string {
  let out = manifestXml;
  for (const { path, mediaType } of entries) {
    if (!out.includes(`manifest:full-path="${path}"`)) {
      out = out.replace(
        /<\/manifest:manifest>\s*$/,
        `<manifest:file-entry manifest:full-path="${path}" manifest:media-type="${mediaType}"/>` +
          "</manifest:manifest>",
      );
    }
  }
  return out;
}

/** The inverse of withManifestEntries: strips every file-entry whose
 * `manifest:full-path` is `dir/` or `dir/content.xml` — a deleted chart's
 * embedded-object entries (SPEC-052 AC "removes its part, its relationship"
 * — the manifest entry is ODF's relationship). Textual, matching
 * withManifestEntries' own approach. */
function withoutManifestObjectDirs(manifestXml: string, dirs: Set<string>): string {
  if (dirs.size === 0) {
    return manifestXml;
  }
  return manifestXml.replace(/<manifest:file-entry\b[^>]*\/>/g, (entry) => {
    const match = /manifest:full-path="([^"]+)"/.exec(entry);
    const path = match?.[1];
    if (!path) {
      return entry;
    }
    for (const dir of dirs) {
      if (path === `${dir}/` || path === `${dir}/content.xml`) {
        return "";
      }
    }
    return entry;
  });
}

/** Serialises the model into `.odt` bytes plus the list of items that could
 * not be represented (SPEC-008).
 *
 * A model opened from a .odt keeps its retained source package: every part
 * the reader did not model (styles.xml, images, embedded objects, …) is
 * written back byte-identical, and only content.xml and meta.xml
 * (title/timestamp) are rebuilt — `dropped` is empty. Anything else (blank
 * documents, models opened from .docx) is written from scratch as a valid
 * ODF package whose uncompressed mimetype entry comes first (SPEC-007 AC-3,
 * AC-9, AC-10), followed by content.xml, styles.xml, meta.xml, settings.xml
 * and META-INF/manifest.xml (AC-4..AC-8), with unrepresentable .docx
 * content reported in `dropped`. */
export async function writeOdt(
  sourceModel: DocumentModel,
  options: WriteOdtOptions = {},
): Promise<WriteResult> {
  // "Remove personal information on save" (SPEC-041 AC-12..14) — see
  // docx/writer.ts's writeDocx for why this is a single model-level
  // transform shared by both writers rather than format-specific logic.
  const model = sourceModel.properties?.removePersonalInfoOnSave ? stripPersonalInfo(sourceModel) : sourceModel;
  const modified = options.modified ?? new Date();
  const title = model.title.trim().length > 0 ? model.title : UNTITLED_ODT_TITLE;
  const metaInput = metaInputFrom(model.properties, title);
  const encoder = new TextEncoder();

  const preservedParts = model.format === "odt" ? opaquePartBytes(model) : new Map<string, Uint8Array>();
  // Images (SPEC-024): media names start past whatever the preserved
  // package already uses, so a freshly written Pictures/imageN.ext can
  // never collide (mirrors docx/writer.ts's own media planning).
  const mediaPlan = planOdtMediaWrite(model, new Set(preservedParts.keys()));
  const decoder = new TextDecoder("utf-8");
  const originalStyles = preservedParts.get("styles.xml")
    ? decoder.decode(preservedParts.get("styles.xml")!)
    : undefined;
  // Resolve the per-section master-page/page-layout names once, against the
  // preserved styles.xml (SPEC-028 AC-33). content.xml (the reference) and
  // the styles patch (the definition) must agree; when no styles.xml is
  // preserved the conventional SectionMasterN/pmN names are used, matching
  // buildStylesXml.
  const sectionNames = planOdtSectionNames(model, originalStyles);
  // Charts (SPEC-052): an untouched chart reuses its own preserved embedded
  // object directory (so the verbatim pass-through below keeps re-emitting
  // it byte-identical), an edited or freshly inserted one is regenerated
  // from the model; a fresh/cross-format save (no preserved package)
  // regenerates every catalogue-type chart and drops an unsupported one
  // entirely. `preservedChartObjectDirs` — which "Object N" directories the
  // preserved manifest already declares as charts — is what tells a deleted
  // chart's object apart from an unrelated embedded OLE object sharing the
  // same naming convention.
  const preservedManifestBytes = preservedParts.get("META-INF/manifest.xml");
  const preservedChartObjectDirs = preservedManifestBytes
    ? chartObjectDirsInManifest(decoder.decode(preservedManifestBytes))
    : new Set<string>();
  const chartPlan = planOdfChartWrite(model, preservedChartObjectDirs, preservedParts.size === 0);
  // Equations (SPEC-056): every equation regenerates its embedded formula
  // sub-document fresh (the MathNode body is the source of truth), and any
  // preserved formula object dir is orphaned — the formula counterpart of the
  // chart plan, but with no verbatim-reuse case since equations are fully
  // modelled.
  const preservedFormulaObjectDirs = preservedManifestBytes
    ? formulaObjectDirsInManifest(decoder.decode(preservedManifestBytes))
    : new Set<string>();
  const formulaPlan = planOdfFormulaWrite(model, preservedFormulaObjectDirs);
  // SPEC-073: the shared drawing definitions (gradients, hatches, fill
  // images, markers, dash patterns) shape/connector graphic styles reference
  // — the registry fills while content.xml is built, then the definitions
  // land in styles.xml `<office:styles>` (fresh build or preserved patch).
  const drawingDefs = createOdfDrawingDefs();
  const content = buildContentXml(model, mediaPlan, sectionNames.masters, chartPlan, formulaPlan, drawingDefs);

  if (preservedParts.size === 0) {
    const bytes = await zip([
      // ODF requires the mimetype entry first and STOREd so format sniffers
      // can read it from the raw bytes at a fixed offset.
      { name: "mimetype", data: encoder.encode(ODT_MIMETYPE), store: true },
      { name: "content.xml", data: encoder.encode(content) },
      { name: "styles.xml", data: encoder.encode(buildStylesXml(model, drawingDefs.defsXml())) },
      { name: "meta.xml", data: encoder.encode(buildMetaXml(metaInput, modified)) },
      { name: "settings.xml", data: encoder.encode(buildSettingsXml()) },
      {
        name: "META-INF/manifest.xml",
        data: encoder.encode(
          buildManifestXml([
            ...mediaPlan.manifestEntries,
            ...chartPlan.manifestEntries,
            ...formulaPlan.manifestEntries,
          ]),
        ),
      },
      ...mediaPlan.zipEntries,
      ...chartPlan.zipEntries,
      ...formulaPlan.zipEntries,
    ]);
    return { bytes, dropped: collectDropRecords(model, "odt") };
  }

  // Same-format save of an opened .odt: preserved parts verbatim, rebuilt
  // content.xml/meta.xml, settings.xml generated only when the source
  // package had none (the manifest is patched to match), and styles.xml
  // patched only when the model's style catalogue diverges from it
  // (SPEC-019).
  const addSettings = !preservedParts.has("settings.xml");
  const addedPaths: OdtManifestEntry[] = [
    { path: "content.xml", mediaType: "text/xml" },
    { path: "meta.xml", mediaType: "text/xml" },
    ...(addSettings ? [{ path: "settings.xml", mediaType: "text/xml" }] : []),
    ...mediaPlan.manifestEntries,
    ...chartPlan.manifestEntries,
    ...formulaPlan.manifestEntries,
  ];
  // Named-style changes (SPEC-019) first, then the per-section master pages
  // content.xml references (SPEC-028 AC-33), then header/footer content
  // (SPEC-029) — layered so an unedited save still leaves the part
  // byte-identical (every patch returns null then).
  const preservedStyles = preservedParts.get("styles.xml");
  let patchedStyles: string | null = null;
  if (preservedStyles) {
    const originalStyles = decoder.decode(preservedStyles);
    patchedStyles = patchOdtStylesXml(model, originalStyles, content);
    patchedStyles =
      patchOdtSectionMasterPages(model, patchedStyles ?? originalStyles, sectionNames) ?? patchedStyles;
    patchedStyles =
      patchOdtSectionHeaderFooters(model, patchedStyles ?? originalStyles, sectionNames) ?? patchedStyles;
    patchedStyles =
      patchOdtNotesConfiguration(model, patchedStyles ?? originalStyles) ?? patchedStyles;
    patchedStyles =
      patchOdtDefaultLanguage(model, patchedStyles ?? originalStyles) ?? patchedStyles;
    // Shared drawing definitions (SPEC-073) last: stale de-named defs from a
    // previous save are replaced by the set this content build collected.
    patchedStyles =
      patchOdtShapeDefs(patchedStyles ?? originalStyles, drawingDefs.defsXml()) ?? patchedStyles;
  }

  const entries: ZipEntryInput[] = [
    { name: "mimetype", data: preservedParts.get("mimetype") ?? encoder.encode(ODT_MIMETYPE), store: true },
    { name: "content.xml", data: encoder.encode(content) },
  ];
  for (const [name, data] of preservedParts) {
    if (name === "mimetype" || name === "content.xml" || name === "meta.xml") {
      continue; // Already emitted or rebuilt below.
    }
    if (chartPlan.orphanedPartNames.has(name) || formulaPlan.orphanedPartNames.has(name)) {
      continue; // A deleted chart / regenerated formula's embedded object — not resurrected (SPEC-008).
    }
    if (name === "META-INF/manifest.xml") {
      let manifestXml = withManifestEntries(decoder.decode(data), addedPaths);
      manifestXml = withoutManifestObjectDirs(manifestXml, chartPlan.orphanedObjectDirs);
      manifestXml = withoutManifestObjectDirs(manifestXml, formulaPlan.orphanedObjectDirs);
      entries.push({ name, data: encoder.encode(manifestXml) });
    } else if (patchedStyles !== null && name === "styles.xml") {
      entries.push({ name, data: encoder.encode(patchedStyles) });
    } else if (chartPlan.regeneratedPartXml.has(name)) {
      entries.push({ name, data: encoder.encode(chartPlan.regeneratedPartXml.get(name)!) });
    } else {
      entries.push({ name, data });
    }
  }
  entries.push({ name: "meta.xml", data: encoder.encode(buildMetaXml(metaInput, modified)) });
  entries.push(...mediaPlan.zipEntries);
  entries.push(...chartPlan.zipEntries);
  entries.push(...formulaPlan.zipEntries);
  if (addSettings) {
    entries.push({ name: "settings.xml", data: encoder.encode(buildSettingsXml()) });
  }

  // SPEC-008's rule that a same-format save reports nothing still holds for
  // every part except an edited chart's (SPEC-052 "Rewriting a chart" — the
  // one narrow, documented exception) and a bar tab stop, which OpenDocument
  // Text cannot express: it is stripped from the rebuilt content.xml above
  // (paragraphPropertiesXml filters alignment==='bar'), so every one must be
  // reported here too, not only on the fresh/cross-format path (SPEC-054
  // AC-38 — a document opened from .odt takes this same-format branch).
  return {
    bytes: await zip(entries),
    dropped: [
      ...collectEditedChartDropRecords(model, "odt"),
      ...collectBarTabDropRecords(model, "odt"),
      // SPEC-074 AC-32: ODF has no dash-dot border keyword, on any save.
      ...collectBorderStyleDropRecords(model),
    ],
  };
}
