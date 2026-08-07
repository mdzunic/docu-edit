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
  BorderSide,
  BreakKind,
  DocumentModel,
  FieldSpec,
  Image,
  ImageWrapMode,
  Margins,
  NamedStyle,
  PageNumberFormat,
  PageSize,
  Paragraph,
  ParagraphFormatting,
  Run,
  RunFormatting,
  Section,
  Table,
  TableCell,
  TableRow,
} from "../documentModel";
import { FIELD_PLACEHOLDER_TEXT, sectionMargins, sectionPageSize } from "../documentModel";
import { collectDropRecords, opaquePartBytes, type WriteResult } from "../opaque";
import { equalStructural } from "../equality";
import {
  BULLET_CHARS,
  clampListLevel,
  LIST_MARKER_WIDTH_CM,
  planListLevels,
  type PlannedListLevel,
} from "../lists";
import { zip, type ZipEntryInput } from "../docx/zipWriter";
import { firstChildElement, parseXml, serializeElement } from "../docx/xml";
import { ODT_MIMETYPE } from "./reader";
import { drawFrameXml, graphicStyleXml } from "./drawing";
import { planOdtMediaWrite, type OdtManifestEntry, type OdtMediaWritePlan } from "./media";
import { HIGHLIGHT_HEX } from "./properties";
import { catalogToNamedStyles, parseOdfStyleCatalog } from "./styles";
import { FO_NAMESPACE, OFFICE_NAMESPACE, STYLE_NAMESPACE, TABLE_NAMESPACE, TEXT_NAMESPACE, XLINK_NAMESPACE, styleAttr } from "./xml";
import { inlineWriteLayout, linkTargetOf, type LinkTarget } from "../links";
import { cmToOdfLength, multiplierToOdfPercentage, ptToOdfLength } from "./units";

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

/** Escapes text for use in XML text content or attribute values. */
function escapeXml(text: string): string {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
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
  return attrs.length > 0 ? `<style:text-properties ${attrs.join(" ")}/>` : "";
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
  // "Start numbering at" (SPEC-029 AC-22) — the section-starting
  // paragraph's own restart value; absent means "continue from the
  // previous section", which needs no attribute at all.
  if (pageNumberStart !== undefined) {
    attrs.push(`style:page-number="${pageNumberStart}"`);
  }
  return attrs.length > 0 ? `<style:paragraph-properties ${attrs.join(" ")}/>` : "";
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

/** Assigns and serialises the automatic styles of content.xml. Paragraph
 * styles are named after the paragraph's styleId (all styleIds are reserved
 * up front, so a generated name can never shadow one); styleless formatted
 * paragraphs get "P<n>" and runs share "T<n>" styles per distinct
 * formatting; list runs share "L<n>" list styles per distinct level set
 * (SPEC-018). */
/** A `BorderSide` as ODF's `fo:border`-family shorthand ("0.5pt solid
 * #ff0000", or "none"); SPEC-023's border style enum maps onto ODF's CSS2-
 * derived keywords one for one except "single", which ODF calls "solid". */
function borderShorthand(side: BorderSide): string {
  if (side.style === "none") {
    return "none";
  }
  const styleWord = side.style === "single" ? "solid" : side.style;
  const color = side.color !== "auto" ? side.color : "#000000";
  return `${ptToOdfLength(Math.max(0.25, side.widthPt))} ${styleWord} ${color}`;
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
  textNameFor(formatting: RunFormatting): string;
  listNameFor(levels: PlannedListLevel[]): string;
  tableNamesFor(table: Table, preserveNames: boolean): TableNames;
  /** A floating image's graphic-family style (SPEC-024 AC-13, AC-14) —
   * shared by every image with the same wrap mode; undefined for "inline",
   * which needs no positioning style of its own (see drawFrameXml). */
  graphicStyleNameFor(wrap: ImageWrapMode): string | undefined;
}

function createAutomaticStyles(model: DocumentModel): AutomaticStyles {
  const takenNames = new Set<string>();
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
      if (block.type === "image" || block.type === "break") {
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
  const graphicNames = new Map<ImageWrapMode, string>();
  const boundStyleIds = new Set<string>();

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

    textNameFor(formatting: RunFormatting): string {
      const propertiesXml = textPropertiesXml(formatting);
      const existing = textNames.get(propertiesXml);
      if (existing !== undefined) {
        return existing;
      }
      const name = `T${textNames.size + 1}`;
      textNames.set(propertiesXml, name);
      defs.push(
        propertiesXml === ""
          ? `<style:style style:name="${name}" style:family="text"/>`
          : `<style:style style:name="${name}" style:family="text">${propertiesXml}</style:style>`,
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
function fieldContentXml(field: FieldSpec): string {
  const tag = field.kind === "page" ? "page-number" : "page-count";
  return `<text:${tag}${FIELD_NUM_FORMAT[field.format]}>${FIELD_PLACEHOLDER_TEXT}</text:${tag}>`;
}

function runXml(run: Run, styles: AutomaticStyles, state: WhitespaceState): string {
  const styleName = styles.textNameFor(run.formatting);
  if (run.field) {
    state.lastChar = FIELD_PLACEHOLDER_TEXT;
    return `<text:span text:style-name="${styleName}">${fieldContentXml(run.field)}</text:span>`;
  }
  return `<text:span text:style-name="${styleName}">${runContentXml(run.text, state)}</text:span>`;
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
  if (opaque.length === 0 && !hasLinks && !hasBookmarks) {
    return paragraph.runs.map((run) => runXml(run, styles, state)).join("");
  }
  const { runs, bookmarksAtBoundary, boundaryFor } = inlineWriteLayout(paragraph);
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

/** A paragraph as text:p — or text:h with text:outline-level when the model
 * carries a heading level, so the level survives re-opening regardless of
 * the style's definition. */
function paragraphXml(
  paragraph: Paragraph,
  styles: AutomaticStyles,
  includeOpaque: boolean,
  styleOptions?: ParagraphStyleOptions,
): string {
  const styleName = styles.paragraphNameFor(paragraph, styleOptions);
  const tag = paragraph.headingLevel !== undefined ? "text:h" : "text:p";
  let attrs = styleName !== undefined ? ` text:style-name="${escapeXml(styleName)}"` : "";
  if (paragraph.headingLevel !== undefined) {
    attrs += ` text:outline-level="${paragraph.headingLevel}"`;
  }
  if (paragraph.runs.some((run) => linkTargetOf(run.formatting) !== null)) {
    attrs += LINK_XMLNS;
  }
  const content = inlineContentXml(paragraph, styles, includeOpaque);
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

/** A cell's block content (SPEC-022): paragraphs, with consecutive list
 * items grouped into text:list runs exactly like body content — a cell is
 * its own list context, so runs never join one outside the table. */
function cellContentXml(blocks: Paragraph[], styles: AutomaticStyles, includeOpaque: boolean): string {
  let xml = "";
  let i = 0;
  while (i < blocks.length) {
    if (!blocks[i].list) {
      xml += paragraphXml(blocks[i], styles, includeOpaque);
      i++;
      continue;
    }
    let end = i + 1;
    while (end < blocks.length && blocks[end].list) {
      end++;
    }
    xml += listRunXml(blocks.slice(i, end), styles, includeOpaque);
    i = end;
  }
  return xml.length > 0 ? xml : "<text:p/>";
}

/** A grid slot's owning real cell — the writer's own mirror of the same
 * concept in lib/tables.ts and docx/writer.ts (SPEC-023): every slot a
 * merged cell covers becomes `<table:covered-table-cell/>`, ODF's explicit
 * per-slot placeholder (no gridSpan-style inference needed, unlike .docx). */
interface CellOwner {
  originRow: number;
  originCol: number;
  rowSpan: number;
  colSpan: number;
}

function computeOwnerGrid(table: Table): CellOwner[][] {
  const rowCount = table.rows.length;
  const colCount = table.columnWidthsCm.length;
  const grid: CellOwner[][] = Array.from({ length: rowCount }, () => Array.from<CellOwner>({ length: colCount }));
  for (let r = 0; r < rowCount; r++) {
    for (let c = 0; c < colCount; c++) {
      if (grid[r][c]) {
        continue;
      }
      const cell = table.rows[r].cells[c];
      const owner: CellOwner = {
        originRow: r,
        originCol: c,
        rowSpan: cell.rowSpan ?? 1,
        colSpan: cell.colSpan ?? 1,
      };
      for (let dr = 0; dr < owner.rowSpan && r + dr < rowCount; dr++) {
        for (let dc = 0; dc < owner.colSpan && c + dc < colCount; dc++) {
          grid[r + dr][c + dc] = owner;
        }
      }
    }
  }
  return grid;
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
  let xml =
    `<table:table${TABLE_XMLNS} table:name="${escapeXml(names.name)}"` +
    ` table:style-name="${escapeXml(names.styleName)}"${alignAttr}>`;
  for (const columnName of names.columnStyleNames) {
    xml += `<table:table-column table:style-name="${escapeXml(columnName)}"/>`;
  }
  const ownerGrid = computeOwnerGrid(table);
  table.rows.forEach((row, rowIndex) => {
    const rowStyleName =
      includeOpaque && row.odtStyleName !== undefined ? row.odtStyleName : names.rowStyleName(row);
    const rowStyleAttr = rowStyleName !== undefined ? ` table:style-name="${escapeXml(rowStyleName)}"` : "";
    xml += `<table:table-row${rowStyleAttr}>`;
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
      xml +=
        `<table:table-cell${cellStyleAttr}${spanAttrs} office:value-type="string">` +
        cellContentXml(cell.blocks, styles, includeOpaque) +
        "</table:table-cell>";
    }
    xml += "</table:table-row>";
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
): string {
  const includeOpaque = model.format === "odt";
  const styles = createAutomaticStyles(model);
  const blockXml: string[] = [];
  // SPEC-028 AC-32: ODF has no standalone break element — a BreakBlock
  // becomes fo:break-before on the NEXT emitted paragraph's automatic
  // style, so the break is held pending until that paragraph is reached.
  let pendingBreak: BreakKind | undefined;
  model.sections.forEach((section, sectionIndex) => {
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
      if (block.type === "break") {
        // SPEC-028 AC-32: held until the next paragraph (see pendingBreak);
        // the break block itself emits nothing here.
        pendingBreak = block.kind;
        i++;
        continue;
      }
      if (!block.list) {
        blockXml.push(paragraphXml(block, styles, includeOpaque, takeStyleOptions()));
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
      blockXml.push(listRunXml(blocks.slice(i, end) as Paragraph[], styles, includeOpaque, takeStyleOptions()));
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
  const body = blockXml.length > 0 ? blockXml.join("") : "<text:p/>";
  const fontFaceDecls = includeOpaque ? (model.sourceXml?.odtFontFaceDecls ?? "") : "";
  const preservedStyles = includeOpaque ? (model.sourceXml?.odtAutomaticStyles ?? []) : [];
  return (
    `${XML_DECLARATION}<office:document-content${DOCUMENT_XMLNS} office:version="1.2">` +
    fontFaceDecls +
    `<office:automatic-styles>${styles.defs.join("")}${preservedStyles.join("")}</office:automatic-styles>` +
    `<office:body><office:text>${body}</office:text></office:body>` +
    "</office:document-content>"
  );
}

/** One named paragraph style:style definition (SPEC-019) — the id as
 * style:name (with the display name alongside when it differs), the parent
 * link, the outline level that makes it a heading style, and the style's
 * own paragraph/text properties. */
function namedStyleDefinitionXml(definition: NamedStyle): string {
  let attrs = `style:name="${escapeXml(definition.id)}"`;
  if (definition.name !== undefined && definition.name !== definition.id) {
    attrs += ` style:display-name="${escapeXml(definition.name)}"`;
  }
  attrs += ' style:family="paragraph"';
  if (definition.basedOn !== undefined) {
    attrs += ` style:parent-style-name="${escapeXml(definition.basedOn)}"`;
  }
  if (definition.headingLevel !== undefined) {
    attrs += ` style:default-outline-level="${definition.headingLevel}"`;
  }
  const properties =
    paragraphPropertiesXml(definition.paragraphFormatting ?? {}) +
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
  return {
    defs,
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
    tableNamesFor(): TableNames {
      throw new Error("tables are not supported in header/footer content (SPEC-029)");
    },
    graphicStyleNameFor(): string | undefined {
      throw new Error("images are not supported in header/footer content (SPEC-029)");
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
function pageLayoutXml(name: string, pageSize: PageSize, margins: Margins): string {
  // ODF stores landscape pages with width/height already swapped; the
  // orientation attribute just marks them, mirroring the .docx writer's
  // w:orient.
  const orientation =
    pageSize.widthCm > pageSize.heightCm ? ' style:print-orientation="landscape"' : "";
  return (
    `<style:page-layout style:name="${escapeXml(name)}">` +
    `<style:page-layout-properties fo:page-width="${cmToOdfLength(pageSize.widthCm)}"` +
    ` fo:page-height="${cmToOdfLength(pageSize.heightCm)}"${orientation}` +
    ` fo:margin-top="${cmToOdfLength(margins.topCm)}" fo:margin-right="${cmToOdfLength(margins.rightCm)}"` +
    ` fo:margin-bottom="${cmToOdfLength(margins.bottomCm)}" fo:margin-left="${cmToOdfLength(margins.leftCm)}"/>` +
    "</style:page-layout>"
  );
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
export function buildStylesXml(model: DocumentModel): string {
  const { fontFamily, fontSizePt } = model.defaultStyles;
  const includeOpaque = model.format === "odt";
  const namedStyles = (model.styles ?? [])
    .filter((style) => style.name !== undefined)
    .map(namedStyleDefinitionXml)
    .join("");
  const pageLayouts = model.sections
    .map((_, index) => pageLayoutXml(sectionLayoutName(index), sectionPageSize(model, index), sectionMargins(model, index)))
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
    `<style:text-properties fo:font-family="${escapeXml(fontFamily)}" fo:font-size="${ptToOdfLength(fontSizePt)}"/>` +
    "</style:default-style>" +
    namedStyles +
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
    const imported = stylesDoc.importNode(wrapper.documentElement.firstElementChild!, true);
    const existing = Array.from(officeStyles.children).find(
      (el) =>
        el.namespaceURI === STYLE_NAMESPACE &&
        el.localName === "style" &&
        el.getAttributeNS(STYLE_NAMESPACE, "family") === "paragraph" &&
        el.getAttributeNS(STYLE_NAMESPACE, "name") === definition.id,
    );
    if (existing) {
      officeStyles.replaceChild(imported, existing);
    } else {
      officeStyles.appendChild(imported);
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
      pageLayoutXml(layoutName, sectionPageSize(model, index), sectionMargins(model, index)),
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

/** A timestamp in the W3CDTF profile ODF metadata uses (whole seconds,
 * UTC), e.g. "2026-08-03T12:00:00Z". */
function toW3CDTF(date: Date): string {
  return date.toISOString().replace(/\.\d+Z$/, "Z");
}

/** meta.xml with the document title (AC-19) and dc:date, ODF's modification
 * timestamp (AC-20). */
export function buildMetaXml(title: string, modified: Date): string {
  return (
    `${XML_DECLARATION}<office:document-meta xmlns:office="${OFFICE_NAMESPACE}"` +
    ` xmlns:dc="${DC_NAMESPACE}" office:version="1.2">` +
    `<office:meta><dc:title>${escapeXml(title)}</dc:title>` +
    `<dc:date>${toW3CDTF(modified)}</dc:date></office:meta>` +
    "</office:document-meta>"
  );
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
  model: DocumentModel,
  options: WriteOdtOptions = {},
): Promise<WriteResult> {
  const modified = options.modified ?? new Date();
  const title = model.title.trim().length > 0 ? model.title : UNTITLED_ODT_TITLE;
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
  const content = buildContentXml(model, mediaPlan, sectionNames.masters);

  if (preservedParts.size === 0) {
    const bytes = await zip([
      // ODF requires the mimetype entry first and STOREd so format sniffers
      // can read it from the raw bytes at a fixed offset.
      { name: "mimetype", data: encoder.encode(ODT_MIMETYPE), store: true },
      { name: "content.xml", data: encoder.encode(content) },
      { name: "styles.xml", data: encoder.encode(buildStylesXml(model)) },
      { name: "meta.xml", data: encoder.encode(buildMetaXml(title, modified)) },
      { name: "settings.xml", data: encoder.encode(buildSettingsXml()) },
      { name: "META-INF/manifest.xml", data: encoder.encode(buildManifestXml(mediaPlan.manifestEntries)) },
      ...mediaPlan.zipEntries,
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
  }

  const entries: ZipEntryInput[] = [
    { name: "mimetype", data: preservedParts.get("mimetype") ?? encoder.encode(ODT_MIMETYPE), store: true },
    { name: "content.xml", data: encoder.encode(content) },
  ];
  for (const [name, data] of preservedParts) {
    if (name === "mimetype" || name === "content.xml" || name === "meta.xml") {
      continue; // Already emitted or rebuilt below.
    }
    if (name === "META-INF/manifest.xml") {
      entries.push({ name, data: encoder.encode(withManifestEntries(decoder.decode(data), addedPaths)) });
    } else if (patchedStyles !== null && name === "styles.xml") {
      entries.push({ name, data: encoder.encode(patchedStyles) });
    } else {
      entries.push({ name, data });
    }
  }
  entries.push({ name: "meta.xml", data: encoder.encode(buildMetaXml(title, modified)) });
  entries.push(...mediaPlan.zipEntries);
  if (addSettings) {
    entries.push({ name: "settings.xml", data: encoder.encode(buildSettingsXml()) });
  }

  return { bytes: await zip(entries), dropped: [] };
}
