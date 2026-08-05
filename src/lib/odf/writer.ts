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
  DocumentModel,
  Paragraph,
  ParagraphFormatting,
  Run,
  RunFormatting,
  Table,
} from "../documentModel";
import { collectDropRecords, opaquePartBytes, type WriteResult } from "../opaque";
import {
  BULLET_CHARS,
  clampListLevel,
  LIST_MARKER_WIDTH_CM,
  planListLevels,
  type PlannedListLevel,
} from "../lists";
import { zip, type ZipEntryInput } from "../docx/zipWriter";
import { ODT_MIMETYPE } from "./reader";
import { HIGHLIGHT_HEX } from "./properties";
import { FO_NAMESPACE, OFFICE_NAMESPACE, STYLE_NAMESPACE, TABLE_NAMESPACE, TEXT_NAMESPACE } from "./xml";
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
 * property off, and dropping that would change the re-opened model. */
function textPropertiesXml(formatting: RunFormatting): string {
  const attrs: string[] = [];
  if (formatting.fontFamily !== undefined) {
    attrs.push(`fo:font-family="${escapeXml(formatting.fontFamily)}"`);
  }
  if (formatting.fontSizePt !== undefined) {
    attrs.push(`fo:font-size="${ptToOdfLength(formatting.fontSizePt)}"`);
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
 * when there is nothing to say. */
function paragraphPropertiesXml(formatting: ParagraphFormatting): string {
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
  if (rtl) {
    attrs.push('style:writing-mode="rl-tb"');
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
/** The names a serialised table references: its `table:name`, its
 * table-family style, one table-column style per column, and the style
 * generated cells share. `cellStyleName` is a lazy accessor so a table
 * whose cells all carry preserved names never emits an unreferenced
 * definition. */
interface TableNames {
  name: string;
  styleName: string;
  columnStyleNames: string[];
  /** True when the table re-uses its source's own style name — its cells
   * then keep exactly the style references they came with (possibly none),
   * so a same-format save reproduces the source. */
  preserved: boolean;
  cellStyleName(): string;
}

interface AutomaticStyles {
  /** Serialised style:style elements, in first-use order. */
  defs: string[];
  paragraphNameFor(paragraph: Paragraph): string | undefined;
  textNameFor(formatting: RunFormatting): string;
  listNameFor(levels: PlannedListLevel[]): string;
  tableNamesFor(table: Table, preserveNames: boolean): TableNames;
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

    paragraphNameFor(paragraph: Paragraph): string | undefined {
      const propertiesXml = paragraphPropertiesXml(paragraph.formatting ?? {});
      // Nothing to reference and nothing to say: emit no style at all, so
      // the paragraph re-opens without a styleId, matching the model.
      if (!paragraph.styleId && propertiesXml === "") {
        return undefined;
      }
      const key = `${paragraph.styleId ?? ""} ${propertiesXml}`;
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
          ? `<style:style style:name="${escapeXml(name)}" style:family="paragraph"/>`
          : `<style:style style:name="${escapeXml(name)}" style:family="paragraph">${propertiesXml}</style:style>`,
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
      // automatic styles — and only nameless parts (added in the editor)
      // get generated styles. A cross-format save generates everything.
      const totalCm = table.columnWidthsCm.reduce((sum, cm) => sum + cm, 0);
      let styleName = preserveNames ? table.odtStyleName : undefined;
      let base: string;
      if (styleName === undefined) {
        base = claimName("Table", "");
        styleName = base;
        defs.push(
          `<style:style style:name="${escapeXml(base)}" style:family="table">` +
            `<style:table-properties style:width="${cmToOdfLength(totalCm)}"/>` +
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
      // One shared style for cells without a preserved name — LibreOffice's
      // default cell look (hairline border, small padding) so an inserted
      // table is visibly a table in any consumer. Lazy: claimed only when a
      // cell actually needs it, so a fully preserved table's styles are
      // byte-identical to its source.
      let generatedCellStyleName: string | undefined;
      return {
        name: table.odtName ?? base,
        styleName,
        columnStyleNames,
        preserved: preserveNames && table.odtStyleName !== undefined,
        cellStyleName(): string {
          if (generatedCellStyleName === undefined) {
            generatedCellStyleName = claimName(base, ".Cell");
            defs.push(
              `<style:style style:name="${escapeXml(generatedCellStyleName)}" style:family="table-cell">` +
                '<style:table-cell-properties fo:padding="0.097cm" fo:border="0.05pt solid #000000"/>' +
                "</style:style>",
            );
          }
          return generatedCellStyleName;
        },
      };
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
): string {
  const styleName = styles.listNameFor(planListLevels(paragraphs.map((p) => p.list!)));
  let xml = "";
  let depth = 0;
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
    xml += paragraphXml(paragraph, styles, includeOpaque);
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

function runXml(run: Run, styles: AutomaticStyles, state: WhitespaceState): string {
  const styleName = styles.textNameFor(run.formatting);
  return `<text:span text:style-name="${styleName}">${runContentXml(run.text, state)}</text:span>`;
}

/** The paragraph's inline content: its runs, with any preserved opaque
 * inline nodes re-emitted verbatim at their anchored positions (SPEC-008).
 * The whitespace state passes through opaque nodes unchanged — elements are
 * transparent to ODF whitespace collapsing, so the reader's view matches. */
function inlineContentXml(
  paragraph: Paragraph,
  styles: AutomaticStyles,
  includeOpaque: boolean,
): string {
  const state: WhitespaceState = { lastChar: "" };
  const opaque = includeOpaque ? (paragraph.opaqueInlines ?? []) : [];
  if (opaque.length === 0) {
    return paragraph.runs.map((run) => runXml(run, styles, state)).join("");
  }
  let content = "";
  for (let i = 0; i <= paragraph.runs.length; i++) {
    for (const node of opaque) {
      if (Math.min(node.runIndex, paragraph.runs.length) === i) {
        content += node.xml;
      }
    }
    if (i < paragraph.runs.length) {
      content += runXml(paragraph.runs[i], styles, state);
    }
  }
  return content;
}

/** A paragraph as text:p — or text:h with text:outline-level when the model
 * carries a heading level, so the level survives re-opening regardless of
 * the style's definition. */
function paragraphXml(paragraph: Paragraph, styles: AutomaticStyles, includeOpaque: boolean): string {
  const styleName = styles.paragraphNameFor(paragraph);
  const tag = paragraph.headingLevel !== undefined ? "text:h" : "text:p";
  let attrs = styleName !== undefined ? ` text:style-name="${escapeXml(styleName)}"` : "";
  if (paragraph.headingLevel !== undefined) {
    attrs += ` text:outline-level="${paragraph.headingLevel}"`;
  }
  const content = inlineContentXml(paragraph, styles, includeOpaque);
  return content.length > 0 ? `<${tag}${attrs}>${content}</${tag}>` : `<${tag}${attrs}/>`;
}

function isListParagraph(block: Block): boolean {
  return block.type === "paragraph" && block.list !== undefined;
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

/** One table block as table:table markup (SPEC-022): column definitions
 * referencing width-carrying styles, then rows of string-typed cells. The
 * table: namespace is declared on the element itself — see TABLE_XMLNS. */
function tableXml(table: Table, styles: AutomaticStyles, includeOpaque: boolean): string {
  const names = styles.tableNamesFor(table, includeOpaque);
  let xml =
    `<table:table${TABLE_XMLNS} table:name="${escapeXml(names.name)}"` +
    ` table:style-name="${escapeXml(names.styleName)}">`;
  for (const columnName of names.columnStyleNames) {
    xml += `<table:table-column table:style-name="${escapeXml(columnName)}"/>`;
  }
  for (const row of table.rows) {
    const rowStyle =
      includeOpaque && row.odtStyleName !== undefined
        ? ` table:style-name="${escapeXml(row.odtStyleName)}"`
        : "";
    xml += `<table:table-row${rowStyle}>`;
    for (const cell of row.cells) {
      // A preserved table's cells keep exactly the style references they
      // came with — including none at all; only a table this writer styles
      // itself gets the generated bordered cell style.
      const cellStyleName =
        includeOpaque && cell.odtStyleName !== undefined
          ? cell.odtStyleName
          : names.preserved
            ? undefined
            : names.cellStyleName();
      const cellStyleAttr =
        cellStyleName !== undefined ? ` table:style-name="${escapeXml(cellStyleName)}"` : "";
      xml +=
        `<table:table-cell${cellStyleAttr} office:value-type="string">` +
        cellContentXml(cell.blocks, styles, includeOpaque) +
        "</table:table-cell>";
    }
    xml += "</table:table-row>";
  }
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
export function buildContentXml(model: DocumentModel): string {
  const includeOpaque = model.format === "odt";
  const styles = createAutomaticStyles(model);
  const blockXml: string[] = [];
  for (const section of model.sections) {
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
        blockXml.push(tableXml(block, styles, includeOpaque));
        i++;
        continue;
      }
      if (!block.list) {
        blockXml.push(paragraphXml(block, styles, includeOpaque));
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
      blockXml.push(listRunXml(blocks.slice(i, end) as Paragraph[], styles, includeOpaque));
      i = end;
    }
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

/** The styles part: a paragraph default-style carrying the model's default
 * font (the reader's source for defaultStyles.fontFamily/fontSizePt), and
 * the page layout applied through the first master page (AC-15). */
export function buildStylesXml(model: DocumentModel): string {
  const { fontFamily, fontSizePt, pageSize, margins } = model.defaultStyles;
  // ODF stores landscape pages with width/height already swapped; the
  // orientation attribute just marks them, mirroring the .docx writer's
  // w:orient.
  const orientation =
    pageSize.widthCm > pageSize.heightCm ? ' style:print-orientation="landscape"' : "";
  return (
    `${XML_DECLARATION}<office:document-styles${DOCUMENT_XMLNS} office:version="1.2">` +
    "<office:styles>" +
    '<style:default-style style:family="paragraph">' +
    `<style:text-properties fo:font-family="${escapeXml(fontFamily)}" fo:font-size="${ptToOdfLength(fontSizePt)}"/>` +
    "</style:default-style>" +
    "</office:styles>" +
    "<office:automatic-styles>" +
    '<style:page-layout style:name="pm1">' +
    `<style:page-layout-properties fo:page-width="${cmToOdfLength(pageSize.widthCm)}"` +
    ` fo:page-height="${cmToOdfLength(pageSize.heightCm)}"${orientation}` +
    ` fo:margin-top="${cmToOdfLength(margins.topCm)}" fo:margin-right="${cmToOdfLength(margins.rightCm)}"` +
    ` fo:margin-bottom="${cmToOdfLength(margins.bottomCm)}" fo:margin-left="${cmToOdfLength(margins.leftCm)}"/>` +
    "</style:page-layout>" +
    "</office:automatic-styles>" +
    "<office:master-styles>" +
    '<style:master-page style:name="Standard" style:page-layout-name="pm1"/>' +
    "</office:master-styles>" +
    "</office:document-styles>"
  );
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

/** META-INF/manifest.xml listing every part (AC-8). The mimetype entry is
 * not listed — per ODF its media type belongs to the "/" entry. */
export function buildManifestXml(): string {
  const part = (path: string) =>
    `<manifest:file-entry manifest:full-path="${path}" manifest:media-type="text/xml"/>`;
  return (
    `${XML_DECLARATION}<manifest:manifest xmlns:manifest="${MANIFEST_NAMESPACE}" manifest:version="1.2">` +
    `<manifest:file-entry manifest:full-path="/" manifest:version="1.2" manifest:media-type="${ODT_MIMETYPE}"/>` +
    part("content.xml") +
    part("styles.xml") +
    part("meta.xml") +
    part("settings.xml") +
    "</manifest:manifest>"
  );
}

export interface WriteOdtOptions {
  /** The dc:date modification timestamp; defaults to now. Injectable so
   * tests can assert an exact meta.xml. */
  modified?: Date;
}

/** Ensures a preserved manifest lists the parts this save (re)builds —
 * needed when the source package lacked meta.xml or settings.xml. The patch
 * is textual and no-ops gracefully on producers that bound the manifest
 * namespace to another prefix (their packages carry the parts already). */
function withManifestEntries(manifestXml: string, paths: string[]): string {
  let out = manifestXml;
  for (const path of paths) {
    if (!out.includes(`manifest:full-path="${path}"`)) {
      out = out.replace(
        /<\/manifest:manifest>\s*$/,
        `<manifest:file-entry manifest:full-path="${path}" manifest:media-type="text/xml"/>` +
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
  const content = buildContentXml(model);

  const preservedParts = model.format === "odt" ? opaquePartBytes(model) : new Map<string, Uint8Array>();
  if (preservedParts.size === 0) {
    const bytes = await zip([
      // ODF requires the mimetype entry first and STOREd so format sniffers
      // can read it from the raw bytes at a fixed offset.
      { name: "mimetype", data: encoder.encode(ODT_MIMETYPE), store: true },
      { name: "content.xml", data: encoder.encode(content) },
      { name: "styles.xml", data: encoder.encode(buildStylesXml(model)) },
      { name: "meta.xml", data: encoder.encode(buildMetaXml(title, modified)) },
      { name: "settings.xml", data: encoder.encode(buildSettingsXml()) },
      { name: "META-INF/manifest.xml", data: encoder.encode(buildManifestXml()) },
    ]);
    return { bytes, dropped: collectDropRecords(model, "odt") };
  }

  // Same-format save of an opened .odt: preserved parts verbatim, rebuilt
  // content.xml/meta.xml, and settings.xml generated only when the source
  // package had none (the manifest is patched to match).
  const decoder = new TextDecoder("utf-8");
  const addSettings = !preservedParts.has("settings.xml");
  const addedPaths = ["content.xml", "meta.xml", ...(addSettings ? ["settings.xml"] : [])];

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
    } else {
      entries.push({ name, data });
    }
  }
  entries.push({ name: "meta.xml", data: encoder.encode(buildMetaXml(title, modified)) });
  if (addSettings) {
    entries.push({ name: "settings.xml", data: encoder.encode(buildSettingsXml()) });
  }

  return { bytes: await zip(entries), dropped: [] };
}
