// Serialises a DocumentModel into a WordprocessingML OPC package — the
// inverse of reader.ts. Every field the reader maps into the model is
// written back out (text, run formatting, paragraph formatting, style
// references, section page setup), each as the exact inverse of the
// reader's mapping, so open -> save -> re-open reproduces the model.
//
// Emission choices that keep the round trip exact:
//  - Effective formatting is written as *direct* formatting on each run and
//    paragraph. The reader resolves style formatting into the model (a
//    paragraph's `formatting` is docDefaults + style chain + direct), so
//    re-emitting the merged result directly reproduces it without having to
//    reconstruct the original style definitions.
//  - Referenced paragraph styles are declared in styles.xml by id and name
//    only, with no formatting of their own — the reader merges style
//    formatting into paragraphs, so a formatting-carrying style would be
//    applied twice on re-open and break equality.
//  - docDefaults carries the model's default font family/size; that is the
//    one place the reader looks for `defaultStyles.fontFamily`/`fontSizePt`.
//  - Child elements inside w:pPr/w:rPr/w:sectPr are emitted in the schema's
//    sequence order (see validate.ts), which is what Word's strict reader
//    checks before offering to "repair" a file.
import type {
  Block,
  BorderLineStyle,
  BorderSide,
  CellBorders,
  CellVerticalAlign,
  DocumentModel,
  DocumentProperties,
  FieldSpec,
  HeaderFooterKind,
  HeaderFooterPart,
  Image,
  Margins,
  NoteContent,
  NoteKind,
  NoteNumberFormat,
  NoteSettings,
  PageNumberFormat,
  NamedStyle,
  PageSize,
  Paragraph,
  ParagraphAlignment,
  ParagraphFormatting,
  Run,
  RunFormatChange,
  RunFormatting,
  Section,
  SectionColumns,
  SectionStart,
  Table,
  TableCell,
  TableRow,
  TableOfContents,
  TocEntry,
  TocOptions,
} from "../documentModel";
import {
  DEFAULT_TITLE,
  isMergeChipField,
  resolveColumnWidthsCm,
  sectionContentWidthCm,
  sectionMargins,
  sectionPageSize,
} from "../documentModel";
import { chipLabel } from "../mergeChips";
import { collectDropRecords, opaquePartBytes, type WriteResult } from "../opaque";
import { equalStructural } from "../equality";
import { inlineWriteLayout, linkTargetOf } from "../links";
import { TOC_PLACEHOLDER_TEXT } from "../toc";
import { referencedNotes } from "../notes";
import { replyParagraphText, visibleCommentThreads, type VisibleCommentThread } from "../comments";
import {
  childElements,
  parseRelationships,
  parseXml,
  relsPathFor,
  resolvePartPath,
  serializeElement,
  wordAttr,
  WORD_NAMESPACE,
} from "./xml";
import { catalogToNamedStyles, parseStylesCatalog } from "./styles";
import {
  cmToTwips,
  cssColorToOoxml,
  ptToEighthPoints,
  ptToHalfPoints,
  ptToTwentiethsOfPoint,
} from "./units";
import {
  buildNumberingXml,
  mergeNumberingXml,
  nextNumberingIds,
  NUMBERING_CONTENT_TYPE,
  NUMBERING_RELATIONSHIP,
  cellNumberingRefFor,
  numberingRefFor,
  planNumbering,
  type NumberingPlan,
  type NumberingRef,
} from "./numbering";
import { CONTENT_TYPE_BY_EXTENSION, drawingXml } from "./drawing";
import {
  nextFreeRelId,
  planMediaWrite,
  planMediaWriteForBlocks,
  takenRelIdsOf,
  type MediaWritePlan,
} from "./media";
import { zip, type ZipEntryInput } from "./zipWriter";
import { buildCorePropertiesXml, buildCustomPropertiesXml, type CorePropertiesInput } from "./docProps";
import { stripPersonalInfo } from "../personalInfo";

export { buildCorePropertiesXml, buildCustomPropertiesXml } from "./docProps";

export const DOCX_MIME_TYPE =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

const CONTENT_TYPES_NAMESPACE = "http://schemas.openxmlformats.org/package/2006/content-types";
const RELATIONSHIPS_NAMESPACE = "http://schemas.openxmlformats.org/package/2006/relationships";
const OFFICE_DOCUMENT_RELATIONSHIP =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
const STYLES_RELATIONSHIP =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
const CORE_PROPERTIES_RELATIONSHIP =
  "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
const CUSTOM_PROPERTIES_RELATIONSHIP =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties";
const CUSTOM_PROPERTIES_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.custom-properties+xml";

// Word's default header/footer distance (0.5"), in twips. The model does not
// carry header/footer distances, but CT_PageMar requires the attributes.
const HEADER_FOOTER_TWIPS = 720;

/** The suggested download file name: `<title>.docx` (AC-2), falling back to
 * the product's default title when the title is empty/whitespace. */
export function docxFileName(title: string): string {
  const trimmed = title.trim();
  return `${trimmed.length > 0 ? trimmed : DEFAULT_TITLE}.docx`;
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

// w:jc values are the inverse of the reader's ALIGNMENT_MAP ("both" is
// OOXML's name for justified).
const JC_VALUES: Record<ParagraphAlignment, string> = {
  left: "left",
  center: "center",
  right: "right",
  justify: "both",
};

function trackedChangeAttrs(id: number, author: string, date: string): string {
  return ` w:id="${id}" w:author="${escapeXml(author)}" w:date="${escapeXml(date)}"`;
}

/** A run's formatting, plus — with `formatChange` set (SPEC-039) — a
 * trailing `w:rPrChange` recording the pre-change formatting (`previous`)
 * under a fresh numeric id from `ctx.changeIds`. */
function runPropertiesXml(
  formatting: RunFormatting,
  styleId?: string,
  formatChange?: RunFormatChange,
  ctx?: InlineWriteContext,
): string {
  const parts: string[] = [];
  // Emitted in CT_RPr sequence order: rStyle, rFonts, b, i, strike, noProof,
  // color, sz, highlight, u, shd, vertAlign, lang. Explicit false values are
  // written as w:val="0"/"none" — the model only holds false where the
  // source explicitly switched the property off, and dropping that would
  // change the re-opened model.
  if (styleId !== undefined) {
    parts.push(`<w:rStyle w:val="${escapeXml(styleId)}"/>`);
  }
  if (formatting.fontFamily !== undefined) {
    const family = escapeXml(formatting.fontFamily);
    parts.push(`<w:rFonts w:ascii="${family}" w:hAnsi="${family}"/>`);
  }
  if (formatting.bold !== undefined) {
    parts.push(formatting.bold ? "<w:b/>" : '<w:b w:val="0"/>');
  }
  if (formatting.italic !== undefined) {
    parts.push(formatting.italic ? "<w:i/>" : '<w:i w:val="0"/>');
  }
  if (formatting.strikethrough !== undefined) {
    parts.push(formatting.strikethrough ? "<w:strike/>" : '<w:strike w:val="0"/>');
  }
  // SPEC-035: "do not check spelling".
  if (formatting.noProof !== undefined) {
    parts.push(formatting.noProof ? "<w:noProof/>" : '<w:noProof w:val="0"/>');
  }
  const color = cssColorToOoxml(formatting.color);
  if (color !== undefined) {
    parts.push(`<w:color w:val="${color}"/>`);
  }
  if (formatting.fontSizePt !== undefined) {
    parts.push(`<w:sz w:val="${ptToHalfPoints(formatting.fontSizePt)}"/>`);
  }
  // Highlight (SPEC-013): a palette name maps to w:highlight, whose value
  // set is closed; an explicit "No colour" is w:val="none"; any other
  // colour has no ST_HighlightColor encoding and degrades to run shading.
  // The readers normalise palette-coloured fills back to names, so a bare
  // "#rrggbb" here is guaranteed to be outside the palette.
  const highlightHex =
    typeof formatting.highlight === "string" && formatting.highlight.startsWith("#")
      ? cssColorToOoxml(formatting.highlight)
      : undefined;
  if (formatting.highlight === false) {
    parts.push('<w:highlight w:val="none"/>');
  } else if (
    typeof formatting.highlight === "string" &&
    !formatting.highlight.startsWith("#")
  ) {
    parts.push(`<w:highlight w:val="${escapeXml(formatting.highlight)}"/>`);
  }
  if (formatting.underline !== undefined) {
    parts.push(`<w:u w:val="${formatting.underline ? "single" : "none"}"/>`);
  }
  // One w:shd slot: an out-of-palette highlight claims it; otherwise the
  // preserved source shading (kept beside a palette highlight) is
  // re-emitted. When an edit leaves both (the user picked a custom
  // highlight on a run that carried shading), the visible highlight wins.
  const shadingHex = highlightHex ?? cssColorToOoxml(formatting.shadingFill);
  if (shadingHex !== undefined) {
    parts.push(`<w:shd w:val="clear" w:color="auto" w:fill="${shadingHex}"/>`);
  }
  if (formatting.vertAlign !== undefined) {
    parts.push(`<w:vertAlign w:val="${formatting.vertAlign}"/>`);
  }
  // SPEC-035: run language, e.g. from the language selector applied to a
  // selection, or a source style/run that already diverged from the
  // document default (see docx/styles.ts's docDefault lift-out — the
  // document default itself is written into docDefaults, not here).
  if (formatting.lang !== undefined) {
    parts.push(`<w:lang w:val="${escapeXml(formatting.lang)}"/>`);
  }
  if (formatChange && ctx) {
    const id = changeNumericId(ctx, formatChange.changeId);
    const previous = runPropertiesXml(formatChange.previous, undefined);
    parts.push(`<w:rPrChange${trackedChangeAttrs(id, formatChange.author, formatChange.date)}>${previous}</w:rPrChange>`);
  }
  return parts.length > 0 ? `<w:rPr>${parts.join("")}</w:rPr>` : "";
}

/** A run's text as w:t/w:tab/w:br children — the inverse of the reader's
 * literal "\t"/"\n" mapping. xml:space="preserve" keeps leading/trailing
 * spaces through Word's whitespace normalisation. `textTag` is "delText"
 * for a tracked deletion (SPEC-039): Word's substitute for w:t inside
 * w:del, same content model otherwise. */
function runContentXml(text: string, textTag: "t" | "delText" = "t"): string {
  let xml = "";
  let buffer = "";
  const flush = () => {
    if (buffer.length > 0) {
      const preserve = buffer.trim() !== buffer ? ' xml:space="preserve"' : "";
      xml += `<w:${textTag}${preserve}>${escapeXml(buffer)}</w:${textTag}>`;
      buffer = "";
    }
  };
  for (const char of text) {
    if (char === "\t") {
      flush();
      xml += "<w:tab/>";
    } else if (char === "\n") {
      flush();
      xml += "<w:br/>";
    } else if (char === "‑") {
      flush();
      xml += "<w:noBreakHyphen/>";
    } else if (char === "­") {
      flush();
      xml += "<w:softHyphen/>";
    } else {
      buffer += char;
    }
  }
  flush();
  return xml;
}

// The field-instruction switch for each numbering format (SPEC-029
// AC-16..18) — absent (decimal) needs none, OOXML's own default.
const FIELD_FORMAT_SWITCH: Record<PageNumberFormat, string> = {
  decimal: "",
  lowerRoman: " \\* roman",
  lowerAlpha: " \\* alphabetic",
};

function fieldInstr(field: FieldSpec): string {
  if (isMergeChipField(field)) {
    // Always quoted (SPEC-046 "Chip serialisation"), so the writer has no
    // has-a-space branch and the reader has one pattern.
    return `MERGEFIELD "${chipLabel(field)}"`;
  }
  const kind = field.kind === "page" ? "PAGE" : "NUMPAGES";
  return `${kind}${FIELD_FORMAT_SWITCH[field.format]}`;
}

/** A run, or — when it carries a field spec (SPEC-029) — a `w:fldSimple`
 * wrapping one: `w:fldSimple` schema-wraps a cached-result run the same
 * shape an ordinary run would be, so the run's own formatting/placeholder
 * text still goes through runPropertiesXml/runContentXml unchanged; only
 * the instruction wrapper differs. A note reference run (SPEC-031) emits a
 * `w:footnoteReference`/`w:endnoteReference` instead of any text — Word
 * renders the number itself — resolved through `noteIds` (the notes write
 * plan's numeric ids); a dangling reference emits nothing.
 *
 * A tracked-change run (SPEC-039) is wrapped in `w:ins`/`w:del` — both when
 * nested (AC "editing over an existing tracked insertion... produces a
 * nested change"), `w:ins` outermost since the insertion happened first —
 * and a tracked deletion's text is `w:delText`, never `w:t`. Field/note
 * runs are never wrapped: their content is a computed placeholder, not
 * text a tracked edit could sensibly mark. */
function runXml(run: Run, noteIds?: Map<string, number>, ctx?: InlineWriteContext): string {
  if (run.note) {
    const id = noteIds?.get(run.note.noteId);
    if (id === undefined) {
      return "";
    }
    const tag = run.note.kind === "footnote" ? "footnoteReference" : "endnoteReference";
    return `<w:r>${runPropertiesXml(run.formatting, run.styleId)}<w:${tag} w:id="${id}"/></w:r>`;
  }
  const deleted = run.changes?.delete;
  const inserted = run.changes?.insert;
  const formatChange = run.changes?.formatChange;
  const rPr = runPropertiesXml(run.formatting, run.styleId, formatChange, ctx);
  if (run.field) {
    // The cached run carries the canonical chip text (not the model's
    // one-character placeholder) for a merge chip, so a consumer that does
    // not evaluate fields still shows `{{client_name}}`; the reader
    // rebuilds from the instruction and ignores this cached text, exactly
    // as it does for PAGE (SPEC-046 "Chip serialisation").
    const cachedText = isMergeChipField(run.field) ? chipLabel(run.field) : run.text;
    const fieldContent = runContentXml(cachedText, deleted ? "delText" : "t");
    const fieldInner = `<w:r>${rPr}${fieldContent}</w:r>`;
    return `<w:fldSimple w:instr="${escapeXml(fieldInstr(run.field))}">${fieldInner}</w:fldSimple>`;
  }
  const content = runContentXml(run.text, deleted ? "delText" : "t");
  const inner = `<w:r>${rPr}${content}</w:r>`;
  let xml = inner;
  if (deleted && ctx) {
    const id = changeNumericId(ctx, deleted.changeId);
    xml = `<w:del${trackedChangeAttrs(id, deleted.author, deleted.date)}>${xml}</w:del>`;
  }
  if (inserted && ctx) {
    const id = changeNumericId(ctx, inserted.changeId);
    xml = `<w:ins${trackedChangeAttrs(id, inserted.author, inserted.date)}>${xml}</w:ins>`;
  }
  return xml;
}

// w:tab/@w:leader values for the model's tab leaders (SPEC-032 AC-23).
const TAB_LEADER_VALUES: Record<string, string> = {
  dot: "dot",
  hyphen: "hyphen",
  underscore: "underscore",
};

/** The pPr children a bare ParagraphFormatting produces — the shared tail
 * of a paragraph's own w:pPr and a style definition's (SPEC-019), in CT_PPr
 * sequence order: tabs, bidi, spacing, ind, jc. */
function paragraphFormattingParts(formatting: ParagraphFormatting): string[] {
  const parts: string[] = [];

  // Custom tab stops (SPEC-032 AC-19..25): one w:tab per stop, positions in
  // twips from the leading text margin.
  if (formatting.tabStops !== undefined && formatting.tabStops.length > 0) {
    const tabs = formatting.tabStops
      .map((stop) => {
        const leader =
          stop.leader !== undefined ? ` w:leader="${TAB_LEADER_VALUES[stop.leader]}"` : "";
        return `<w:tab w:val="${stop.alignment}"${leader} w:pos="${cmToTwips(stop.positionCm)}"/>`;
      })
      .join("");
    parts.push(`<w:tabs>${tabs}</w:tabs>`);
  }

  // The model's indents are logical start/end values, exactly what w:ind's
  // left/right mean, so RTL needs only the w:bidi flag — the indent values
  // are emitted unchanged (SPEC-014).
  if (formatting.direction === "rtl") {
    parts.push("<w:bidi/>");
  }

  const spacingAttrs: string[] = [];
  if (formatting.spaceBeforePt !== undefined) {
    spacingAttrs.push(`w:before="${ptToTwentiethsOfPoint(formatting.spaceBeforePt)}"`);
  }
  if (formatting.spaceAfterPt !== undefined) {
    spacingAttrs.push(`w:after="${ptToTwentiethsOfPoint(formatting.spaceAfterPt)}"`);
  }
  // A single w:line attribute carries either rule; lineSpacingPt wins when
  // both are set, matching the model's documented precedence. The absent
  // lineSpacingRule default is "atLeast" — converted content can never clip
  // tall glyphs — while an explicit "exact" survives as itself (SPEC-014).
  if (formatting.lineSpacingPt !== undefined) {
    const rule = formatting.lineSpacingRule === "exact" ? "exact" : "atLeast";
    spacingAttrs.push(
      `w:line="${ptToTwentiethsOfPoint(formatting.lineSpacingPt)}" w:lineRule="${rule}"`,
    );
  } else if (formatting.lineSpacing !== undefined) {
    spacingAttrs.push(`w:line="${Math.round(formatting.lineSpacing * 240)}" w:lineRule="auto"`);
  }
  if (spacingAttrs.length > 0) {
    parts.push(`<w:spacing ${spacingAttrs.join(" ")}/>`);
  }

  const indAttrs: string[] = [];
  if (formatting.indentStartCm !== undefined) {
    indAttrs.push(`w:left="${cmToTwips(formatting.indentStartCm)}"`);
  }
  if (formatting.indentEndCm !== undefined) {
    indAttrs.push(`w:right="${cmToTwips(formatting.indentEndCm)}"`);
  }
  if (formatting.indentFirstLineCm !== undefined) {
    indAttrs.push(
      formatting.indentFirstLineCm >= 0
        ? `w:firstLine="${cmToTwips(formatting.indentFirstLineCm)}"`
        : `w:hanging="${cmToTwips(-formatting.indentFirstLineCm)}"`,
    );
  }
  if (indAttrs.length > 0) {
    parts.push(`<w:ind ${indAttrs.join(" ")}/>`);
  }

  if (formatting.alignment !== undefined) {
    parts.push(`<w:jc w:val="${JC_VALUES[formatting.alignment]}"/>`);
  }
  return parts;
}

function paragraphPropertiesXml(
  paragraph: Paragraph,
  numRef?: NumberingRef,
  sectPr?: string,
  ctx?: InlineWriteContext,
): string {
  const parts: string[] = [];
  const formatting = paragraph.formatting ?? {};

  // Emitted in CT_PPr sequence order: pStyle, keepNext, keepLines,
  // pageBreakBefore, widowControl, numPr, then the shared formatting parts,
  // then outlineLvl — with a section-ending w:sectPr (SPEC-028 AC-31) last.
  if (paragraph.styleId !== undefined) {
    parts.push(`<w:pStyle w:val="${escapeXml(paragraph.styleId)}"/>`);
  }

  // The SPEC-028 pagination toggles, mirroring the model's canonical form:
  // the keeps and pageBreakBefore only when on (empty on-toggles), widow
  // control only when explicitly off — absent means on in both the model
  // and OOXML, so nothing is written for the default.
  if (formatting.keepNext === true) {
    parts.push("<w:keepNext/>");
  }
  if (formatting.keepLines === true) {
    parts.push("<w:keepLines/>");
  }
  if (formatting.pageBreakBefore === true) {
    parts.push("<w:pageBreakBefore/>");
  }
  if (formatting.widowControl === false) {
    parts.push('<w:widowControl w:val="0"/>');
  }

  // The list reference (SPEC-018 AC-18): the numbering plan assigned this
  // paragraph its num instance and 0-based level.
  if (numRef) {
    parts.push(
      `<w:numPr><w:ilvl w:val="${numRef.ilvl}"/><w:numId w:val="${numRef.numId}"/></w:numPr>`,
    );
  }

  parts.push(...paragraphFormattingParts(formatting));
  if (paragraph.headingLevel !== undefined) {
    parts.push(`<w:outlineLvl w:val="${paragraph.headingLevel - 1}"/>`);
  }
  // The paragraph mark's own tracked insert/delete (SPEC-039): nested
  // inside the mark's own w:rPr, distinct from the paragraph's other
  // children — a split's new mark, or a pending merge's mark-for-deletion.
  const markInsert = paragraph.changes?.markInsert;
  const markDelete = paragraph.changes?.markDelete;
  if ((markInsert || markDelete) && ctx) {
    const markParts: string[] = [];
    if (markInsert) {
      const id = changeNumericId(ctx, markInsert.changeId);
      markParts.push(`<w:ins${trackedChangeAttrs(id, markInsert.author, markInsert.date)}/>`);
    }
    if (markDelete) {
      const id = changeNumericId(ctx, markDelete.changeId);
      markParts.push(`<w:del${trackedChangeAttrs(id, markDelete.author, markDelete.date)}/>`);
    }
    parts.push(`<w:rPr>${markParts.join("")}</w:rPr>`);
  }
  if (sectPr !== undefined) {
    parts.push(sectPr);
  }
  const formatChange = paragraph.changes?.formatChange;
  if (formatChange && ctx) {
    const id = changeNumericId(ctx, formatChange.changeId);
    const previousFormatting = formatChange.previous ?? {};
    const previous = `<w:pPr>${paragraphFormattingParts(previousFormatting).join("")}</w:pPr>`;
    parts.push(`<w:pPrChange${trackedChangeAttrs(id, formatChange.author, formatChange.date)}>${previous}</w:pPrChange>`);
  }

  return parts.length > 0 ? `<w:pPr>${parts.join("")}</w:pPr>` : "";
}

// --- hyperlinks and bookmarks (SPEC-025) -------------------------------------

const HYPERLINK_RELATIONSHIP =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
// The `r:` attribute namespace, declared locally on each w:hyperlink the
// same way drawing.ts declares it on a:blip.
const R_ATTRIBUTE_NAMESPACE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/** Visits every paragraph reachable from a block list, descending into table
 * cells — the shared iteration `forEachModelParagraph` (the whole model)
 * and SPEC-029's per-header/footer-part planning (one Block[]) both need. */
function forEachParagraphIn(blocks: Block[], visit: (paragraph: Paragraph) => void): void {
  for (const block of blocks) {
    if (block.type === "table") {
      for (const row of block.rows) {
        for (const cell of row.cells) {
          for (const paragraph of cell.blocks) {
            visit(paragraph);
          }
        }
      }
    } else if (block.type === "paragraph") {
      visit(block);
    }
  }
}

function forEachModelParagraph(model: DocumentModel, visit: (paragraph: Paragraph) => void): void {
  for (const section of model.sections) {
    forEachParagraphIn(section.blocks, visit);
  }
}

export interface HyperlinkWritePlan {
  /** The relationship id a run's external `linkHref` maps to. */
  relIdFor(href: string): string;
  /** `<Relationship>` elements to append to the document rels part. */
  relationshipsXml: string;
  /** The ids this plan claimed — seeds the media planner so the two
   * appenders never collide. */
  relIds: Set<string>;
}

/** One `TargetMode="External"` hyperlink relationship per distinct external
 * link target a block list references, ids allocated past `takenRelIds` —
 * the same additive scheme as planMediaWrite, so a same-format save only
 * ever appends to the preserved rels part. The engine `planHyperlinkWrite`
 * (the whole model) and SPEC-029's per-header/footer-part planning (one
 * Block[], its own independent rels file) both call this. */
function planHyperlinkWriteFor(blocks: Block[], takenRelIds: Set<string>): HyperlinkWritePlan {
  const hrefs: string[] = [];
  const seen = new Set<string>();
  forEachParagraphIn(blocks, (paragraph) => {
    for (const run of paragraph.runs) {
      const href = run.formatting.linkHref;
      if (href !== undefined && !seen.has(href)) {
        seen.add(href);
        hrefs.push(href);
      }
    }
  });
  const relIdOf = new Map<string, string>();
  const relParts: string[] = [];
  const taken = new Set(takenRelIds);
  for (const href of hrefs) {
    const relId = nextFreeRelId(taken);
    taken.add(relId);
    relIdOf.set(href, relId);
    relParts.push(
      `<Relationship Id="${relId}" Type="${HYPERLINK_RELATIONSHIP}" Target="${escapeXml(href)}" TargetMode="External"/>`,
    );
  }
  return {
    relIdFor: (href: string) => relIdOf.get(href) ?? "",
    relationshipsXml: relParts.join(""),
    relIds: new Set(relIdOf.values()),
  };
}

export function planHyperlinkWrite(
  model: DocumentModel,
  takenRelIds: Set<string> = new Set(),
): HyperlinkWritePlan {
  return planHyperlinkWriteFor(
    model.sections.flatMap((section) => section.blocks),
    takenRelIds,
  );
}

/** Document-order numeric ids for w:bookmarkStart/End pairs (SPEC-025). */
function bookmarkIdsFor(model: DocumentModel): Map<string, number> {
  const ids = new Map<string, number>();
  forEachModelParagraph(model, (paragraph) => {
    for (const bookmark of paragraph.bookmarks ?? []) {
      if (!ids.has(bookmark.name)) {
        ids.set(bookmark.name, ids.size + 1);
      }
    }
  });
  return ids;
}

interface InlineWriteContext {
  hyperlinkPlan: HyperlinkWritePlan;
  bookmarkIds: Map<string, number>;
  /** Numeric footnotes.xml/endnotes.xml ids by model noteId (SPEC-031);
   * absent in contexts that can hold no note references (the note parts
   * themselves, headers/footers). */
  noteIds?: Map<string, number>;
  /** Numeric word/comments.xml ids by model threadId (SPEC-038); absent in
   * contexts that can hold no comment anchors (the comments part itself,
   * headers/footers, note content). */
  commentIds?: Map<string, number>;
  /** Numeric `w:id`s for every `w:ins`/`w:del`/`w:rPrChange`/`w:pPrChange`/
   * row-change element (SPEC-039), keyed by model changeId and assigned
   * lazily in document-write order — the same one pool Word itself uses
   * across every tracked-change element kind in a document. */
  changeIds: Map<string, number>;
}

/** The numeric `w:id` for a tracked-change element (SPEC-039): the same
 * changeId always gets the same number within one write, assigned on first
 * encounter in document order — mirrors commentIds' assignment shape. */
function changeNumericId(ctx: InlineWriteContext, changeId: string): number {
  const existing = ctx.changeIds.get(changeId);
  if (existing !== undefined) {
    return existing;
  }
  const id = ctx.changeIds.size + 1;
  ctx.changeIds.set(changeId, id);
  return id;
}

/** The paragraph's inline content: its runs, with any preserved opaque
 * inline nodes re-emitted verbatim at their anchored positions (SPEC-008),
 * consecutive same-target link runs grouped under one `w:hyperlink`
 * (external targets via an r:id relationship, internal via w:anchor), and
 * bookmarks emitted as adjacent `w:bookmarkStart`/`w:bookmarkEnd` pairs at
 * their character offsets (SPEC-025). Emitted on cross-format saves too —
 * links/bookmarks are modelled content, unlike opaque nodes. */
function inlineContentXml(paragraph: Paragraph, includeOpaque: boolean, ctx: InlineWriteContext): string {
  const opaque = includeOpaque ? (paragraph.opaqueInlines ?? []) : [];
  const hasLinks = paragraph.runs.some((run) => linkTargetOf(run.formatting) !== null);
  const hasBookmarks = (paragraph.bookmarks ?? []).length > 0;
  const hasComments = (paragraph.commentAnchors ?? []).length > 0;
  if (opaque.length === 0 && !hasLinks && !hasBookmarks && !hasComments) {
    return paragraph.runs.map((run) => runXml(run, ctx.noteIds, ctx)).join("");
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
  let xml = "";
  let openTarget: { href?: string; anchor?: string } | null = null;
  const closeHyperlink = () => {
    if (openTarget) {
      xml += "</w:hyperlink>";
      openTarget = null;
    }
  };
  for (let i = 0; i <= runs.length; i++) {
    const run = i < runs.length ? runs[i] : null;
    const target = run ? linkTargetOf(run.formatting) : null;
    const opaqueHere = opaqueAtBoundary.get(i) ?? [];
    // An opaque node's XML is arbitrary — never emit it inside a hyperlink;
    // the wrapper is split around it instead (two w:hyperlink elements with
    // the same target are equivalent).
    if (opaqueHere.length > 0) {
      closeHyperlink();
    }
    if (
      openTarget &&
      (!target || target.href !== openTarget.href || target.anchor !== openTarget.anchor)
    ) {
      closeHyperlink();
    }
    xml += opaqueHere.join("");
    for (const name of bookmarksAtBoundary.get(i) ?? []) {
      const id = ctx.bookmarkIds.get(name) ?? 0;
      xml += `<w:bookmarkStart w:id="${id}" w:name="${escapeXml(name)}"/><w:bookmarkEnd w:id="${id}"/>`;
    }
    // Comment ranges (SPEC-038): a dangling anchor (not in ctx.commentIds —
    // its thread turned out not to be visible, see visibleCommentThreads)
    // emits nothing, the same convention runXml's note references follow.
    for (const threadId of commentStartsAtBoundary.get(i) ?? []) {
      const id = ctx.commentIds?.get(threadId);
      if (id !== undefined) {
        xml += `<w:commentRangeStart w:id="${id}"/>`;
      }
    }
    for (const threadId of commentEndsAtBoundary.get(i) ?? []) {
      const id = ctx.commentIds?.get(threadId);
      if (id !== undefined) {
        xml += `<w:commentRangeEnd w:id="${id}"/><w:r><w:commentReference w:id="${id}"/></w:r>`;
      }
    }
    if (!run) {
      break;
    }
    if (target && !openTarget) {
      xml +=
        target.href !== undefined
          ? `<w:hyperlink xmlns:r="${R_ATTRIBUTE_NAMESPACE}" r:id="${ctx.hyperlinkPlan.relIdFor(target.href)}">`
          : `<w:hyperlink w:anchor="${escapeXml(target.anchor ?? "")}">`;
      openTarget = target;
    }
    xml += runXml(run, ctx.noteIds, ctx);
  }
  closeHyperlink();
  return xml;
}

function paragraphXml(
  paragraph: Paragraph,
  includeOpaque: boolean,
  ctx: InlineWriteContext,
  numRef?: NumberingRef,
  sectPr?: string,
): string {
  return `<w:p>${paragraphPropertiesXml(paragraph, numRef, sectPr, ctx)}${inlineContentXml(paragraph, includeOpaque, ctx)}</w:p>`;
}

/** An Image block (SPEC-024): OOXML has no body-level picture element, so
 * it is wrapped exactly the way Word itself requires — one paragraph
 * holding one run holding the `w:drawing` (see drawing.ts's drawingXml).
 * Returns "" when `mediaPlan` dropped this image's media (an SVG with no
 * .docx representation, see media.ts) — the caller omits the block
 * entirely, matching the drop record opaque.ts's collectDropRecords already
 * reports for it. */
function imageParagraphXml(image: Image, mediaPlan: MediaWritePlan, docPrId: number): string {
  const relId = mediaPlan.relIdFor(image.mediaId);
  if (!relId) {
    return "";
  }
  return `<w:p><w:r>${drawingXml(image, relId, docPrId)}</w:r></w:p>`;
}

// Word's default single-line borders for a table created in this editor —
// self-contained (no style reference), so the table is visible in any
// consumer without touching styles.xml. Sized 4 eighth-points, Word's own
// default for an inserted table. An explicit per-cell w:tcBorders (SPEC-023)
// overrides this default for whichever edges it sets, the normal OOXML
// cell-over-table precedence.
const GENERATED_TABLE_BORDERS =
  "<w:tblBorders>" +
  '<w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>' +
  '<w:start w:val="single" w:sz="4" w:space="0" w:color="auto"/>' +
  '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>' +
  '<w:end w:val="single" w:sz="4" w:space="0" w:color="auto"/>' +
  '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>' +
  '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>' +
  "</w:tblBorders>";

/** A grid slot's owning real cell — the writer's own mirror of lib/tables.ts's
 * same concept (SPEC-023), needed to know, for a merged table, which grid
 * columns/rows a `w:tc` must skip (already covered by an earlier column's
 * `w:gridSpan` in the same row) versus re-emit as a `w:vMerge` continuation
 * (a later row within the origin's `rowSpan`). */
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

function borderStyleVal(style: BorderLineStyle): string {
  if (style === "none") {
    return "nil";
  }
  return style;
}

function borderSideXml(tag: string, side: BorderSide): string {
  const sz = side.style === "none" ? 0 : Math.max(1, ptToEighthPoints(side.widthPt));
  const color = side.color !== "auto" ? (cssColorToOoxml(side.color) ?? "auto") : "auto";
  return `<w:${tag} w:val="${borderStyleVal(side.style)}" w:sz="${sz}" w:space="0" w:color="${color}"/>`;
}

/** A cell's `w:tcBorders` (SPEC-023 AC-10..12, AC-25): only the edges the
 * cell explicitly sets — an unset edge falls back to the table's own
 * `w:tblBorders` default, the normal OOXML precedence. */
function cellBordersXml(borders: CellBorders | undefined): string {
  if (!borders) {
    return "";
  }
  const parts: string[] = [];
  if (borders.top) {
    parts.push(borderSideXml("top", borders.top));
  }
  if (borders.bottom) {
    parts.push(borderSideXml("bottom", borders.bottom));
  }
  if (borders.left) {
    parts.push(borderSideXml("start", borders.left));
  }
  if (borders.right) {
    parts.push(borderSideXml("end", borders.right));
  }
  return parts.length > 0 ? `<w:tcBorders>${parts.join("")}</w:tcBorders>` : "";
}

/** A cell's `w:shd` fill (SPEC-023 AC-15, AC-26). */
function cellShadingXml(fill: string | undefined): string {
  const hex = cssColorToOoxml(fill);
  return hex !== undefined ? `<w:shd w:val="clear" w:color="auto" w:fill="${hex}"/>` : "";
}

/** A cell's `w:tcMar` (SPEC-023 AC-16): the model holds one uniform padding,
 * written to all four sides. */
function cellPaddingXml(paddingCm: number | undefined): string {
  if (paddingCm === undefined) {
    return "";
  }
  const w = cmToTwips(paddingCm);
  return (
    "<w:tcMar>" +
    `<w:top w:w="${w}" w:type="dxa"/>` +
    `<w:start w:w="${w}" w:type="dxa"/>` +
    `<w:bottom w:w="${w}" w:type="dxa"/>` +
    `<w:end w:w="${w}" w:type="dxa"/>` +
    "</w:tcMar>"
  );
}

function cellVAlignXml(align: CellVerticalAlign | undefined): string {
  return align !== undefined ? `<w:vAlign w:val="${align}"/>` : "";
}

/** A cell's tracked insert/delete marks (SPEC-039 AC-15): `w:cellIns`/
 * `w:cellDel` inside `w:tcPr` — the cell-level analogue of a row's
 * `w:trPr/w:ins`|`w:del`. A cell can carry only one of the two (a column
 * insert and a column delete are different edit events). */
function cellInsertDeleteXml(cell: TableCell, ctx: InlineWriteContext): string {
  const insert = cell.changes?.insert;
  const del = cell.changes?.delete;
  const parts: string[] = [];
  if (insert) {
    parts.push(`<w:cellIns${trackedChangeAttrs(changeNumericId(ctx, insert.changeId), insert.author, insert.date)}/>`);
  }
  if (del) {
    parts.push(`<w:cellDel${trackedChangeAttrs(changeNumericId(ctx, del.changeId), del.author, del.date)}/>`);
  }
  return parts.join("");
}

/** A cell's pending property change (SPEC-039 AC-15: a merge/split, or a
 * border/shading/padding/vertical-alignment edit) as `w:tcPrChange` inside
 * `w:tcPr`: the previous span/borders/shading/padding/vertical-alignment,
 * nested in their own `w:tcPr` (the same shape `w:rPrChange`/`w:pPrChange`
 * already use for a run/paragraph's previous formatting). The previous
 * rowSpan has no standard `w:tcPr` slot (rowSpan is normally derived from
 * `w:vMerge` chains spanning several rows, not a per-cell property), so it
 * rides a non-standard `w:prevRowSpan` attribute on `w:tcPrChange` itself. */
function cellPropertyChangeXml(cell: TableCell, ctx: InlineWriteContext): string {
  const change = cell.changes?.propertyChange;
  if (!change) {
    return "";
  }
  const id = changeNumericId(ctx, change.changeId);
  const prev = change.previous;
  const prevRowSpanAttr = prev.rowSpan !== undefined ? ` w:prevRowSpan="${prev.rowSpan}"` : "";
  const parts: string[] = [];
  if (prev.colSpan !== undefined && prev.colSpan > 1) {
    parts.push(`<w:gridSpan w:val="${prev.colSpan}"/>`);
  }
  const borders = cellBordersXml(prev.borders);
  if (borders) {
    parts.push(borders);
  }
  const shading = cellShadingXml(prev.shadingFill);
  if (shading) {
    parts.push(shading);
  }
  const margin = cellPaddingXml(prev.paddingCm);
  if (margin) {
    parts.push(margin);
  }
  const valign = cellVAlignXml(prev.verticalAlign);
  if (valign) {
    parts.push(valign);
  }
  return (
    `<w:tcPrChange${trackedChangeAttrs(id, change.author, change.date)}${prevRowSpanAttr}>` +
    `<w:tcPr>${parts.join("")}</w:tcPr></w:tcPrChange>`
  );
}

/** A freshly generated `w:tcPr` (no preserved source XML to reuse): width,
 * then — in the schema's CT_TcPrBase sequence — gridSpan, vMerge,
 * tcBorders, shd, tcMar, vAlign, then any tracked-change marks (SPEC-039
 * AC-15) last, mirroring `w:rPr`'s own `w:ins`/`w:del`/`w:rPrChange`
 * ordering. */
function generatedTcPrXml(
  cell: TableCell,
  widthTwips: number,
  colSpan: number,
  vMergeRestart: boolean,
  ctx?: InlineWriteContext,
): string {
  const parts = [`<w:tcW w:w="${widthTwips}" w:type="dxa"/>`];
  if (colSpan > 1) {
    parts.push(`<w:gridSpan w:val="${colSpan}"/>`);
  }
  if (vMergeRestart) {
    parts.push('<w:vMerge w:val="restart"/>');
  }
  const borders = cellBordersXml(cell.borders);
  if (borders) {
    parts.push(borders);
  }
  const shading = cellShadingXml(cell.shadingFill);
  if (shading) {
    parts.push(shading);
  }
  const margin = cellPaddingXml(cell.paddingCm);
  if (margin) {
    parts.push(margin);
  }
  const valign = cellVAlignXml(cell.verticalAlign);
  if (valign) {
    parts.push(valign);
  }
  if (ctx) {
    parts.push(cellInsertDeleteXml(cell, ctx));
    parts.push(cellPropertyChangeXml(cell, ctx));
  }
  return `<w:tcPr>${parts.join("")}</w:tcPr>`;
}

/** A row's `w:trPr` (SPEC-023 AC-4, AC-5): the source's own, preserved
 * verbatim, or — for a row this editor set a minimum height on —
 * `w:trHeight` with `hRule="atLeast"`, OOXML's own "minimum, grows for
 * content" row height rule. A tracked row insert/delete (SPEC-039 AC-15)
 * always wins over a preserved verbatim trPr — reusing stale XML here
 * would silently drop the tracked change, which SPEC-008's own "never
 * discard what wasn't understood" rule (and this spec's own "never
 * silently drop a tracked change") both forbid. */
/** A row's pending property change (SPEC-039 AC-15: a height edit) as
 * `w:trPrChange` inside `w:trPr`: the previous height, nested in its own
 * `w:trPr` — same shape as `w:tcPrChange`. */
function rowPropertyChangeXml(row: TableRow, ctx: InlineWriteContext): string {
  const change = row.changes?.propertyChange;
  if (!change) {
    return "";
  }
  const id = changeNumericId(ctx, change.changeId);
  const inner = change.previous.minHeightCm !== undefined
    ? `<w:trHeight w:val="${cmToTwips(change.previous.minHeightCm)}" w:hRule="atLeast"/>`
    : "";
  return `<w:trPrChange${trackedChangeAttrs(id, change.author, change.date)}><w:trPr>${inner}</w:trPr></w:trPrChange>`;
}

function rowPropertiesXml(row: TableRow, includeOpaque: boolean, ctx?: InlineWriteContext): string {
  const insert = row.changes?.insert;
  const del = row.changes?.delete;
  const propertyChange = row.changes?.propertyChange;
  if ((insert || del || propertyChange) && ctx) {
    // A tracked change always regenerates trPr from modelled properties
    // only — reusing preserved verbatim XML here risks producing an
    // invalid element if its source namespace prefix isn't "w:"; the only
    // loss is unmodelled row properties on a row that is, by definition,
    // pending insertion, deletion or a property edit.
    const inner =
      (row.isHeader ? "<w:tblHeader/>" : "") +
      (row.minHeightCm !== undefined
        ? `<w:trHeight w:val="${cmToTwips(row.minHeightCm)}" w:hRule="atLeast"/>`
        : "");
    const marks: string[] = [];
    if (insert) {
      marks.push(`<w:ins${trackedChangeAttrs(changeNumericId(ctx, insert.changeId), insert.author, insert.date)}/>`);
    }
    if (del) {
      marks.push(`<w:del${trackedChangeAttrs(changeNumericId(ctx, del.changeId), del.author, del.date)}/>`);
    }
    if (propertyChange) {
      marks.push(rowPropertyChangeXml(row, ctx));
    }
    return `<w:trPr>${inner}${marks.join("")}</w:trPr>`;
  }
  if (includeOpaque && row.docxTrPr !== undefined) {
    return row.docxTrPr;
  }
  if (row.minHeightCm !== undefined || row.isHeader) {
    const height =
      row.minHeightCm !== undefined
        ? `<w:trHeight w:val="${cmToTwips(row.minHeightCm)}" w:hRule="atLeast"/>`
        : "";
    return `<w:trPr>${row.isHeader ? "<w:tblHeader/>" : ""}${height}</w:trPr>`;
  }
  return "";
}

/** A table's `w:tblW` (SPEC-023 AC-23, AC-27): fixed (the column widths'
 * own sum), a percentage of the text area (fiftieths of a percent is
 * OOXML's unit — 5000 = 100%), or auto (content-sized). */
function tableWidthXml(table: Table, totalTwips: number): string {
  if (table.widthMode === "auto") {
    return '<w:tblW w:w="0" w:type="auto"/>';
  }
  if (table.widthMode === "percent") {
    const pct = Math.round((table.widthPercent ?? 100) * 50);
    return `<w:tblW w:w="${pct}" w:type="pct"/>`;
  }
  return `<w:tblW w:w="${totalTwips}" w:type="dxa"/>`;
}

/** A table's `w:jc` (SPEC-023 AC-22): absent means left, OOXML's own
 * default, so only center/right need an explicit element. */
function tableAlignmentXml(table: Table): string {
  return table.alignment === "center" || table.alignment === "right"
    ? `<w:jc w:val="${table.alignment}"/>`
    : "";
}

/** One table block (SPEC-022; formatting and merged cells SPEC-023):
 * w:tblPr (the source's own, preserved verbatim, or a generated one from
 * the modelled width/alignment plus the default borders), the w:tblGrid
 * column widths, then every row's cells.
 *
 * A merged cell's grid rectangle is walked via computeOwnerGrid: a column
 * covered by an earlier column's `w:gridSpan` in the same row emits nothing
 * (already part of that `w:tc`); the origin row emits the real `w:tc` with
 * its content and (if merged) `w:gridSpan`/`w:vMerge="restart"`; any further
 * row within a vertical merge's `rowSpan` emits an empty continuation
 * `w:tc` marked `w:vMerge` (no `w:val` — "continue" is the schema default),
 * repeating the same `w:gridSpan` so every row's cells still sum to the
 * grid's column count. */
/** A table's pending property change (SPEC-039 AC-15: alignment, width
 * mode/percent, or column widths) as `w:tblPrChange` inside `w:tblPr`: the
 * previous alignment/width, nested in their own `w:tblPr`, plus — when the
 * column widths themselves changed — a nested previous `w:tblGrid` (OOXML
 * gives grid changes their own `w:tblGridChange` sibling element; nesting
 * it here instead keeps one edit event as one element with one id, the
 * same "one changeId per edit" rule every other tracked-change kind
 * follows). */
function tablePropertyChangeXml(table: Table, ctx: InlineWriteContext): string {
  const change = table.changes?.propertyChange;
  if (!change) {
    return "";
  }
  const id = changeNumericId(ctx, change.changeId);
  const prev = change.previous;
  const previousTable: Table = {
    ...table,
    alignment: prev.alignment,
    widthMode: prev.widthMode,
    widthPercent: prev.widthPercent,
    columnWidthsCm: prev.columnWidthsCm ?? table.columnWidthsCm,
  };
  const previousTotalTwips = previousTable.columnWidthsCm.reduce((sum, cm) => sum + cmToTwips(cm), 0);
  const innerTblPr = tableWidthXml(previousTable, previousTotalTwips) + tableAlignmentXml(previousTable);
  const grid =
    prev.columnWidthsCm !== undefined
      ? "<w:tblGrid>" + prev.columnWidthsCm.map((cm) => `<w:gridCol w:w="${cmToTwips(cm)}"/>`).join("") + "</w:tblGrid>"
      : "";
  return (
    `<w:tblPrChange${trackedChangeAttrs(id, change.author, change.date)}>` +
    `<w:tblPr>${innerTblPr}</w:tblPr>${grid}</w:tblPrChange>`
  );
}

function tableXml(
  table: Table,
  includeOpaque: boolean,
  ctx: InlineWriteContext,
  cellNumRef: (row: number, col: number, cellBlock: number) => NumberingRef | undefined,
): string {
  const totalTwips = table.columnWidthsCm.reduce((sum, cm) => sum + cmToTwips(cm), 0);
  const hasTrackedTableChange = !!table.changes?.propertyChange;
  const tblPr =
    includeOpaque && table.docxTblPr !== undefined && !hasTrackedTableChange
      ? table.docxTblPr
      : "<w:tblPr>" +
        tableWidthXml(table, totalTwips) +
        tableAlignmentXml(table) +
        GENERATED_TABLE_BORDERS +
        '<w:tblLayout w:type="fixed"/>' +
        tablePropertyChangeXml(table, ctx) +
        "</w:tblPr>";
  const grid =
    "<w:tblGrid>" +
    table.columnWidthsCm.map((cm) => `<w:gridCol w:w="${cmToTwips(cm)}"/>`).join("") +
    "</w:tblGrid>";
  const ownerGrid = computeOwnerGrid(table);
  const rows = table.rows
    .map((row, rowIndex) => {
      const trPr = rowPropertiesXml(row, includeOpaque, ctx);
      const cells: string[] = [];
      for (let colIndex = 0; colIndex < table.columnWidthsCm.length; colIndex++) {
        const owner = ownerGrid[rowIndex][colIndex];
        if (owner.originCol !== colIndex) {
          // Covered by an earlier column's gridSpan in this row — already
          // part of that w:tc.
          continue;
        }
        const widthTwips = table.columnWidthsCm
          .slice(colIndex, colIndex + owner.colSpan)
          .reduce((sum, cm) => sum + cmToTwips(cm), 0);
        if (owner.originRow === rowIndex) {
          const cell = row.cells[colIndex];
          // A tracked cell change always regenerates tcPr from modelled
          // properties (SPEC-039 AC-15) — same "never reuse stale preserved
          // XML over a pending change" rule rowPropertiesXml already
          // follows.
          const hasTrackedChange = !!(cell.changes?.insert || cell.changes?.delete || cell.changes?.propertyChange);
          const tcPr =
            includeOpaque && cell.docxTcPr !== undefined && !hasTrackedChange
              ? cell.docxTcPr
              : generatedTcPrXml(cell, widthTwips, owner.colSpan, owner.rowSpan > 1, ctx);
          const content =
            cell.blocks.length > 0
              ? cell.blocks
                  .map((paragraph, blockIndex) =>
                    paragraphXml(paragraph, includeOpaque, ctx, cellNumRef(rowIndex, colIndex, blockIndex)),
                  )
                  .join("")
              : "<w:p/>";
          cells.push(`<w:tc>${tcPr}${content}</w:tc>`);
        } else {
          const tcPr =
            `<w:tcPr><w:tcW w:w="${widthTwips}" w:type="dxa"/>` +
            (owner.colSpan > 1 ? `<w:gridSpan w:val="${owner.colSpan}"/>` : "") +
            "<w:vMerge/></w:tcPr>";
          cells.push(`<w:tc>${tcPr}<w:p/></w:tc>`);
        }
      }
      return `<w:tr>${trPr}${cells.join("")}</w:tr>`;
    })
    .join("");
  return `<w:tbl>${tblPr}${grid}${rows}</w:tbl>`;
}

// --- headers, footers and page numbers (SPEC-029) ---------------------------

export const HEADER_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml";
export const FOOTER_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml";
const HEADER_RELATIONSHIP = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header";
const FOOTER_RELATIONSHIP = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer";
const SETTINGS_RELATIONSHIP =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings";
const SETTINGS_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml";

interface HeaderFooterPlanEntry {
  kind: HeaderFooterKind;
  relId: string;
}

interface SectionHeaderFooterPlan {
  headers: HeaderFooterPlanEntry[];
  footers: HeaderFooterPlanEntry[];
}

interface HeaderFooterPartPlan {
  partName: string;
  blocks: Block[];
  mediaPlan: MediaWritePlan;
  hyperlinkPlan: HyperlinkWritePlan;
}

export interface HeaderFooterWritePlan {
  /** Per section (by index), the header/footer references its own sectPr
   * needs — empty for a section with no header/footer content. */
  bySection: SectionHeaderFooterPlan[];
  /** Every header1.xml/footer1.xml-shaped part this plan mints, with the
   * content and (part-scoped, own-rels-file) media/hyperlink plans it needs
   * to serialise. */
  parts: HeaderFooterPartPlan[];
  /** `<Relationship>` elements for word/_rels/document.xml.rels. */
  relationshipsXml: string;
  /** `<Override>` elements for `[Content_Types].xml`. */
  contentTypesXml: string;
  /** True when any section wants different odd/even pages — OOXML only
   * has one document-wide flag for this (`w:evenAndOddHeaders` in
   * word/settings.xml), so every section's own toggle is OR'd together. */
  evenAndOddHeaders: boolean;
}

/** The header/footer kinds a section actually needs a part for — a kind
 * with no content (an unedited "first"/"even" the section's toggle merely
 * enabled) needs none, matching resolveHeaderFooterKind's "genuinely
 * blank" contract on the read side. */
function neededHeaderFooterKinds(section: Section, part: HeaderFooterPart): HeaderFooterKind[] {
  const map = part === "header" ? section.headers : section.footers;
  if (!map) {
    return [];
  }
  return (Object.keys(map) as HeaderFooterKind[]).filter((kind) => map[kind] !== undefined);
}

/** Plans every header1.xml/footer1.xml-shaped part a save needs (SPEC-029):
 * headers/footers are always regenerated fresh from the model — never
 * preserved-then-patched the way styles.xml/numbering.xml are — since their
 * content is now fully modelled, so this mints a brand new part name/
 * relationship id for every one regardless of same-format vs. from-scratch.
 * `existingMediaPartNames` is shared with the caller's own body-image
 * MediaWritePlan (mutated in place by planMediaWriteForBlocks) so a
 * header's images can never collide with the document's own
 * `word/media/imageN.ext` names. */
function planHeaderFooterWrite(
  model: DocumentModel,
  takenRelIds: Set<string>,
  existingMediaPartNames: Set<string>,
  existingPartNames: Set<string> = new Set(),
): HeaderFooterWritePlan {
  const bySection: SectionHeaderFooterPlan[] = [];
  const parts: HeaderFooterPartPlan[] = [];
  const relParts: string[] = [];
  const overrideParts: string[] = [];
  const taken = new Set(takenRelIds);
  let headerCount = 0;
  let footerCount = 0;
  let evenAndOddHeaders = false;

  function planOne(
    part: HeaderFooterPart,
    kind: HeaderFooterKind,
    blocks: Block[],
    count: number,
  ): HeaderFooterPlanEntry {
    let n = count;
    let fileName = `${part}${n}.xml`;
    let partName = `word/${fileName}`;
    // A source package can carry an orphaned word/headerN.xml the reader
    // never consumed (unreferenced by any current sectPr) — never overwrite
    // it, the same collision-avoidance planMediaWrite gives images.
    while (existingPartNames.has(partName)) {
      n++;
      fileName = `${part}${n}.xml`;
      partName = `word/${fileName}`;
    }
    const relId = nextFreeRelId(taken);
    taken.add(relId);
    const mediaPlan = planMediaWriteForBlocks(blocks, model.media, existingMediaPartNames, 1);
    const hyperlinkPlan = planHyperlinkWriteFor(blocks, new Set());
    parts.push({ partName, blocks, mediaPlan, hyperlinkPlan });
    const relType = part === "header" ? HEADER_RELATIONSHIP : FOOTER_RELATIONSHIP;
    relParts.push(`<Relationship Id="${relId}" Type="${relType}" Target="${fileName}"/>`);
    const contentType = part === "header" ? HEADER_CONTENT_TYPE : FOOTER_CONTENT_TYPE;
    overrideParts.push(`<Override PartName="/${partName}" ContentType="${contentType}"/>`);
    return { kind, relId };
  }

  for (const section of model.sections) {
    if (section.evenAndOddHeaders) {
      evenAndOddHeaders = true;
    }
    const headers: HeaderFooterPlanEntry[] = [];
    for (const kind of neededHeaderFooterKinds(section, "header")) {
      headerCount += 1;
      headers.push(planOne("header", kind, section.headers![kind]!.blocks, headerCount));
    }
    const footers: HeaderFooterPlanEntry[] = [];
    for (const kind of neededHeaderFooterKinds(section, "footer")) {
      footerCount += 1;
      footers.push(planOne("footer", kind, section.footers![kind]!.blocks, footerCount));
    }
    bySection.push({ headers, footers });
  }

  return {
    bySection,
    parts,
    relationshipsXml: relParts.join(""),
    contentTypesXml: overrideParts.join(""),
    evenAndOddHeaders,
  };
}

// --- table of contents (SPEC-033) -------------------------------------------

// Word's own TOC style indent step (0.25in per level) — cosmetic only; the
// reader never looks at paragraph indentation to parse a TOC back, only the
// field's own structure (see reader.ts's parseTocField).
const TOC_LEVEL_INDENT_TWIPS = 360;

function tocFieldInstr(options: TocOptions): string {
  const levels = options.levels.length > 0 ? options.levels : [1];
  const min = Math.min(...levels);
  const max = Math.max(...levels);
  return ` TOC \\o "${min}-${max}" \\h \\z \\u `;
}

function tocEntryPPr(entry: TocEntry, tabPosTwips: number, includeTab: boolean): string {
  const indent = (Math.max(1, entry.level) - 1) * TOC_LEVEL_INDENT_TWIPS;
  const tabs = includeTab
    ? `<w:tabs><w:tab w:val="right" w:leader="dot" w:pos="${tabPosTwips}"/></w:tabs>`
    : "";
  return `<w:pPr><w:pStyle w:val="TOC${entry.level}"/>${tabs}<w:ind w:left="${indent}"/></w:pPr>`;
}

/** One TOC entry's inline content: its heading text as an internal
 * hyperlink to the heading's anchor (AC-8, AC-9 — every entry is clickable
 * regardless of options), plus, unless `hyperlinksOnly`, a dot-leader tab
 * and the cached page number (AC-2..4, AC-7). The page number is written as
 * plain cached text rather than a nested PAGEREF field — it is already
 * inside the outer TOC field's own cached result, so Word displays it
 * without prompting either way, and a plain run keeps the entry's XML (and
 * the reader that parses it back) far simpler. */
function tocEntryContentXml(entry: TocEntry, options: TocOptions): string {
  const textRun = `<w:r><w:t>${escapeXml(entry.text)}</w:t></w:r>`;
  const trailer =
    options.showPageNumbers && !options.hyperlinksOnly
      ? `<w:r><w:tab/></w:r><w:r><w:t>${escapeXml(String(entry.pageNumber ?? ""))}</w:t></w:r>`
      : "";
  return `<w:hyperlink w:anchor="${escapeXml(entry.anchor)}">${textRun}${trailer}</w:hyperlink>`;
}

/** The `w:p` sequence a TOC block writes (SPEC-033 AC-17): a real complex
 * field — `w:fldChar` begin/separate/end around a `w:instrText` TOC
 * instruction — whose cached RESULT is the entries themselves, so Word
 * shows them immediately on open with no "update fields?" prompt. Every
 * entry is its own paragraph in the TOCn style Word itself uses; the field
 * boundary markers live in the first and last paragraphs, exactly where
 * Word puts them when it saves a document containing a TOC. */
function tocFieldXml(toc: TableOfContents, model: DocumentModel, sectionIndex: number): string {
  const tabPosTwips = cmToTwips(sectionContentWidthCm(model, sectionIndex));
  const beginRuns =
    '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
    `<w:r><w:instrText xml:space="preserve">${escapeXml(tocFieldInstr(toc.options))}</w:instrText></w:r>` +
    '<w:r><w:fldChar w:fldCharType="separate"/></w:r>';
  const endRun = '<w:r><w:fldChar w:fldCharType="end"/></w:r>';
  if (toc.entries.length === 0) {
    return (
      `<w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr>${beginRuns}` +
      `<w:r><w:t>${escapeXml(TOC_PLACEHOLDER_TEXT)}</w:t></w:r>${endRun}</w:p>`
    );
  }
  const includeTab = toc.options.showPageNumbers && !toc.options.hyperlinksOnly;
  return toc.entries
    .map((entry, index) => {
      const pPr = tocEntryPPr(entry, tabPosTwips, includeTab);
      const boundary = index === 0 ? beginRuns : "";
      const trailingEnd = index === toc.entries.length - 1 ? endRun : "";
      return `<w:p>${pPr}${boundary}${tocEntryContentXml(entry, toc.options)}${trailingEnd}</w:p>`;
    })
    .join("");
}

/** One header/footer part's full XML (SPEC-029 AC-3..6, AC-24): its own
 * paragraphs/tables/images, sharing the exact same per-block serialisation
 * body content uses (paragraphXml/tableXml/imageParagraphXml) rather than a
 * second copy of it — resolved against the part's OWN hyperlink/media
 * plans (its own independent rels file, never the main document's). */
function headerFooterPartXml(
  rootTag: "hdr" | "ftr",
  includeOpaque: boolean,
  plan: HeaderFooterPartPlan,
): string {
  const inlineCtx: InlineWriteContext = { hyperlinkPlan: plan.hyperlinkPlan, bookmarkIds: new Map(), changeIds: new Map() };
  let docPrId = 1;
  const body = plan.blocks
    .map((block) => {
      if (block.type === "table") {
        return tableXml(block, includeOpaque, inlineCtx, () => undefined);
      }
      if (block.type === "image") {
        return imageParagraphXml(block, plan.mediaPlan, docPrId++);
      }
      if (block.type === "break") {
        return "";
      }
      if (block.type === "toc") {
        // A TOC is a body-level feature (SPEC-033); one is never inserted
        // into a header/footer, but the Block union covers it structurally.
        return "";
      }
      return paragraphXml(block, includeOpaque, inlineCtx);
    })
    .join("");
  return (
    `${XML_DECLARATION}<w:${rootTag} xmlns:w="${WORD_NAMESPACE}" xmlns:r="${R_ATTRIBUTE_NAMESPACE}">` +
    (body.length > 0 ? body : "<w:p/>") +
    `</w:${rootTag}>`
  );
}

/** A header/footer part's own `_rels/*.rels` — null when it references no
 * media/hyperlinks (most header/footer parts), so the caller adds no part
 * for it at all. */
function headerFooterPartRelsXml(plan: HeaderFooterPartPlan): string | null {
  const relXml = plan.mediaPlan.relationshipsXml + plan.hyperlinkPlan.relationshipsXml;
  if (relXml.length === 0) {
    return null;
  }
  return `${XML_DECLARATION}<Relationships xmlns="${RELATIONSHIPS_NAMESPACE}">${relXml}</Relationships>`;
}

// --- footnotes and endnotes (SPEC-031) ---------------------------------------

export const FOOTNOTES_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml";
export const ENDNOTES_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml";
const FOOTNOTES_RELATIONSHIP =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes";
const ENDNOTES_RELATIONSHIP =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes";

interface NotesPartPlan {
  kind: NoteKind;
  partName: string;
  notes: NoteContent[];
  hyperlinkPlan: HyperlinkWritePlan;
}

export interface NotesWritePlan {
  /** One footnotes.xml and/or endnotes.xml part per stream with content. */
  parts: NotesPartPlan[];
  /** Numeric part-local ids by model noteId — what the body's
   * w:footnoteReference/w:endnoteReference elements carry. */
  noteIds: Map<string, number>;
  /** `<Relationship>` elements for word/_rels/document.xml.rels. */
  relationshipsXml: string;
  /** `<Override>` elements for `[Content_Types].xml`. */
  contentTypesXml: string;
  /** Relationship ids this plan claimed (for later planners' taken sets). */
  relIds: Set<string>;
}

/** Plans the footnotes.xml/endnotes.xml parts a save needs (SPEC-031
 * AC-21/22): like headers/footers, the note parts are always regenerated
 * fresh from the model. Only referenced notes are written (an unreferenced
 * note is invisible everywhere — see NoteContent), numbered 1..n in
 * first-reference order; ids -1/0 are the conventional separator entries.
 * Notes' own hyperlinks resolve through the part's own rels file. */
export function planNotesWrite(
  model: DocumentModel,
  takenRelIds: Set<string> = new Set(),
  existingPartNames: Set<string> = new Set(),
): NotesWritePlan {
  const parts: NotesPartPlan[] = [];
  const relParts: string[] = [];
  const overrideParts: string[] = [];
  const noteIds = new Map<string, number>();
  const taken = new Set(takenRelIds);
  const relIds = new Set<string>();
  for (const kind of ["footnote", "endnote"] as NoteKind[]) {
    const notes = referencedNotes(model, kind);
    if (notes.length === 0) {
      continue;
    }
    const base = kind === "footnote" ? "footnotes" : "endnotes";
    let n = 1;
    let fileName = `${base}.xml`;
    let partName = `word/${fileName}`;
    while (existingPartNames.has(partName)) {
      n += 1;
      fileName = `${base}${n}.xml`;
      partName = `word/${fileName}`;
    }
    const relId = nextFreeRelId(taken);
    taken.add(relId);
    relIds.add(relId);
    notes.forEach((note, index) => noteIds.set(note.noteId, index + 1));
    const hyperlinkPlan = planHyperlinkWriteFor(
      notes.flatMap((note) => note.blocks),
      new Set(),
    );
    parts.push({ kind, partName, notes, hyperlinkPlan });
    const relType = kind === "footnote" ? FOOTNOTES_RELATIONSHIP : ENDNOTES_RELATIONSHIP;
    relParts.push(`<Relationship Id="${relId}" Type="${relType}" Target="${fileName}"/>`);
    const contentType = kind === "footnote" ? FOOTNOTES_CONTENT_TYPE : ENDNOTES_CONTENT_TYPE;
    overrideParts.push(`<Override PartName="/${partName}" ContentType="${contentType}"/>`);
  }
  return {
    parts,
    noteIds,
    relationshipsXml: relParts.join(""),
    contentTypesXml: overrideParts.join(""),
    relIds,
  };
}

/** One footnotes.xml/endnotes.xml part's full XML (SPEC-031 AC-21/22): the
 * conventional separator/continuationSeparator entries at ids -1/0, then
 * every referenced note at its planned id, its content sharing the body's
 * per-block serialisation. The note's first paragraph opens with a fresh
 * superscript `w:footnoteRef`/`w:endnoteRef` marker run — Word's own shape,
 * and the element the reader drops on open, so the round trip is stable. */
function notesPartXml(plan: NotesPartPlan, includeOpaque: boolean, noteIds: Map<string, number>): string {
  const tag = plan.kind === "footnote" ? "footnote" : "endnote";
  const rootTag = plan.kind === "footnote" ? "footnotes" : "endnotes";
  const refTag = plan.kind === "footnote" ? "footnoteRef" : "endnoteRef";
  const inlineCtx: InlineWriteContext = { hyperlinkPlan: plan.hyperlinkPlan, bookmarkIds: new Map(), changeIds: new Map() };
  const markerRun = `<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:${refTag}/></w:r>`;
  const noteXml = plan.notes
    .map((note) => {
      const id = noteIds.get(note.noteId) ?? 0;
      let markerPending = true;
      const body = note.blocks
        .map((block) => {
          if (block.type === "table") {
            return tableXml(block, includeOpaque, inlineCtx, () => undefined);
          }
          if (block.type === "image" || block.type === "break" || block.type === "toc") {
            // A note body holds paragraphs (and its first block must, to
            // carry the reference marker); images/breaks/TOCs have no valid
            // per-note serialisation, so they are omitted here the way they
            // are omitted from the note areas.
            return "";
          }
          if (markerPending) {
            markerPending = false;
            return `<w:p>${paragraphPropertiesXml(block)}${markerRun}${inlineContentXml(block, includeOpaque, inlineCtx)}</w:p>`;
          }
          return paragraphXml(block, includeOpaque, inlineCtx);
        })
        .join("");
      return `<w:${tag} w:id="${id}">${body.length > 0 ? body : `<w:p>${markerRun}</w:p>`}</w:${tag}>`;
    })
    .join("");
  return (
    `${XML_DECLARATION}<w:${rootTag} xmlns:w="${WORD_NAMESPACE}" xmlns:r="${R_ATTRIBUTE_NAMESPACE}">` +
    `<w:${tag} w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:${tag}>` +
    `<w:${tag} w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:${tag}>` +
    noteXml +
    `</w:${rootTag}>`
  );
}

/** A note part's own `_rels/*.rels` — null when its notes reference no
 * external hyperlinks (the common case). */
function notesPartRelsXml(plan: NotesPartPlan): string | null {
  const relXml = plan.hyperlinkPlan.relationshipsXml;
  if (relXml.length === 0) {
    return null;
  }
  return `${XML_DECLARATION}<Relationships xmlns="${RELATIONSHIPS_NAMESPACE}">${relXml}</Relationships>`;
}

// --- comments (SPEC-038) -----------------------------------------------------

export const COMMENTS_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml";
const COMMENTS_RELATIONSHIP =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";

interface CommentsPartPlan {
  partName: string;
  threads: VisibleCommentThread[];
}

export interface CommentsWritePlan {
  /** Zero or one word/comments.xml part — a plan-shaped array (rather than
   * `| null`) so writeDocx's per-notes-part loops generalise to comments
   * without a separate null check at each call site. */
  parts: CommentsPartPlan[];
  /** Numeric part-local ids by model threadId — what the body's
   * w:commentRangeStart/End/commentReference elements carry. */
  commentIds: Map<string, number>;
  relationshipsXml: string;
  contentTypesXml: string;
  relIds: Set<string>;
}

/** Plans the comments.xml part a save needs (SPEC-038): like the note
 * parts, always regenerated fresh from the model. Only threads whose
 * anchor range is still present in the body are written (see
 * lib/comments.ts's visibleCommentThreads — the same "invisible unless
 * anchored" contract NoteContent uses), numbered 1..n in document order;
 * that numeric id is shared by the range markers and the `<w:comment>`
 * element, so no separate reference-id lookup is needed to match them up. */
export function planCommentsWrite(
  model: DocumentModel,
  takenRelIds: Set<string> = new Set(),
  existingPartNames: Set<string> = new Set(),
): CommentsWritePlan {
  const threads = visibleCommentThreads(model);
  if (threads.length === 0) {
    return { parts: [], commentIds: new Map(), relationshipsXml: "", contentTypesXml: "", relIds: new Set() };
  }
  let n = 1;
  let fileName = "comments.xml";
  let partName = `word/${fileName}`;
  while (existingPartNames.has(partName)) {
    n += 1;
    fileName = `comments${n}.xml`;
    partName = `word/${fileName}`;
  }
  const taken = new Set(takenRelIds);
  const relId = nextFreeRelId(taken);
  const commentIds = new Map<string, number>();
  threads.forEach((entry, index) => commentIds.set(entry.thread.threadId, index + 1));
  return {
    parts: [{ partName, threads }],
    commentIds,
    relationshipsXml: `<Relationship Id="${relId}" Type="${COMMENTS_RELATIONSHIP}" Target="${fileName}"/>`,
    contentTypesXml: `<Override PartName="/${partName}" ContentType="${COMMENTS_CONTENT_TYPE}"/>`,
    relIds: new Set([relId]),
  };
}

/** Up to 3 initials from an author's name, for `w:comment/@w:initials`
 * (Word shows these in the margin balloon's collapsed state). */
function initialsOf(author: string): string {
  return author
    .trim()
    .split(/\s+/)
    .map((word) => word[0]?.toUpperCase() ?? "")
    .join("")
    .slice(0, 3);
}

/** word/comments.xml's full XML (SPEC-038): one `<w:comment>` per visible
 * thread, its `w:author`/`w:date` from the opening message, its body one
 * `<w:p>` per message — the opening message's text verbatim, then each
 * reply as `replyParagraphText`. */
function commentsPartXml(plan: CommentsPartPlan, commentIds: Map<string, number>): string {
  const body = plan.threads
    .map(({ thread }) => {
      const id = commentIds.get(thread.threadId) ?? 0;
      const opening = thread.messages[0];
      const paragraphs = [opening.text, ...thread.messages.slice(1).map(replyParagraphText)]
        .map((text) => `<w:p>${text.length > 0 ? `<w:r>${runContentXml(text)}</w:r>` : ""}</w:p>`)
        .join("");
      return (
        `<w:comment w:id="${id}" w:author="${escapeXml(opening.author)}" w:date="${escapeXml(opening.createdAt)}" w:initials="${escapeXml(initialsOf(opening.author))}">` +
        paragraphs +
        `</w:comment>`
      );
    })
    .join("");
  return `${XML_DECLARATION}<w:comments xmlns:w="${WORD_NAMESPACE}">${body}</w:comments>`;
}

// The w:numFmt value for each model note numbering format — "chicago" is
// OOXML's name for the * † ‡ § symbol cycle.
const NOTE_NUM_FMT_VALUES: Record<NoteNumberFormat, string> = {
  decimal: "decimal",
  lowerRoman: "lowerRoman",
  lowerAlpha: "lowerLetter",
  symbol: "chicago",
};

/** A sectPr's `w:footnotePr`/`w:endnotePr` (SPEC-031 AC-17..20) from the
 * model's note settings — "" when nothing is set, so untouched documents
 * emit exactly what they always did. Children in CT_FtnProps order:
 * numFmt, numStart, numRestart. */
function notePrXml(tag: "footnotePr" | "endnotePr", settings: NoteSettings | undefined): string {
  if (!settings) {
    return "";
  }
  const parts: string[] = [];
  if (settings.numberFormat !== undefined) {
    parts.push(`<w:numFmt w:val="${NOTE_NUM_FMT_VALUES[settings.numberFormat]}"/>`);
  }
  if (settings.startAt !== undefined) {
    parts.push(`<w:numStart w:val="${settings.startAt}"/>`);
  }
  if (settings.restart !== undefined && settings.restart !== "continuous") {
    parts.push(`<w:numRestart w:val="${settings.restart === "page" ? "eachPage" : "eachSect"}"/>`);
  }
  return parts.length > 0 ? `<w:${tag}>${parts.join("")}</w:${tag}>` : "";
}

/** A from-scratch word/settings.xml (SPEC-029, SPEC-039): created when a
 * section wants different odd/even headers and/or the document's
 * recording-enabled flag (`w:trackChanges`) is set — OOXML's only
 * representation of either is a document-wide flag, so a document that
 * never asked for either never gains a settings part at all. */
function buildSettingsXml(trackChangesRecording: boolean): string {
  const trackChanges = trackChangesRecording ? "<w:trackChanges/>" : "";
  return `${XML_DECLARATION}<w:settings xmlns:w="${WORD_NAMESPACE}"><w:evenAndOddHeaders/>${trackChanges}</w:settings>`;
}

/** Adds or removes `w:evenAndOddHeaders`/`w:trackChanges` on a preserved
 * settings.xml (SPEC-029, SPEC-039) — "patch, don't regenerate", like
 * patchStylesXml, since settings.xml carries many other unmodelled
 * document-wide flags this spec must not disturb. Returns null when both
 * flags already match. */
export function patchSettingsXml(
  evenAndOddHeaders: boolean,
  trackChangesRecording: boolean,
  originalXml: string,
): string | null {
  let xml = originalXml;
  let changed = false;
  const hasEvenOdd = /<[\w:]*evenAndOddHeaders\b/.test(xml);
  if (evenAndOddHeaders !== hasEvenOdd) {
    changed = true;
    xml = evenAndOddHeaders
      ? xml.replace(/(<w:settings[^>]*>)/, "$1<w:evenAndOddHeaders/>")
      : xml.replace(/<[\w:]*evenAndOddHeaders\s*\/>/, "");
  }
  const hasTrackChanges = /<[\w:]*trackChanges\b/.test(xml);
  if (trackChangesRecording !== hasTrackChanges) {
    changed = true;
    xml = trackChangesRecording
      ? xml.replace(/(<w:settings[^>]*>)/, "$1<w:trackChanges/>")
      : xml.replace(/<[\w:]*trackChanges\s*\/>/, "");
  }
  return changed ? xml : null;
}

// The section's `w:cols` (SPEC-030 AC-16/AC-16a): equal columns emit just
// `w:num`/`w:space` (no children); unequal columns emit an explicit `w:col`
// per column carrying its width in twips and — for all but the last — its
// trailing gutter. Widths are resolved against the section's content width so
// they add up exactly.
function columnsXml(columns: SectionColumns, contentWidthCm: number): string {
  const space = cmToTwips(Math.max(0, columns.gapCm));
  const sep = columns.separator ? ' w:sep="1"' : "";
  if (columns.equalWidth || !columns.widthsCm || columns.widthsCm.length !== columns.count) {
    return `<w:cols w:num="${columns.count}" w:space="${space}"${sep}/>`;
  }
  const widths = resolveColumnWidthsCm(columns, contentWidthCm);
  const children = widths
    .map((width, index) => {
      const trailing = index < widths.length - 1 ? ` w:space="${space}"` : "";
      return `<w:col w:w="${cmToTwips(width)}"${trailing}/>`;
    })
    .join("");
  return `<w:cols w:num="${columns.count}" w:space="${space}"${sep}>${children}</w:cols>`;
}

/** A `w:sectPr` generated from modelled section geometry, in CT_SectPr
 * sequence order: header/footer references (SPEC-029), w:type (continuous
 * starts only — nextPage is OOXML's own default, matching the model's
 * absent `start`), the required w:pgSz/w:pgMar, then w:pgNumType (SPEC-029
 * AC-22), w:cols (SPEC-030) and w:titlePg (AC-11). Shared by the body-level
 * sectPr describing the final section and the paragraph-level sectPr every
 * non-final section is written as (SPEC-028 AC-31). */
function sectionPropertiesXml(
  pageSize: PageSize,
  margins: Margins,
  start: SectionStart | undefined,
  headerFooter: SectionHeaderFooterPlan,
  titlePage: boolean,
  pageNumberStart: number | undefined,
  columns?: SectionColumns,
  notesPr = "",
): string {
  const width = cmToTwips(pageSize.widthCm);
  const height = cmToTwips(pageSize.heightCm);
  const orient = width > height ? ' w:orient="landscape"' : "";
  const type = start === "continuous" ? '<w:type w:val="continuous"/>' : "";
  const { topCm, rightCm, bottomCm, leftCm } = margins;
  const cols =
    columns && columns.count > 1
      ? columnsXml(columns, Math.max(0.1, pageSize.widthCm - leftCm - rightCm))
      : "";
  const headerRefs = headerFooter.headers
    .map((ref) => `<w:headerReference w:type="${ref.kind}" r:id="${ref.relId}"/>`)
    .join("");
  const footerRefs = headerFooter.footers
    .map((ref) => `<w:footerReference w:type="${ref.kind}" r:id="${ref.relId}"/>`)
    .join("");
  // xmlns:r is only ever needed when a reference's r:id attribute is
  // present — declared locally here rather than on the document root, the
  // same minimal-declaration convention w:hyperlink already follows.
  const needsR = headerRefs.length > 0 || footerRefs.length > 0;
  const pgNumType = pageNumberStart !== undefined ? `<w:pgNumType w:start="${pageNumberStart}"/>` : "";
  const titlePg = titlePage ? "<w:titlePg/>" : "";
  return (
    `<w:sectPr${needsR ? ` xmlns:r="${R_ATTRIBUTE_NAMESPACE}"` : ""}>` +
    headerRefs +
    footerRefs +
    notesPr +
    type +
    `<w:pgSz w:w="${width}" w:h="${height}"${orient}/>` +
    `<w:pgMar w:top="${cmToTwips(topCm)}" w:right="${cmToTwips(rightCm)}"` +
    ` w:bottom="${cmToTwips(bottomCm)}" w:left="${cmToTwips(leftCm)}"` +
    ` w:header="${HEADER_FOOTER_TWIPS}" w:footer="${HEADER_FOOTER_TWIPS}" w:gutter="0"/>` +
    pgNumType +
    cols +
    titlePg +
    "</w:sectPr>"
  );
}

/** The main document part. Every non-final section ends with a w:sectPr
 * inside its last paragraph's w:pPr (SPEC-028 AC-31) — the section's own
 * preserved XML when the source provided one, else generated from its
 * modelled geometry; a section whose last block is not a paragraph gets an
 * extra empty section-break paragraph to host it. The final section is the
 * body-level w:sectPr, exactly as before SPEC-028 — from its geometry,
 * which for a single-section document is `defaultStyles` unchanged. An
 * empty document still gets one w:p — a body whose only child is w:sectPr
 * triggers Word's repair prompt (AC-24).
 *
 * SPEC-008: when the model came from a .docx, preserved opaque blocks and
 * inlines are re-emitted verbatim at their anchored positions, and the
 * original body-level w:sectPr replaces the generated one so unmodelled
 * section properties survive. Cross-format models carry foreign (ODF) XML
 * in their opaque nodes, so those are excluded — the writer reports them as
 * dropped instead. */
export function buildDocumentXml(
  model: DocumentModel,
  plan: NumberingPlan = planNumbering(model),
  mediaPlan: MediaWritePlan = planMediaWrite(model),
  hyperlinkPlan: HyperlinkWritePlan = planHyperlinkWrite(model),
  headerFooterPlan: HeaderFooterWritePlan = planHeaderFooterWrite(model, new Set(), new Set()),
  notesPlan: NotesWritePlan = planNotesWrite(model),
  commentsPlan: CommentsWritePlan = planCommentsWrite(model),
): string {
  const includeOpaque = model.format === "docx";
  const inlineCtx: InlineWriteContext = {
    hyperlinkPlan,
    bookmarkIds: bookmarkIdsFor(model),
    noteIds: notesPlan.noteIds,
    commentIds: commentsPlan.commentIds,
    changeIds: new Map(),
  };
  const blockXml: string[] = [];
  let docPrId = 1;
  // A section with any header/footer reference, title page or explicit
  // page-number start must never reuse a preserved verbatim sectPr
  // (SPEC-029): the writer mints fresh relationship ids for regenerated
  // header/footer parts every save, so a stale preserved sectPr's r:id
  // values would dangle — sectionHeaderFooterOf's bySection entry (built
  // from the *current* model, not whatever the source file had) is always
  // the source of truth for whether that applies.
  const sectionHeaderFooterOf = (index: number): SectionHeaderFooterPlan =>
    headerFooterPlan.bySection[index] ?? { headers: [], footers: [] };
  const sectionNeedsFreshSectPr = (section: Section, index: number): boolean => {
    const hf = sectionHeaderFooterOf(index);
    return (
      hf.headers.length > 0 ||
      hf.footers.length > 0 ||
      section.titlePage === true ||
      section.pageNumberStart !== undefined
    );
  };
  model.sections.forEach((section, sectionIndex) => {
    const isFinal = sectionIndex === model.sections.length - 1;
    const sectPrXml = isFinal
      ? undefined
      : includeOpaque && section.docxSectPr !== undefined && !sectionNeedsFreshSectPr(section, sectionIndex)
        ? section.docxSectPr
        : sectionPropertiesXml(
            sectionPageSize(model, sectionIndex),
            sectionMargins(model, sectionIndex),
            section.start,
            sectionHeaderFooterOf(sectionIndex),
            section.titlePage === true,
            section.pageNumberStart,
            section.columns,
          );
    const lastBlock = section.blocks[section.blocks.length - 1];
    const embedSectPr = sectPrXml !== undefined && lastBlock?.type === "paragraph";
    const opaqueBlocks = includeOpaque
      ? [...(section.opaqueBlocks ?? [])].sort((a, b) => a.blockIndex - b.blockIndex)
      : [];
    for (let i = 0; i <= section.blocks.length; i++) {
      for (const node of opaqueBlocks) {
        if (Math.min(node.blockIndex, section.blocks.length) === i) {
          blockXml.push(node.xml);
        }
      }
      if (i < section.blocks.length) {
        const block = section.blocks[i];
        if (block.type === "table") {
          blockXml.push(
            tableXml(block, includeOpaque, inlineCtx, (row, col, cellBlock) =>
              cellNumberingRefFor(plan, sectionIndex, i, row, col, cellBlock),
            ),
          );
        } else if (block.type === "image") {
          blockXml.push(imageParagraphXml(block, mediaPlan, docPrId++));
        } else if (block.type === "break") {
          // A manual break block (SPEC-028 AC-30): a paragraph holding just
          // the typed w:br run, the shape the reader folds back into a
          // BreakBlock on open.
          blockXml.push(`<w:p><w:r><w:br w:type="${block.kind}"/></w:r></w:p>`);
        } else if (block.type === "toc") {
          blockXml.push(tocFieldXml(block, model, sectionIndex));
        } else {
          blockXml.push(
            paragraphXml(
              block,
              includeOpaque,
              inlineCtx,
              numberingRefFor(plan, sectionIndex, i),
              embedSectPr && i === section.blocks.length - 1 ? sectPrXml : undefined,
            ),
          );
        }
      }
    }
    if (sectPrXml !== undefined && !embedSectPr) {
      // No paragraph to end the section on (a table/image/break last, or an
      // empty section): the section break needs a paragraph mark to live
      // on, so an empty one is appended — the same shape Word itself uses
      // after a section-ending table.
      blockXml.push(`<w:p><w:pPr>${sectPrXml}</w:pPr></w:p>`);
    }
  });
  const body = blockXml.length > 0 ? blockXml.join("") : "<w:p/>";
  const lastSection = model.sections.length - 1;
  const lastSectionModel = model.sections[lastSection];
  const sectPr =
    includeOpaque &&
    model.sourceXml?.docxSectPr !== undefined &&
    !(lastSectionModel && sectionNeedsFreshSectPr(lastSectionModel, lastSection))
      ? model.sourceXml.docxSectPr
      : sectionPropertiesXml(
          sectionPageSize(model, lastSection),
          sectionMargins(model, lastSection),
          lastSectionModel?.start,
          sectionHeaderFooterOf(lastSection),
          lastSectionModel?.titlePage === true,
          lastSectionModel?.pageNumberStart,
          lastSectionModel?.columns,
          // Note numbering settings (SPEC-031) live on the body-level sectPr
          // — the one the reader reads them back from.
          notePrXml("footnotePr", model.footnoteSettings) +
            notePrXml("endnotePr", model.endnoteSettings),
        );
  return (
    `${XML_DECLARATION}<w:document xmlns:w="${WORD_NAMESPACE}">` +
    `<w:body>${body}${sectPr}</w:body></w:document>`
  );
}

// "Heading1".."Heading9" get Word's built-in style names ("heading 1"...);
// anything else keeps its id as the name. Used only for ids no definition
// exists for — defined styles (the model's catalogue or the built-ins,
// SPEC-019) carry their own names.
function styleNameFor(styleId: string): string {
  const match = /^Heading([1-9])$/.exec(styleId);
  return match ? `heading ${match[1]}` : styleId;
}

/** Every style id the model's paragraphs reference, in first-use order. */
function collectReferencedStyleIds(model: DocumentModel): Set<string> {
  const referencedIds = new Set<string>();
  for (const section of model.sections) {
    for (const block of section.blocks) {
      if (block.type === "image" || block.type === "break" || block.type === "toc") {
        continue;
      }
      const paragraphs =
        block.type === "table"
          ? block.rows.flatMap((row) => row.cells.flatMap((cell) => cell.blocks))
          : [block];
      for (const paragraph of paragraphs) {
        if (paragraph.styleId) {
          referencedIds.add(paragraph.styleId);
        }
      }
    }
  }
  return referencedIds;
}

/** A style definition's w:pPr/w:rPr from its own formatting (SPEC-019).
 * Style formatting is safe to declare even though every paragraph also
 * carries its merged result as direct formatting: the reader resolves the
 * style first and merges the (complete) direct formatting over it, which
 * reproduces the identical model. */
function stylePropertiesXml(definition: NamedStyle): string {
  if (definition.type === "table") {
    // Table styles (SPEC-042): the resolved border/shading set bakes onto
    // every cell at apply time (commands.ts's applyTableStyle) — this
    // declaration is written for a real .docx consumer's benefit (a plain
    // Word open still sees a valid table style), never re-parsed by this
    // app's own reader (table styles always come back `unmodeled`, see
    // docx/styles.ts — banding regions are out of scope, SPEC-042).
    const tcPrParts = [cellBordersXml(definition.cellBorders), cellShadingXml(definition.cellShadingFill)].filter(
      (part) => part.length > 0,
    );
    const tcPr = tcPrParts.length > 0 ? `<w:tcPr>${tcPrParts.join("")}</w:tcPr>` : "";
    return runPropertiesXml(definition.runFormatting ?? {}, undefined) + tcPr;
  }
  const pPrParts = paragraphFormattingParts(definition.paragraphFormatting ?? {});
  if (definition.headingLevel !== undefined) {
    pPrParts.push(`<w:outlineLvl w:val="${definition.headingLevel - 1}"/>`);
  }
  const pPr =
    definition.type !== "character" && pPrParts.length > 0 ? `<w:pPr>${pPrParts.join("")}</w:pPr>` : "";
  return pPr + runPropertiesXml(definition.runFormatting ?? {}, undefined);
}

/** The `w:type` a catalogue entry declares (SPEC-042): "list" has no OOXML
 * family of its own — a list style is an ordinary paragraph style whose
 * paragraphs also carry numbering (see styles.ts's `inferListStyleTypes`
 * doc comment) — so it declares "paragraph" like every plain paragraph
 * style. */
function wireStyleType(type: NamedStyle["type"]): "paragraph" | "character" | "table" {
  return type === "character" || type === "table" ? type : "paragraph";
}

/** An `unmodeled` style's original XML (SPEC-042 AC-18), re-emitted
 * verbatim except for the user-visible name — renaming (AC-6) is still
 * supported even for a style whose formatting the editor cannot express. */
function unmodeledStyleXml(definition: NamedStyle): string {
  if (!definition.rawXml) {
    return (
      `<w:style w:type="${wireStyleType(definition.type)}" w:styleId="${escapeXml(definition.id)}">` +
      `<w:name w:val="${escapeXml(definition.name ?? definition.id)}"/></w:style>`
    );
  }
  try {
    const wrapper = parseXml(`<w:root xmlns:w="${WORD_NAMESPACE}">${definition.rawXml}</w:root>`);
    const el = wrapper.documentElement.firstElementChild;
    if (!el) {
      return definition.rawXml;
    }
    if (definition.name !== undefined) {
      const nameEl = childElements(el, WORD_NAMESPACE, "name")[0];
      if (nameEl) {
        nameEl.setAttributeNS(WORD_NAMESPACE, "w:val", definition.name);
      } else {
        const created = wrapper.createElementNS(WORD_NAMESPACE, "w:name");
        created.setAttributeNS(WORD_NAMESPACE, "w:val", definition.name);
        el.insertBefore(created, el.firstElementChild);
      }
    }
    return serializeElement(el);
  } catch {
    return definition.rawXml;
  }
}

/** One w:style element for a defined style (SPEC-019 paragraph styles;
 * SPEC-042 widens this to character/list/table styles and to styles the
 * editor cannot fully model, which are re-emitted verbatim instead). */
function styleDefinitionXml(definition: NamedStyle, isDefault: boolean): string {
  if (definition.unmodeled) {
    return unmodeledStyleXml(definition);
  }
  const defaultAttr = isDefault ? ' w:default="1"' : "";
  const basedOn =
    definition.basedOn !== undefined
      ? `<w:basedOn w:val="${escapeXml(definition.basedOn)}"/>`
      : "";
  const next = definition.next !== undefined ? `<w:next w:val="${escapeXml(definition.next)}"/>` : "";
  return (
    `<w:style w:type="${wireStyleType(definition.type)}"${defaultAttr} w:styleId="${escapeXml(definition.id)}">` +
    `<w:name w:val="${escapeXml(definition.name ?? definition.id)}"/>` +
    basedOn +
    next +
    stylePropertiesXml(definition) +
    "</w:style>"
  );
}

/** The styles part: docDefaults carrying the document's default font, the
 * required default "Normal" style, and a definition for every style the
 * model's catalogue holds or its paragraphs reference (SPEC-019) — with the
 * style's own formatting when a definition exists (the catalogue, or the
 * built-in set for ids applied in the editor), name-only otherwise. */
export function buildStylesXml(model: DocumentModel): string {
  const { fontFamily, fontSizePt, language } = model.defaultStyles;
  const family = escapeXml(fontFamily);

  // Definitions come from the model's catalogue alone — never the built-in
  // set: a style applied in the editor is materialized into the catalogue
  // by the command (see commands.applyParagraphStyle), while a model
  // carrying no catalogue must reproduce the old name-only declarations
  // exactly (its styles' formatting is baked into the paragraphs).
  const catalogueById = new Map((model.styles ?? []).map((style) => [style.id, style]));

  // Declared ids: the catalogue in its own order, then referenced ids, then
  // basedOn ancestors of anything declared — closed so no declared style
  // references an undeclared parent.
  const queue = [...catalogueById.keys(), ...collectReferencedStyleIds(model)];
  const declared: string[] = [];
  const seen = new Set<string>(["Normal"]);
  while (queue.length > 0) {
    const id = queue.shift()!;
    if (seen.has(id)) {
      continue;
    }
    seen.add(id);
    declared.push(id);
    const parent = catalogueById.get(id)?.basedOn;
    if (parent !== undefined) {
      queue.push(parent);
    }
  }

  const normal = catalogueById.get("Normal");
  let styles =
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">' +
    '<w:name w:val="Normal"/>' +
    (normal ? stylePropertiesXml(normal) : "") +
    "</w:style>";
  for (const id of declared) {
    const definition = catalogueById.get(id);
    if (definition) {
      styles += styleDefinitionXml(definition, false);
    } else {
      styles +=
        `<w:style w:type="paragraph" w:styleId="${escapeXml(id)}">` +
        `<w:name w:val="${escapeXml(styleNameFor(id))}"/>` +
        '<w:basedOn w:val="Normal"/></w:style>';
    }
  }

  return (
    `${XML_DECLARATION}<w:styles xmlns:w="${WORD_NAMESPACE}">` +
    "<w:docDefaults><w:rPrDefault><w:rPr>" +
    `<w:rFonts w:ascii="${family}" w:hAnsi="${family}"/>` +
    `<w:sz w:val="${ptToHalfPoints(fontSizePt)}"/>` +
    (language !== undefined ? `<w:lang w:val="${escapeXml(language)}"/>` : "") +
    "</w:rPr></w:rPrDefault></w:docDefaults>" +
    styles +
    "</w:styles>"
  );
}

/** A preserved styles.xml with the model's style changes applied (SPEC-019),
 * or null when nothing needs changing — the caller then keeps the original
 * part byte-identical (SPEC-008). A style is rewritten only when the
 * model's catalogue entry no longer matches what the part defines ("Update
 * <style> to match"), and a definition is added for a catalogue entry or a
 * referenced built-in the part doesn't define (a heading style applied in
 * the editor must be a real style in the saved file, AC-15). Rewritten
 * styles keep their unmodelled attributes and children (w:default, w:qFormat,
 * w:next, …); only w:name, w:basedOn, w:pPr and w:rPr are regenerated. */
export function patchStylesXml(model: DocumentModel, originalXml: string): string | null {
  if (!model.styles) {
    return null;
  }
  let doc: Document;
  try {
    doc = parseXml(originalXml);
  } catch {
    return null;
  }
  const original = new Map(
    catalogToNamedStyles(parseStylesCatalog(doc)).map((style) => [style.id, style]),
  );
  const upserts: NamedStyle[] = [];
  for (const entry of model.styles) {
    const existing = original.get(entry.id);
    if (!existing || !equalStructural(existing, entry)) {
      upserts.push(entry);
    }
  }
  if (upserts.length === 0) {
    return null;
  }

  const root = doc.documentElement;
  const importStyleElement = (definition: NamedStyle): Element => {
    const wrapper = parseXml(
      `<w:styles xmlns:w="${WORD_NAMESPACE}">${styleDefinitionXml(definition, false)}</w:styles>`,
    );
    return doc.importNode(wrapper.documentElement.firstElementChild!, true) as Element;
  };
  for (const definition of upserts) {
    // Matched by styleId alone (SPEC-042): a style's family no longer maps
    // 1:1 onto "paragraph", and a styleId is unique across families in a
    // real styles.xml.
    const existing = childElements(root, WORD_NAMESPACE, "style").find(
      (el) => wordAttr(el, "styleId") === definition.id,
    );
    const replacement = importStyleElement(definition);
    if (!existing) {
      root.appendChild(replacement);
      continue;
    }
    if (definition.unmodeled) {
      // unmodeledStyleXml already re-emitted `existing`'s own XML verbatim
      // (only the name patched) — no attribute/child merge needed, unlike
      // the modeled branch below which regenerates from scratch.
      root.replaceChild(replacement, existing);
      continue;
    }
    // Carry the original element's unmodelled attributes (w:default, …) and
    // children over: the regenerated element replaces only what the model
    // expresses.
    for (const attr of Array.from(existing.attributes)) {
      if (
        attr.namespaceURI === WORD_NAMESPACE &&
        (attr.localName === "type" || attr.localName === "styleId")
      ) {
        continue;
      }
      replacement.setAttributeNodeNS(attr.cloneNode(true) as Attr);
    }
    const regenerated = new Set(["name", "basedOn", "next", "pPr", "rPr", "tcPr"]);
    // CT_Style puts pPr/rPr/tcPr last, so preserved children slot in before them.
    const propertiesStart =
      Array.from(replacement.children).find(
        (el) =>
          el.namespaceURI === WORD_NAMESPACE &&
          (el.localName === "pPr" || el.localName === "rPr" || el.localName === "tcPr"),
      ) ?? null;
    for (const child of Array.from(existing.children)) {
      if (child.namespaceURI === WORD_NAMESPACE && regenerated.has(child.localName)) {
        continue;
      }
      replacement.insertBefore(doc.importNode(child, true), propertiesStart);
    }
    root.replaceChild(replacement, existing);
  }
  return `${XML_DECLARATION}${serializeElement(root)}`;
}

/** Resolves the model's `properties` into `buildCorePropertiesXml`'s input:
 * the Description-tab "Title" field falls back to the model's own title
 * (SPEC-002) when blank, matching this writer's pre-SPEC-041 behaviour for a
 * document that never set one explicitly. */
function coreInputFrom(properties: DocumentProperties | undefined, fallbackTitle: string): CorePropertiesInput {
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
  };
}

/** `<Default>` entries for the image extensions the model's media plan
 * actually uses (SPEC-024 AC-26) — content types are declared once per
 * extension, package-wide, never per part (unlike word/document.xml's own
 * Override). */
function imageContentTypesXml(extensions: string[]): string {
  return extensions
    .map((ext) => `<Default Extension="${ext}" ContentType="${CONTENT_TYPE_BY_EXTENSION[ext]}"/>`)
    .join("");
}

export function buildContentTypesXml(withNumbering = false, imageExtensions: string[] = []): string {
  const numbering = withNumbering
    ? `<Override PartName="/word/numbering.xml" ContentType="${NUMBERING_CONTENT_TYPE}"/>`
    : "";
  return (
    `${XML_DECLARATION}<Types xmlns="${CONTENT_TYPES_NAMESPACE}">` +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    imageContentTypesXml(imageExtensions) +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
    numbering +
    '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' +
    "</Types>"
  );
}

export function buildPackageRelsXml(): string {
  return (
    `${XML_DECLARATION}<Relationships xmlns="${RELATIONSHIPS_NAMESPACE}">` +
    `<Relationship Id="rId1" Type="${OFFICE_DOCUMENT_RELATIONSHIP}" Target="word/document.xml"/>` +
    `<Relationship Id="rId2" Type="${CORE_PROPERTIES_RELATIONSHIP}" Target="docProps/core.xml"/>` +
    "</Relationships>"
  );
}

export function buildDocumentRelsXml(withNumbering = false, mediaRelationshipsXml = ""): string {
  const numbering = withNumbering
    ? `<Relationship Id="rId2" Type="${NUMBERING_RELATIONSHIP}" Target="numbering.xml"/>`
    : "";
  return (
    `${XML_DECLARATION}<Relationships xmlns="${RELATIONSHIPS_NAMESPACE}">` +
    `<Relationship Id="rId1" Type="${STYLES_RELATIONSHIP}" Target="styles.xml"/>` +
    numbering +
    mediaRelationshipsXml +
    "</Relationships>"
  );
}

export interface WriteDocxOptions {
  /** The dcterms:created/modified timestamp; defaults to now. Injectable so
   * tests can assert an exact core.xml. */
  modified?: Date;
}

const CORE_CONTENT_TYPE = "application/vnd.openxmlformats-package.core-properties+xml";

/** Adds the core-properties Override to a preserved [Content_Types].xml —
 * needed when the source package had no docProps/core.xml but the save adds
 * one to carry the edited title. */
function withCoreContentType(contentTypesXml: string): string {
  if (contentTypesXml.includes('PartName="/docProps/core.xml"')) {
    return contentTypesXml;
  }
  return contentTypesXml.replace(
    /<\/Types>\s*$/,
    `<Override PartName="/docProps/core.xml" ContentType="${CORE_CONTENT_TYPE}"/></Types>`,
  );
}

/** Adds a core-properties relationship (under a fresh, unused rId) to a
 * preserved _rels/.rels for the same missing-core.xml case. */
function withCoreRelationship(relsXml: string): string {
  const rels = parseRelationships(relsXml);
  let n = rels.size + 1;
  while (rels.has(`rId${n}`)) {
    n++;
  }
  return relsXml.replace(
    /<\/Relationships>\s*$/,
    `<Relationship Id="rId${n}" Type="${CORE_PROPERTIES_RELATIONSHIP}" Target="docProps/core.xml"/>` +
      "</Relationships>",
  );
}

/** Adds the custom-properties Override to a preserved [Content_Types].xml —
 * needed when a save adds docProps/custom.xml to a package that had none
 * (SPEC-041: the model gained its first custom property). */
function withCustomPropsContentType(contentTypesXml: string): string {
  if (contentTypesXml.includes('PartName="/docProps/custom.xml"')) {
    return contentTypesXml;
  }
  return contentTypesXml.replace(
    /<\/Types>\s*$/,
    `<Override PartName="/docProps/custom.xml" ContentType="${CUSTOM_PROPERTIES_CONTENT_TYPE}"/></Types>`,
  );
}

/** Adds a custom-properties relationship (under a fresh, unused rId) to a
 * preserved _rels/.rels for the same missing-custom.xml case. */
function withCustomPropsRelationship(relsXml: string): string {
  const rels = parseRelationships(relsXml);
  let n = rels.size + 1;
  while (rels.has(`rId${n}`)) {
    n++;
  }
  return relsXml.replace(
    /<\/Relationships>\s*$/,
    `<Relationship Id="rId${n}" Type="${CUSTOM_PROPERTIES_RELATIONSHIP}" Target="docProps/custom.xml"/>` +
      "</Relationships>",
  );
}

/** Removes the custom-properties Override from a preserved
 * [Content_Types].xml — the inverse of `withCustomPropsContentType`, needed
 * when the last custom property is deleted from a source that had declared
 * the part. */
function withoutCustomPropsContentType(contentTypesXml: string): string {
  return contentTypesXml.replace(/<Override PartName="\/docProps\/custom\.xml"[^>]*\/>/, "");
}

/** Removes the custom-properties relationship — the inverse of
 * `withCustomPropsRelationship`. */
function withoutCustomPropsRelationship(relsXml: string): string {
  return relsXml.replace(new RegExp(`<Relationship[^>]*Type="${CUSTOM_PROPERTIES_RELATIONSHIP}"[^>]*/>`), "");
}

/** Adds the numbering Override to a preserved [Content_Types].xml — needed
 * when a save adds a numbering part to a package that had none (SPEC-018). */
function withNumberingContentType(contentTypesXml: string, partName: string): string {
  if (contentTypesXml.includes(`PartName="/${partName}"`)) {
    return contentTypesXml;
  }
  return contentTypesXml.replace(
    /<\/Types>\s*$/,
    `<Override PartName="/${partName}" ContentType="${NUMBERING_CONTENT_TYPE}"/></Types>`,
  );
}

/** Adds a numbering relationship (under a fresh, unused rId) to a preserved
 * document rels part for the same added-numbering case. */
function withNumberingRelationship(relsXml: string, target: string): string {
  const rels = parseRelationships(relsXml);
  let n = rels.size + 1;
  while (rels.has(`rId${n}`)) {
    n++;
  }
  return relsXml.replace(
    /<\/Relationships>\s*$/,
    `<Relationship Id="rId${n}" Type="${NUMBERING_RELATIONSHIP}" Target="${target}"/>` +
      "</Relationships>",
  );
}

/** Adds Default content-type entries for the image extensions a save's
 * media plan introduces (SPEC-024 AC-26) — needed when a package that
 * previously had no images gets one, or gains an extension it didn't
 * already declare. */
function withImageContentTypes(contentTypesXml: string, extensions: string[]): string {
  const missing = extensions.filter((ext) => !contentTypesXml.includes(`Extension="${ext}"`));
  if (missing.length === 0) {
    return contentTypesXml;
  }
  const defaults = missing
    .map((ext) => `<Default Extension="${ext}" ContentType="${CONTENT_TYPE_BY_EXTENSION[ext]}"/>`)
    .join("");
  return contentTypesXml.replace(/<\/Types>\s*$/, `${defaults}</Types>`);
}

/** Appends the media plan's `<Relationship>` elements to a preserved
 * document rels part — the ids were already allocated past every id that
 * part uses (planMediaWrite's `takenRelIds`), so this is always a pure
 * append. */
function withMediaRelationships(relsXml: string, relationshipsXml: string): string {
  return relsXml.replace(/<\/Relationships>\s*$/, `${relationshipsXml}</Relationships>`);
}

/** Serialises the model into `.docx` bytes plus the list of items that could
 * not be represented (SPEC-008).
 *
 * A model opened from a .docx keeps its retained source package: every part
 * the reader did not model is written back byte-identical, and only the main
 * document part and docProps/core.xml (title/timestamp) are rebuilt —
 * `dropped` is empty. Anything else (blank documents, models opened from
 * .odt) is written from scratch as a valid OPC ZIP containing
 * [Content_Types].xml, _rels/.rels, word/document.xml, word/styles.xml and
 * docProps/core.xml (SPEC-006 AC-3..8), with unrepresentable .odt content
 * reported in `dropped`. */
export async function writeDocx(
  sourceModel: DocumentModel,
  options: WriteDocxOptions = {},
): Promise<WriteResult> {
  // "Remove personal information on save" (SPEC-041 AC-12..14): a pure
  // model transform applied before anything below ever reads an author name
  // or a comment/tracked-change timestamp, so every write path downstream
  // (fresh or same-format) automatically writes the sanitised values with no
  // format-specific stripping logic of its own.
  const model = sourceModel.properties?.removePersonalInfoOnSave ? stripPersonalInfo(sourceModel) : sourceModel;
  const modified = options.modified ?? new Date();
  const title = model.title.trim().length > 0 ? model.title : DEFAULT_TITLE;
  const coreInput = coreInputFrom(model.properties, title);
  const customPropsXml = buildCustomPropertiesXml(
    model.properties?.customProperties ?? [],
    model.properties?.unsupportedCustomProperties ?? [],
  );
  const encoder = new TextEncoder();

  const preservedParts = model.format === "docx" ? opaquePartBytes(model) : new Map<string, Uint8Array>();
  if (preservedParts.size === 0) {
    const plan = planNumbering(model);
    // rId1 is styles.xml and (if present) rId2 is numbering.xml, both
    // hardcoded by buildDocumentRelsXml — hyperlink and media ids must start
    // past them (and past each other).
    const baseTaken = new Set(plan.isEmpty ? ["rId1"] : ["rId1", "rId2"]);
    const hyperlinkPlan = planHyperlinkWrite(model, baseTaken);
    const existingMediaPartNames = new Set<string>();
    const mediaPlan = planMediaWrite(
      model,
      new Set([...baseTaken, ...hyperlinkPlan.relIds]),
      existingMediaPartNames,
    );
    // Note parts (SPEC-031) mint their relationship ids past the base set,
    // the comments part (SPEC-038) past notes, header/footer parts
    // (SPEC-029) past everything — their own images share the same
    // word/media/imageN.ext naming pool (existingMediaPartNames, mutated by
    // planMediaWrite above).
    const notesPlan = planNotesWrite(
      model,
      new Set([...baseTaken, ...hyperlinkPlan.relIds, ...mediaPlan.relIds]),
    );
    const commentsPlan = planCommentsWrite(
      model,
      new Set([...baseTaken, ...hyperlinkPlan.relIds, ...mediaPlan.relIds, ...notesPlan.relIds]),
    );
    const headerFooterPlan = planHeaderFooterWrite(
      model,
      new Set([
        ...baseTaken,
        ...hyperlinkPlan.relIds,
        ...mediaPlan.relIds,
        ...notesPlan.relIds,
        ...commentsPlan.relIds,
      ]),
      existingMediaPartNames,
    );
    const needsSettingsPart = headerFooterPlan.evenAndOddHeaders || !!model.trackChangesRecording;
    const settingsOverride = needsSettingsPart
      ? `<Override PartName="/word/settings.xml" ContentType="${SETTINGS_CONTENT_TYPE}"/>`
      : "";
    const customPropsOverride =
      customPropsXml !== undefined
        ? `<Override PartName="/docProps/custom.xml" ContentType="${CUSTOM_PROPERTIES_CONTENT_TYPE}"/>`
        : "";
    const contentTypesXml = buildContentTypesXml(!plan.isEmpty, mediaPlan.extensions).replace(
      /<\/Types>\s*$/,
      `${notesPlan.contentTypesXml}${commentsPlan.contentTypesXml}${headerFooterPlan.contentTypesXml}${settingsOverride}${customPropsOverride}</Types>`,
    );
    const customPropsRel =
      customPropsXml !== undefined
        ? `<Relationship Id="rId3" Type="${CUSTOM_PROPERTIES_RELATIONSHIP}" Target="docProps/custom.xml"/>`
        : "";
    const entries: ZipEntryInput[] = [
      { name: "[Content_Types].xml", data: encoder.encode(contentTypesXml) },
      {
        name: "_rels/.rels",
        data: encoder.encode(buildPackageRelsXml().replace(/<\/Relationships>\s*$/, `${customPropsRel}</Relationships>`)),
      },
      {
        name: "word/document.xml",
        data: encoder.encode(
          buildDocumentXml(model, plan, mediaPlan, hyperlinkPlan, headerFooterPlan, notesPlan, commentsPlan),
        ),
      },
      {
        name: "word/_rels/document.xml.rels",
        data: encoder.encode(
          buildDocumentRelsXml(
            !plan.isEmpty,
            hyperlinkPlan.relationshipsXml +
              mediaPlan.relationshipsXml +
              notesPlan.relationshipsXml +
              commentsPlan.relationshipsXml +
              headerFooterPlan.relationshipsXml,
          ),
        ),
      },
      { name: "word/styles.xml", data: encoder.encode(buildStylesXml(model)) },
      { name: "docProps/core.xml", data: encoder.encode(buildCorePropertiesXml(coreInput, modified)) },
      ...mediaPlan.zipEntries,
    ];
    if (customPropsXml !== undefined) {
      entries.push({ name: "docProps/custom.xml", data: encoder.encode(customPropsXml) });
    }
    if (!plan.isEmpty) {
      entries.splice(5, 0, {
        name: "word/numbering.xml",
        data: encoder.encode(buildNumberingXml(plan)),
      });
    }
    if (needsSettingsPart) {
      entries.push({
        name: "word/settings.xml",
        data: encoder.encode(buildSettingsXml(!!model.trackChangesRecording)),
      });
    }
    for (const partPlan of headerFooterPlan.parts) {
      const rootTag = partPlan.partName.includes("/header") ? "hdr" : "ftr";
      entries.push({
        name: partPlan.partName,
        data: encoder.encode(headerFooterPartXml(rootTag, model.format === "docx", partPlan)),
      });
      entries.push(...partPlan.mediaPlan.zipEntries);
      const relsXml = headerFooterPartRelsXml(partPlan);
      if (relsXml !== null) {
        entries.push({ name: relsPathFor(partPlan.partName), data: encoder.encode(relsXml) });
      }
    }
    for (const notesPart of notesPlan.parts) {
      entries.push({
        name: notesPart.partName,
        data: encoder.encode(notesPartXml(notesPart, model.format === "docx", notesPlan.noteIds)),
      });
      const relsXml = notesPartRelsXml(notesPart);
      if (relsXml !== null) {
        entries.push({ name: relsPathFor(notesPart.partName), data: encoder.encode(relsXml) });
      }
    }
    for (const commentsPart of commentsPlan.parts) {
      entries.push({
        name: commentsPart.partName,
        data: encoder.encode(commentsPartXml(commentsPart, commentsPlan.commentIds)),
      });
    }
    const bytes = await zip(entries);
    return { bytes, dropped: collectDropRecords(model, "docx") };
  }

  // Same-format save of an opened .docx. The preserved _rels/.rels names the
  // main document and core-properties parts, so the rebuilt parts land at
  // the package's own paths regardless of producer layout.
  const decoder = new TextDecoder("utf-8");
  const relsBytes = preservedParts.get("_rels/.rels");
  const relsText = relsBytes ? decoder.decode(relsBytes) : buildPackageRelsXml();
  const rels = Array.from(parseRelationships(relsText).values());
  const officeDocRel = rels.find((rel) => rel.type === OFFICE_DOCUMENT_RELATIONSHIP);
  const mainPartName = officeDocRel ? resolvePartPath("", officeDocRel.target) : "word/document.xml";
  const coreRel = rels.find((rel) => rel.type === CORE_PROPERTIES_RELATIONSHIP);
  const corePartName = coreRel ? resolvePartPath("", coreRel.target) : "docProps/core.xml";
  const customRel = rels.find((rel) => rel.type === CUSTOM_PROPERTIES_RELATIONSHIP);
  const customPartName = customRel ? resolvePartPath("", customRel.target) : "docProps/custom.xml";
  // A save always fully regenerates custom.xml from the model's typed
  // properties (SPEC-041), so it is dropped from `preservedParts` below like
  // core.xml. Its content-type/relationship declarations are added when the
  // model gains its first custom property, and removed when the last one is
  // deleted from a source that had declared the part (an Override pointing
  // at a part the save no longer writes would make the package invalid).
  const addCustomPropsDecls = customPropsXml !== undefined && !customRel;
  const removeCustomPropsDecls = customPropsXml === undefined && !!customRel;

  // Numbering (SPEC-018). When the model holds lists, its plan allocates
  // ids above whatever the preserved numbering part already defines and the
  // definitions are merged in — the original ones stay verbatim because
  // opaque content (lists inside tables, say) may reference them. A package
  // that had no numbering part gets one, declared in the rels/content types
  // like the added-core.xml path just below.
  const mainRelsName = relsPathFor(mainPartName);
  const mainRelsBytes = preservedParts.get(mainRelsName);
  const mainRelsText = mainRelsBytes ? decoder.decode(mainRelsBytes) : undefined;
  const numberingRel = mainRelsText
    ? Array.from(parseRelationships(mainRelsText).values()).find(
        (rel) => rel.type === NUMBERING_RELATIONSHIP,
      )
    : undefined;
  const mainDir = mainPartName.includes("/")
    ? mainPartName.slice(0, mainPartName.lastIndexOf("/") + 1)
    : "";
  const numberingPartName = numberingRel
    ? resolvePartPath(mainPartName, numberingRel.target)
    : `${mainDir}numbering.xml`;
  const preservedNumbering = preservedParts.get(numberingPartName);
  const plan = planNumbering(
    model,
    preservedNumbering ? nextNumberingIds(decoder.decode(preservedNumbering)) : undefined,
  );
  const numberingOut = plan.isEmpty
    ? null
    : preservedNumbering
      ? mergeNumberingXml(decoder.decode(preservedNumbering), plan)
      : buildNumberingXml(plan);
  const addNumberingDecls = numberingOut !== null && !numberingRel;

  // Styles (SPEC-019): the preserved styles part is kept byte-identical
  // unless the model's catalogue diverges from it — a style updated in the
  // editor, or a built-in style applied that the part doesn't define.
  const stylesRel = mainRelsText
    ? Array.from(parseRelationships(mainRelsText).values()).find(
        (rel) => rel.type === STYLES_RELATIONSHIP,
      )
    : undefined;
  const stylesPartName = stylesRel
    ? resolvePartPath(mainPartName, stylesRel.target)
    : `${mainDir}styles.xml`;
  const preservedStyles = preservedParts.get(stylesPartName);
  const patchedStyles = preservedStyles
    ? patchStylesXml(model, decoder.decode(preservedStyles))
    : null;

  // Images (SPEC-024): media ids start past whatever the preserved main
  // rels part already uses, and past whatever package part names already
  // exist, so a freshly written word/media/imageN.ext can never collide.
  const mainRelsTaken = mainRelsText ? takenRelIdsOf(parseRelationships(mainRelsText).values()) : new Set<string>();
  // Hyperlinks (SPEC-025) mint their rel ids first, past the preserved
  // part's own ids; media ids then start past both.
  const hyperlinkPlan = planHyperlinkWrite(model, mainRelsTaken);
  const existingPartNames = new Set(preservedParts.keys());
  const mediaPlan = planMediaWrite(
    model,
    new Set([...mainRelsTaken, ...hyperlinkPlan.relIds]),
    existingPartNames,
  );
  const addMediaDecls = mediaPlan.relationshipsXml.length > 0;
  const addHyperlinkDecls = hyperlinkPlan.relationshipsXml.length > 0;

  // Note parts (SPEC-031): the reader consumed the source's own
  // footnotes.xml/endnotes.xml (they are not in preservedParts), so the
  // canonical names are normally free again; ids mint past everything above.
  const notesPlan = planNotesWrite(
    model,
    new Set([...mainRelsTaken, ...hyperlinkPlan.relIds, ...mediaPlan.relIds]),
    existingPartNames,
  );
  const addNotesDecls = notesPlan.relationshipsXml.length > 0;

  // The comments part (SPEC-038): the reader consumed the source's own
  // comments.xml, so its canonical name is normally free again too; ids
  // mint past the note parts.
  const commentsPlan = planCommentsWrite(
    model,
    new Set([...mainRelsTaken, ...hyperlinkPlan.relIds, ...mediaPlan.relIds, ...notesPlan.relIds]),
    existingPartNames,
  );
  const addCommentsDecls = commentsPlan.relationshipsXml.length > 0;

  // Headers, footers and page numbering (SPEC-029): relationship ids mint
  // past everything above; header/footer part names and their own images
  // avoid every preserved/already-minted name the same way media does.
  const headerFooterPlan = planHeaderFooterWrite(
    model,
    new Set([
      ...mainRelsTaken,
      ...hyperlinkPlan.relIds,
      ...mediaPlan.relIds,
      ...notesPlan.relIds,
      ...commentsPlan.relIds,
    ]),
    existingPartNames,
    existingPartNames,
  );
  const addHeaderFooterDecls = headerFooterPlan.relationshipsXml.length > 0;

  // word/settings.xml (SPEC-029): patched in place when present (it carries
  // many other unmodelled document-wide flags), created fresh only when the
  // model wants different odd/even headers and the source had no settings
  // part at all.
  const settingsRel = mainRelsText
    ? Array.from(parseRelationships(mainRelsText).values()).find((rel) => rel.type === SETTINGS_RELATIONSHIP)
    : undefined;
  const settingsPartName = settingsRel ? resolvePartPath(mainPartName, settingsRel.target) : `${mainDir}settings.xml`;
  const preservedSettings = preservedParts.get(settingsPartName);
  const trackChangesRecording = !!model.trackChangesRecording;
  const patchedSettings = preservedSettings
    ? patchSettingsXml(headerFooterPlan.evenAndOddHeaders, trackChangesRecording, decoder.decode(preservedSettings))
    : null;
  const addSettingsPart = !preservedSettings && (headerFooterPlan.evenAndOddHeaders || trackChangesRecording);

  const entries: ZipEntryInput[] = [];
  if (!relsBytes) {
    entries.push({ name: "_rels/.rels", data: encoder.encode(relsText) });
  }
  for (const [name, data] of preservedParts) {
    if (name === mainPartName || name === corePartName || name === customPartName) {
      continue; // Rebuilt below (the reader excludes them from opaque anyway).
    }
    if (numberingOut !== null && name === numberingPartName) {
      entries.push({ name, data: encoder.encode(numberingOut) });
    } else if (patchedStyles !== null && name === stylesPartName) {
      entries.push({ name, data: encoder.encode(patchedStyles) });
    } else if (patchedSettings !== null && name === settingsPartName) {
      entries.push({ name, data: encoder.encode(patchedSettings) });
    } else if (
      (addNumberingDecls ||
        addMediaDecls ||
        addHyperlinkDecls ||
        addNotesDecls ||
        addCommentsDecls ||
        addHeaderFooterDecls ||
        addSettingsPart) &&
      name === mainRelsName
    ) {
      let rels = mainRelsText!;
      if (addNumberingDecls) {
        rels = withNumberingRelationship(rels, "numbering.xml");
      }
      if (addHyperlinkDecls) {
        rels = withMediaRelationships(rels, hyperlinkPlan.relationshipsXml);
      }
      if (addMediaDecls) {
        rels = withMediaRelationships(rels, mediaPlan.relationshipsXml);
      }
      if (addNotesDecls) {
        rels = withMediaRelationships(rels, notesPlan.relationshipsXml);
      }
      if (addCommentsDecls) {
        rels = withMediaRelationships(rels, commentsPlan.relationshipsXml);
      }
      if (addHeaderFooterDecls) {
        rels = withMediaRelationships(rels, headerFooterPlan.relationshipsXml);
      }
      if (addSettingsPart) {
        rels = withMediaRelationships(
          rels,
          `<Relationship Id="${nextFreeRelId(new Set([...mainRelsTaken, ...hyperlinkPlan.relIds, ...mediaPlan.relIds, ...notesPlan.relIds, ...commentsPlan.relIds]))}" Type="${SETTINGS_RELATIONSHIP}" Target="settings.xml"/>`,
        );
      }
      entries.push({ name, data: encoder.encode(rels) });
    } else if ((!coreRel || addCustomPropsDecls || removeCustomPropsDecls) && name === "_rels/.rels") {
      let rootRels = relsText;
      if (!coreRel) {
        rootRels = withCoreRelationship(rootRels);
      }
      if (addCustomPropsDecls) {
        rootRels = withCustomPropsRelationship(rootRels);
      }
      if (removeCustomPropsDecls) {
        rootRels = withoutCustomPropsRelationship(rootRels);
      }
      entries.push({ name, data: encoder.encode(rootRels) });
    } else if (
      name === "[Content_Types].xml" &&
      (!coreRel ||
        addNumberingDecls ||
        addMediaDecls ||
        addNotesDecls ||
        addCommentsDecls ||
        addHeaderFooterDecls ||
        addSettingsPart ||
        addCustomPropsDecls ||
        removeCustomPropsDecls)
    ) {
      let contentTypes = decoder.decode(data);
      if (!coreRel) {
        contentTypes = withCoreContentType(contentTypes);
      }
      if (addCustomPropsDecls) {
        contentTypes = withCustomPropsContentType(contentTypes);
      }
      if (removeCustomPropsDecls) {
        contentTypes = withoutCustomPropsContentType(contentTypes);
      }
      if (addNumberingDecls) {
        contentTypes = withNumberingContentType(contentTypes, numberingPartName);
      }
      if (addMediaDecls) {
        contentTypes = withImageContentTypes(contentTypes, mediaPlan.extensions);
      }
      if (addNotesDecls) {
        contentTypes = contentTypes.replace(/<\/Types>\s*$/, `${notesPlan.contentTypesXml}</Types>`);
      }
      if (addCommentsDecls) {
        contentTypes = contentTypes.replace(/<\/Types>\s*$/, `${commentsPlan.contentTypesXml}</Types>`);
      }
      if (addHeaderFooterDecls) {
        contentTypes = contentTypes.replace(/<\/Types>\s*$/, `${headerFooterPlan.contentTypesXml}</Types>`);
      }
      if (addSettingsPart) {
        contentTypes = contentTypes.replace(
          /<\/Types>\s*$/,
          `<Override PartName="/word/settings.xml" ContentType="${SETTINGS_CONTENT_TYPE}"/></Types>`,
        );
      }
      entries.push({ name, data: encoder.encode(contentTypes) });
    } else {
      entries.push({ name, data });
    }
  }
  if (numberingOut !== null && !preservedNumbering) {
    entries.push({ name: numberingPartName, data: encoder.encode(numberingOut) });
  }
  if (addSettingsPart) {
    entries.push({ name: "word/settings.xml", data: encoder.encode(buildSettingsXml(trackChangesRecording)) });
  }
  for (const partPlan of headerFooterPlan.parts) {
    const rootTag = partPlan.partName.includes("/header") ? "hdr" : "ftr";
    entries.push({
      name: partPlan.partName,
      data: encoder.encode(headerFooterPartXml(rootTag, model.format === "docx", partPlan)),
    });
    entries.push(...partPlan.mediaPlan.zipEntries);
    const relsXml = headerFooterPartRelsXml(partPlan);
    if (relsXml !== null) {
      entries.push({ name: relsPathFor(partPlan.partName), data: encoder.encode(relsXml) });
    }
  }
  for (const notesPart of notesPlan.parts) {
    entries.push({
      name: notesPart.partName,
      data: encoder.encode(notesPartXml(notesPart, model.format === "docx", notesPlan.noteIds)),
    });
    const relsXml = notesPartRelsXml(notesPart);
    if (relsXml !== null) {
      entries.push({ name: relsPathFor(notesPart.partName), data: encoder.encode(relsXml) });
    }
  }
  for (const commentsPart of commentsPlan.parts) {
    entries.push({
      name: commentsPart.partName,
      data: encoder.encode(commentsPartXml(commentsPart, commentsPlan.commentIds)),
    });
  }
  entries.push({
    name: mainPartName,
    data: encoder.encode(
      buildDocumentXml(model, plan, mediaPlan, hyperlinkPlan, headerFooterPlan, notesPlan, commentsPlan),
    ),
  });
  entries.push({ name: corePartName, data: encoder.encode(buildCorePropertiesXml(coreInput, modified)) });
  if (customPropsXml !== undefined) {
    entries.push({ name: customPartName, data: encoder.encode(customPropsXml) });
  }
  entries.push(...mediaPlan.zipEntries);

  return { bytes: await zip(entries), dropped: [] };
}
