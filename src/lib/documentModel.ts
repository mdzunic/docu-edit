// The in-memory document model: a serialisable tree every later spec reads
// and writes. Plain data only (no classes, no functions on the objects) so
// `JSON.stringify`/`JSON.parse` round-trips it without loss — that property
// is what SPEC-006/007 (save) and SPEC-004/005 (open) will rely on.

/** Character-level formatting applied to a run's text. Empty until SPEC-012
 * (character formatting) and later specs add fields such as bold/italic. */
export interface RunFormatting {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strikethrough?: boolean;
  /** e.g. "Calibri" */
  fontFamily?: string;
  fontSizePt?: number;
  /** CSS-ready colour, e.g. "#ff0000". */
  color?: string;
  /** OOXML highlight colour name, e.g. "yellow". */
  highlight?: string;
}

/** The smallest unit of text: a string plus the formatting applied to it.
 * Tabs and forced line breaks are preserved as literal "\t" / "\n"
 * characters in `text` — renderers use `white-space: pre-wrap` so both
 * display exactly as the source document intended. */
export interface Run {
  text: string;
  formatting: RunFormatting;
}

export type ParagraphAlignment = "left" | "center" | "right" | "justify";

/** Block-level formatting applied to a paragraph. Empty until SPEC-014 adds
 * a UI for it; SPEC-004 (the .docx reader) is the first writer. */
export interface ParagraphFormatting {
  alignment?: ParagraphAlignment;
  indentLeftCm?: number;
  indentRightCm?: number;
  /** Positive indents the first line; negative produces a hanging indent. */
  indentFirstLineCm?: number;
  spaceBeforePt?: number;
  spaceAfterPt?: number;
  /** Line height as a multiplier of the font size (CSS `line-height`-style),
   * used when the source expresses spacing as a proportional line rule. */
  lineSpacing?: number;
  /** Line height as an absolute point size, used when the source expresses
   * spacing as a fixed/minimum line rule. Takes precedence over
   * `lineSpacing` when both would otherwise apply. */
  lineSpacingPt?: number;
}

/** An inline XML element the reader does not model (SPEC-008) — e.g. a
 * `w:drawing`-carrying run or an ODF `draw:frame` — held verbatim and
 * anchored to its position between the paragraph's runs so the writer can
 * re-emit it in place. `xml` is a self-contained serialisation (namespace
 * declarations included); it is compared with XML canonicalization, never
 * byte-for-byte. */
export interface OpaqueInline {
  xml: string;
  /** The node sits immediately before `runs[runIndex]`; `runs.length` means
   * after the last run. */
  runIndex: number;
}

/** A paragraph is a block-level element made of inline runs. */
export interface Paragraph {
  type: "paragraph";
  runs: Run[];
  formatting?: ParagraphFormatting;
  /** Raw style reference (e.g. `w:pStyle`'s value), kept alongside the
   * resolved `formatting`/run formatting so later specs (save) can re-emit
   * a style reference instead of only direct formatting. */
  styleId?: string;
  /** 1-based outline level resolved from the paragraph's style chain;
   * undefined for body text. */
  headingLevel?: number;
  /** Unmodelled inline XML preserved through a same-format save (SPEC-008).
   * A side-list (rather than a run union) so `runs` stays a plain `Run[]`
   * for every renderer and editing consumer; order follows `runIndex`. */
  opaqueInlines?: OpaqueInline[];
}

/** A block-level XML element the reader does not model (SPEC-008) — e.g. a
 * `w:tbl`, a block-level `w:sdt` content control or an ODF `table:table` —
 * held verbatim, anchored to its position among the section's blocks, and
 * re-emitted in place on a same-format save. */
export interface OpaqueBlock {
  /** Self-contained XML serialisation, compared canonically (see
   * OpaqueInline.xml). */
  xml: string;
  /** The node sits immediately before `blocks[blockIndex]`; `blocks.length`
   * means after the last block. */
  blockIndex: number;
}

/** Block-level children of a section. Paragraph is the only kind for now;
 * later specs (tables, images) extend this union. */
export type Block = Paragraph;

/** A section groups block-level content; documents have one or more. */
export interface Section {
  blocks: Block[];
  /** Unmodelled block-level XML preserved through a same-format save
   * (SPEC-008). A side-list (rather than a Block union member) so `blocks`
   * stays renderable/editable without narrowing; order follows
   * `blockIndex`. */
  opaqueBlocks?: OpaqueBlock[];
}

/** Page dimensions in centimetres. */
export interface PageSize {
  widthCm: number;
  heightCm: number;
}

/** Page margins in centimetres. */
export interface Margins {
  topCm: number;
  rightCm: number;
  bottomCm: number;
  leftCm: number;
}

/** Document-wide defaults new content is created with. */
export interface DefaultStyles {
  pageSize: PageSize;
  margins: Margins;
  fontFamily: string;
  fontSizePt: number;
}

/** The source format a document was opened from, if any — drives the format
 * badge in the app bar. Undefined for documents created blank (SPEC-002). */
export type DocumentFormat = "docx" | "odt";

/** A package part the reader does not model (SPEC-008) — `customXml/`,
 * charts, themes, embedded objects, images, `word/webSettings.xml`, … —
 * retained for the lifetime of the document and written back byte-identical
 * on a save to the same format. */
export interface OpaquePart {
  /** Path within the source package, e.g. "word/charts/chart1.xml". */
  name: string;
  /** The part's original (decompressed) bytes, base64-encoded so the model
   * stays JSON-serialisable (see the header comment). */
  data: string;
}

/** Verbatim XML captured from the source package's modelled parts that the
 * model itself cannot express, re-emitted on a same-format save (SPEC-008).
 * Each field is format-specific; all are compared canonically. */
export interface SourceXml {
  /** .docx: the original body-level `w:sectPr`, emitted in place of a
   * generated one so unmodelled section properties (header/footer
   * references, columns, …) survive. Page size/margins cannot yet diverge
   * from `defaultStyles` — no spec edits them. */
  docxSectPr?: string;
  /** .odt: the original `office:font-face-decls` of content.xml. */
  odtFontFaceDecls?: string;
  /** .odt: original automatic-style definitions of content.xml other than
   * paragraph/text `style:style`s (graphic, table, list, … families) —
   * kept because opaque nodes may reference them. Paragraph/text styles are
   * regenerated from the model's resolved formatting instead. */
  odtAutomaticStyles?: string[];
}

/** The root of the document tree. */
export interface DocumentModel {
  id: string;
  title: string;
  sections: Section[];
  defaultStyles: DefaultStyles;
  format?: DocumentFormat;
  /** Unmodelled package parts retained from the opened file (SPEC-008). */
  opaque?: OpaquePart[];
  /** Unmodelled XML retained from the opened file's parsed parts. */
  sourceXml?: SourceXml;
  /** True once the user ticked "Don't show again for this document" on the
   * conversion report dialog (SPEC-008); persisted with the document. */
  suppressConversionReport?: boolean;
}

export const DEFAULT_TITLE = "Untitled document";

const A4_PAGE_SIZE: PageSize = { widthCm: 21, heightCm: 29.7 };

const STANDARD_MARGINS: Margins = {
  topCm: 2.54,
  rightCm: 2.54,
  bottomCm: 2.54,
  leftCm: 2.54,
};

// Calibri 11pt is Word's modern default body font; picked as the product
// default since OOXML (.docx) is this project's primary format.
const DEFAULT_FONT_FAMILY = "Calibri";
const DEFAULT_FONT_SIZE_PT = 11;

/** The product's built-in fallback styles (A4, standard margins, Calibri
 * 11pt) — the starting point for a blank document, and the fallback the
 * `.docx` reader falls back to for anything a source file doesn't specify. */
export function createDefaultStyles(): DefaultStyles {
  return {
    pageSize: { ...A4_PAGE_SIZE },
    margins: { ...STANDARD_MARGINS },
    fontFamily: DEFAULT_FONT_FAMILY,
    fontSizePt: DEFAULT_FONT_SIZE_PT,
  };
}

function createEmptyParagraph(): Paragraph {
  return { type: "paragraph", runs: [] };
}

/** Creates a new blank document: one section, one empty paragraph, and the
 * product's default styles. Pass an id to control it (e.g. a route param);
 * otherwise one is generated. */
export function createDocument(id: string = crypto.randomUUID()): DocumentModel {
  return {
    id,
    title: DEFAULT_TITLE,
    sections: [{ blocks: [createEmptyParagraph()] }],
    defaultStyles: createDefaultStyles(),
  };
}
