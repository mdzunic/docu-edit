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
}

/** The smallest unit of text: a string plus the formatting applied to it. */
export interface Run {
  text: string;
  formatting: RunFormatting;
}

/** A paragraph is a block-level element made of inline runs. */
export interface Paragraph {
  type: "paragraph";
  runs: Run[];
}

/** Block-level children of a section. Paragraph is the only kind for now;
 * later specs (tables, images) extend this union. */
export type Block = Paragraph;

/** A section groups block-level content; documents have one or more. */
export interface Section {
  blocks: Block[];
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

/** The root of the document tree. */
export interface DocumentModel {
  id: string;
  title: string;
  sections: Section[];
  defaultStyles: DefaultStyles;
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

function createDefaultStyles(): DefaultStyles {
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
