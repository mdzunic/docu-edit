// What every library text extractor returns (SPEC-114): the text it could
// read, and how much of the document that is. An extractor never hands back
// an empty string that lies — a document it could not read says `none` and
// why, so the library's "Not searched" list can tell the user. Product-
// neutral: the word processor's documentText, the deck's deckText and
// core/pdfText all return this shape.

/** How much of a document's text an extraction holds: all of it, some of it
 * (the reason says what is missing), or none of it. */
export type TextCoverage = "full" | "partial" | "none";

export interface ExtractedText {
  text: string;
  coverage: TextCoverage;
  /** Why the coverage is not `full`, in words the user reads. */
  reason?: string;
}

/** The most text one document contributes to the index; past it the text
 * is cut and the coverage drops to `partial`. */
export const MAX_INDEXED_CHARS = 2_000_000;

/** The reasons the "Not searched" list shows, one phrase each. */
export const COVERAGE_REASONS = {
  scanned: "scanned image, no text layer",
  noCharMap: "fonts carry no character map",
  sizeCeiling: "over the size ceiling",
  image: "image, no text layer",
  unreadable: "could not be read",
  encrypted: "password-protected",
  importIncomplete: "import not complete",
  spreadsheet: "spreadsheets are not searched",
  indexFull: "index full",
  empty: "no text",
} as const;

/** Cuts an extraction to `max` characters (never splitting a surrogate
 * pair): the text keeps its opening and the coverage becomes `partial`,
 * saying so. */
export function capExtractedText(extracted: ExtractedText, max: number = MAX_INDEXED_CHARS): ExtractedText {
  if (extracted.text.length <= max) {
    return extracted;
  }
  const last = extracted.text.charCodeAt(max - 1);
  const cut = last >= 0xd800 && last <= 0xdbff ? max - 1 : max;
  return {
    text: extracted.text.slice(0, cut),
    coverage: "partial",
    reason: `${COVERAGE_REASONS.sizeCeiling} — only the first ${max.toLocaleString("en-US")} characters are searched`,
  };
}
