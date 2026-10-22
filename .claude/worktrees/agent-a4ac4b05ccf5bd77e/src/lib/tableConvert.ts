// Text ↔ table conversion (SPEC-076): the two pure transforms behind
// `Table → Convert text to table…` and `Table → Convert table to text…`.
// Framework-free by contract (AC-78) — no Vue import, no DOM.
//
// The pair is designed to round-trip: for the simple case (every paragraph
// has the same field count, no field's text contains the separator, no
// paragraph is empty) `tableToText(textToTable(p, s), s)` is
// `equalStructural` to `p` (AC-65). That is why `textToTable` splits runs
// rather than flattening to plain strings — the field keeps the run
// formatting of the text it came from (AC-58), so the trip back rebuilds the
// same runs.
import { cmToTwips, twipsToCm } from "./core/units";
import { normalizeRuns } from "./textEditing";
import type { Paragraph, Run, Table, TableCell, TableRow } from "./documentModel";

/** The separators the dialogs offer (AC-53, AC-61); "custom" takes a single
 * character of the user's choosing. */
export type ConvertSeparatorId = "tab" | "comma" | "semicolon" | "custom";

/** The literal character each named separator stands for. */
export const CONVERT_SEPARATORS: { id: ConvertSeparatorId; label: string; value: string }[] = [
  { id: "tab", label: "Tab", value: "\t" },
  { id: "comma", label: "Comma", value: "," },
  { id: "semicolon", label: "Semicolon", value: ";" },
  { id: "custom", label: "Custom", value: "" },
];

/** The character a separator choice resolves to; `custom` uses the typed
 * character. Returns "" for an unusable custom entry — the dialog's own
 * validation (AC-55) is what stops that ever reaching a conversion. */
export function separatorCharacter(id: ConvertSeparatorId, custom: string): string {
  const entry = CONVERT_SEPARATORS.find((option) => option.id === id);
  return id === "custom" ? custom : (entry?.value ?? "\t");
}

/** Whether a custom separator is usable (AC-55): exactly one character, and
 * not whitespace. Empty, whitespace-only and longer-than-one all fail. */
export function isValidCustomSeparator(custom: string): boolean {
  return custom.length === 1 && custom.trim().length === 1;
}

/** Typed pass-through for a separator id arriving from a `.vue` template's
 * plain string (espree bars the annotation there). */
export function asConvertSeparatorId(value: string): ConvertSeparatorId {
  return value as ConvertSeparatorId;
}

/** One paragraph's runs split at every occurrence of `separator`, each field
 * keeping the run formatting of the text it came from (AC-58). A run
 * containing the separator splits into as many runs as it has fields, all
 * sharing that run's formatting; a field spanning two runs keeps both.
 * Empty trailing fields survive as empty run lists (AC-57). */
function splitParagraphRuns(paragraph: Paragraph, separator: string): Run[][] {
  const fields: Run[][] = [[]];
  for (const run of paragraph.runs) {
    // A field run (a merge chip, a SPEC-053 field) is a single caret atom —
    // never split it, whatever its text happens to contain.
    if (run.field !== undefined || !run.text.includes(separator)) {
      fields[fields.length - 1].push(run);
      continue;
    }
    const pieces = run.text.split(separator);
    pieces.forEach((piece, index) => {
      if (index > 0) {
        fields.push([]);
      }
      if (piece.length > 0) {
        fields[fields.length - 1].push({ ...run, text: piece });
      }
    });
  }
  return fields;
}

/** Equal column widths spanning `contentWidthCm`, snapped to the twip grid —
 * the same rule lib/tables.ts's equalWidthsAcross applies, restated here so
 * this module stays free of the model-shaped helpers. */
function equalWidths(contentWidthCm: number, columns: number): number[] {
  const width = twipsToCm(cmToTwips(Math.max(0.1, contentWidthCm) / Math.max(1, columns)));
  return Array.from({ length: columns }, () => width);
}

/** The row and column count a conversion would produce — the dialog's live
 * preview (AC-54), computed without building the table. */
export function textToTableSize(
  paragraphs: Paragraph[],
  separator: string,
): { rows: number; columns: number } {
  if (paragraphs.length === 0 || separator.length === 0) {
    return { rows: 0, columns: 0 };
  }
  let columns = 1;
  for (const paragraph of paragraphs) {
    columns = Math.max(columns, splitParagraphRuns(paragraph, separator).length);
  }
  return { rows: paragraphs.length, columns };
}

/** One paragraph per row, one field per cell (AC-56..AC-59).
 *
 * The table gets as many columns as the paragraph with the most fields;
 * shorter rows are padded with empty cells. A paragraph ending in the
 * separator yields a trailing *empty cell*, not one column fewer. Column
 * widths divide `contentWidthCm` equally — the caller passes the section's
 * content width (`sectionContentWidthCm`); the default is A4's text area at
 * the product's default 2.54 cm margins, for callers that have no section.
 */
export function textToTable(
  paragraphs: Paragraph[],
  separator: string,
  contentWidthCm = 21 - 2.54 * 2,
): Table {
  const rowFields = paragraphs.map((paragraph) => splitParagraphRuns(paragraph, separator));
  const columns = Math.max(1, ...rowFields.map((fields) => fields.length));
  const rows: TableRow[] = rowFields.map((fields, rowIndex) => {
    const source = paragraphs[rowIndex];
    const cells: TableCell[] = Array.from({ length: columns }, (_unused, colIndex) => {
      const runs = normalizeRuns(fields[colIndex] ?? []);
      // Each cell paragraph keeps the source paragraph's own formatting and
      // style so a converted list of headings still reads as headings.
      const cellParagraph: Paragraph = { type: "paragraph", runs };
      if (source.formatting !== undefined) {
        cellParagraph.formatting = source.formatting;
      }
      if (source.styleId !== undefined) {
        cellParagraph.styleId = source.styleId;
      }
      if (source.headingLevel !== undefined) {
        cellParagraph.headingLevel = source.headingLevel;
      }
      return { blocks: [cellParagraph] };
    });
    return { cells };
  });
  return { type: "table", rows, columnWidthsCm: equalWidths(contentWidthCm, columns) };
}

/** A cell's text contribution: its paragraphs' run text, with a nested
 * table's own flattened text in document order (AC-63). */
function cellText(cell: TableCell, separator: string): string {
  let text = "";
  for (const block of cell.blocks) {
    text += flattenBlockText(block, separator);
  }
  return text;
}

/** A block's text for the conversion — a paragraph's runs, or a nested
 * table flattened by the same rule the conversion itself uses (AC-63): its
 * rows joined by newlines, its cells by the separator. */
function flattenBlockText(block: Paragraph | Table, separator: string): string {
  if ((block as Table).type === "table") {
    const nested = block as Table;
    return nested.rows
      .map((row) => row.cells.filter((cell) => !cell.covered).map((cell) => cellText(cell, separator)).join(separator))
      .join("\n");
  }
  return (block as Paragraph).runs.map((run) => run.text).join("");
}

/** The formatting a joining separator run takes: that of the run it follows,
 * so the trip back through `normalizeRuns` rebuilds the source paragraph's
 * own run boundaries instead of fragmenting it at every separator (AC-65).
 *
 * Which side a separator belongs to is genuinely ambiguous once it has been
 * split out — `"a" + sep` ending one run and `sep + "b"` opening the next are
 * indistinguishable in the table. The convention here is the one typing
 * produces: a separator carries the formatting in force when it was typed,
 * i.e. the preceding run's. The following cell's first run is the fallback
 * for a row opening with an empty cell. */
function separatorFormatting(next: TableCell | undefined, previousRun: Run | undefined): Run["formatting"] {
  const ahead = next?.blocks[0]?.runs.find((run) => run.field === undefined);
  return previousRun?.formatting ?? ahead?.formatting ?? {};
}

/** One paragraph per row, cells joined by `separator` (AC-62).
 *
 * The result's formatting comes from each cell's *first* paragraph — that is
 * the paragraph a round trip has to reproduce — and any further paragraphs
 * in the cell contribute their text inline, appended to the first
 * paragraph's last run rather than starting a new one. Covered cells (the
 * placeholder slots of a merge) contribute nothing.
 */
export function tableToText(table: Table, separator: string): Paragraph[] {
  return table.rows.map((row) => {
    const runs: Run[] = [];
    let paragraphSource: Paragraph | undefined;
    const cells = row.cells.filter((cell) => !cell.covered);
    cells.forEach((cell, index) => {
      if (index > 0) {
        runs.push({ text: separator, formatting: separatorFormatting(cell, runs[runs.length - 1]) });
      }
      const first = cell.blocks[0];
      if (first !== undefined && paragraphSource === undefined) {
        paragraphSource = first;
      }
      if (first !== undefined) {
        for (const run of first.runs) {
          runs.push({ ...run });
        }
      }
      // Further paragraphs (and any nested table) join the same output
      // paragraph inline, carrying their text but not their own block.
      for (const block of cell.blocks.slice(1)) {
        const text = flattenBlockText(block, separator);
        if (text.length > 0) {
          runs.push({ text, formatting: {} });
        }
      }
    });
    const paragraph: Paragraph = { type: "paragraph", runs: normalizeRuns(runs) };
    if (paragraphSource?.formatting !== undefined) {
      paragraph.formatting = paragraphSource.formatting;
    }
    if (paragraphSource?.styleId !== undefined) {
      paragraph.styleId = paragraphSource.styleId;
    }
    if (paragraphSource?.headingLevel !== undefined) {
      paragraph.headingLevel = paragraphSource.headingLevel;
    }
    return paragraph;
  });
}
