// Print-domain helpers for SPEC-034: parsing the preview's page-range
// input, deciding whether a document has any printable content at all, and
// building the injected @page rule that makes the browser's print dialog
// use the document's own page geometry. Framework-free per src/lib.
import {
  isBreak,
  isChart,
  isConnector,
  isImage,
  isShape,
  isTable,
  isToc,
  isTof,
  sectionPageSize,
  type DocumentModel,
} from "./documentModel";

/** A parsed page range: the 1-based page numbers to print, ascending, or a
 * user-facing validation error. Exactly one of the two is set. */
export interface PageRangeResult {
  pages?: number[];
  error?: string;
}

export const RANGE_HINT = 'Use page numbers like "1-3, 5"';

/** Parses a custom page-range string ("2", "1-3", "1-2, 5") against the
 * document's page count. Errors on non-numeric tokens, ranges whose start
 * is after their end, and pages outside 1..totalPages (AC-28). */
export function parsePageRange(input: string, totalPages: number): PageRangeResult {
  const trimmed = input.trim();
  if (trimmed.length === 0) {
    return { error: `Enter a page range. ${RANGE_HINT}` };
  }
  const selected = new Set<number>();
  for (const token of trimmed.split(",")) {
    const part = token.trim();
    const match = /^(\d+)(?:\s*-\s*(\d+))?$/.exec(part);
    if (!match) {
      return { error: `"${part}" is not a page number. ${RANGE_HINT}` };
    }
    const start = Number.parseInt(match[1], 10);
    const end = match[2] === undefined ? start : Number.parseInt(match[2], 10);
    if (end < start) {
      return { error: `"${part}" starts after it ends` };
    }
    if (start < 1 || end > totalPages) {
      return {
        error:
          totalPages === 1
            ? "This document only has 1 page"
            : `Pages go from 1 to ${totalPages}`,
      };
    }
    for (let page = start; page <= end; page += 1) {
      selected.add(page);
    }
  }
  return { pages: Array.from(selected).sort((a, b) => a - b) };
}

/** Whether the document has no printable content: every section's block
 * list is empty, or holds only paragraphs with no text (an untouched new
 * document). Such a document previews as no pages and exports a zero-page
 * PDF (AC-27); any table, image, break, TOC, field or character counts as
 * content. */
export function isDocumentContentEmpty(model: DocumentModel): boolean {
  for (const section of model.sections) {
    for (const block of section.blocks) {
      if (isTable(block) || isImage(block) || isShape(block) || isChart(block) || isConnector(block) || isBreak(block) || isToc(block) || isTof(block)) {
        return false;
      }
      for (const run of block.runs) {
        if (run.field !== undefined || run.text.length > 0) {
          return false;
        }
      }
    }
  }
  return true;
}

/** The injected print rule: sheets sized to the document's own page size
 * (its first section's, or the A4 default) with zero printer margins — the
 * rendered page boxes carry the document margins themselves, so the sheet
 * matches the on-screen page exactly (AC-6, AC-29, AC-30). */
export function printPageRuleCss(model: DocumentModel): string {
  const pageSize = sectionPageSize(model, 0);
  return `@page { size: ${pageSize.widthCm}cm ${pageSize.heightCm}cm; margin: 0; }`;
}
