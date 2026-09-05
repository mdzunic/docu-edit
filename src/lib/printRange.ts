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
  return { pages: Array.from(selected).toSorted((a, b) => a - b) };
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

/** The CSS custom-ident naming one page size, for `@page <name>` and the
 * `page` property that selects it. Derived from the dimensions themselves so
 * the rule and the page box agree without anything having to be threaded
 * between them ("21x29.7cm" -> "docupage-21-29-7"). */
function ruleNamePart(value: number): string {
  return String(value).replace(/[^0-9]/g, "-");
}

export function pageSizeRuleName(pageSize: { widthCm: number; heightCm: number }): string {
  return `docupage-${ruleNamePart(pageSize.widthCm)}-${ruleNamePart(pageSize.heightCm)}`;
}

/** Whether the document's sections disagree about page size — the only case
 * that needs named `@page` rules, and the only case a page box may carry the
 * `page` property. Setting `page` unconditionally is not free: a CHANGE in
 * its value forces a page break in Chromium even when no rule of that name
 * exists, which put an extra blank sheet in every single-size print. */
export function hasMixedPageSizes(model: DocumentModel): boolean {
  return distinctPageSizes(model).length > 1;
}

/** One page box's own CSS geometry. With `named`, `page` selects this size's
 * own `@page` rule (see printPageRuleCss); without it the property is left
 * off entirely. Framework-free so it stays here rather than in the view
 * composable that renders the box. */
export function pageBoxStyle(
  page: { pageSize: { widthCm: number; heightCm: number } },
  named = false,
): Record<string, string> {
  const style: Record<string, string> = {
    width: `${page.pageSize.widthCm}cm`,
    height: `${page.pageSize.heightCm}cm`,
  };
  if (named) {
    style.page = pageSizeRuleName(page.pageSize);
  }
  return style;
}

/** Every distinct page size the document uses, in section order. */
function distinctPageSizes(model: DocumentModel): { widthCm: number; heightCm: number }[] {
  const sizes: { widthCm: number; heightCm: number }[] = [];
  const seen = new Set<string>();
  for (let index = 0; index < model.sections.length; index++) {
    const pageSize = sectionPageSize(model, index);
    const name = pageSizeRuleName(pageSize);
    if (!seen.has(name)) {
      seen.add(name);
      sizes.push(pageSize);
    }
  }
  return sizes;
}

/** The injected print rule: sheets sized to the document's own page size
 * with zero printer margins — the rendered page boxes carry the document
 * margins themselves, so the sheet matches the on-screen page exactly
 * (AC-6, AC-29, AC-30).
 *
 * A document whose sections disagree — a landscape section holding a wide
 * table is the common case — gets one NAMED rule per distinct size, which
 * the page boxes select through the `page` property pageBoxStyle above
 * puts in their style. Sizing every sheet from section 0 printed that landscape
 * section onto portrait paper and cropped it, while the screen showed it
 * correctly. The unnamed rule stays, both as the single-size document's
 * whole answer and as the fallback where named pages are unsupported: such a
 * browser prints what it printed before rather than nothing. */
export function printPageRuleCss(model: DocumentModel): string {
  const sizes = distinctPageSizes(model);
  const first = sizes[0] ?? sectionPageSize(model, 0);
  const rules = [`@page { size: ${first.widthCm}cm ${first.heightCm}cm; margin: 0; }`];
  if (sizes.length > 1) {
    for (const pageSize of sizes) {
      rules.push(
        `@page ${pageSizeRuleName(pageSize)} { size: ${pageSize.widthCm}cm ${pageSize.heightCm}cm; margin: 0; }`,
      );
    }
  }
  return rules.join(" ");
}
