// Page numbering and header/footer resolution (SPEC-029): pure functions
// over an already-computed pagination layout, framework-free like the rest
// of lib/. Deliberately separate from pagination.ts — page numbers and
// which header/footer kind applies to a page are render-time concerns, not
// layout ones (header/footer content never affects where a page breaks).
import {
  sectionHeaderFooter,
  type DocumentModel,
  type HeaderFooterContent,
  type HeaderFooterKind,
  type HeaderFooterPart,
  type PageNumberFormat,
} from "./documentModel";
import { formatOrdinal } from "./lists";
import type { PageLayout } from "./pagination";

/** The displayed page number of every laid-out page, in order (SPEC-029
 * AC-20..23): counting starts at 1 and continues across section
 * boundaries unless a section sets `pageNumberStart`, which restarts the
 * count from that value on the section's first page. */
export function computePageNumbers(pages: PageLayout[], model: DocumentModel): number[] {
  const numbers: number[] = [];
  let counter = 1;
  let lastSection = -1;
  for (const page of pages) {
    if (page.sectionIndex !== lastSection) {
      const start = model.sections[page.sectionIndex]?.pageNumberStart;
      if (start !== undefined) {
        counter = start;
      }
      lastSection = page.sectionIndex;
    }
    numbers.push(counter);
    counter += 1;
  }
  return numbers;
}

/** True for the first laid-out page of its section (SPEC-029 AC-11,
 * AC-27) — pages are produced in section order, so this is just "the
 * section changed since the previous page". */
export function isFirstPageOfSection(pages: PageLayout[], pageIndex: number): boolean {
  return pageIndex === 0 || pages[pageIndex - 1]?.sectionIndex !== pages[pageIndex]?.sectionIndex;
}

/** Which header/footer kind applies to a given laid-out page (SPEC-029
 * AC-11..13): the section's first page uses `"first"` when `titlePage` is
 * set; otherwise an even *displayed* page number uses `"even"` when
 * `evenAndOddHeaders` is set; everything else uses `"default"`. */
export function resolveHeaderFooterKind(
  model: DocumentModel,
  pages: PageLayout[],
  pageIndex: number,
  pageNumber: number,
): HeaderFooterKind {
  const section = model.sections[pages[pageIndex]?.sectionIndex ?? -1];
  if (!section) {
    return "default";
  }
  if (section.titlePage && isFirstPageOfSection(pages, pageIndex)) {
    return "first";
  }
  if (section.evenAndOddHeaders && pageNumber % 2 === 0) {
    return "even";
  }
  return "default";
}

/** The header or footer content that actually renders on a given page
 * (SPEC-029 AC-9): resolves the applicable kind and returns undefined —
 * not the `"default"` content — when that kind has no content of its own,
 * matching Word's "different first page starts blank" behaviour. */
export function headerFooterForPage(
  model: DocumentModel,
  pages: PageLayout[],
  pageIndex: number,
  pageNumber: number,
  part: HeaderFooterPart,
): HeaderFooterContent | undefined {
  const sectionIndex = pages[pageIndex]?.sectionIndex;
  if (sectionIndex === undefined) {
    return undefined;
  }
  const kind = resolveHeaderFooterKind(model, pages, pageIndex, pageNumber);
  return sectionHeaderFooter(model, sectionIndex, part, kind);
}

/** A field's live display value (SPEC-029 AC-20, AC-21): `"page"` is the
 * page's own displayed number, `"numpages"` the document's total page
 * count, both formatted per the field's chosen numbering style. */
export function fieldValueText(
  kind: "page" | "numpages",
  format: PageNumberFormat,
  pageNumber: number,
  totalPages: number,
): string {
  return formatOrdinal(format, kind === "page" ? pageNumber : totalPages);
}
