// SPEC-029 — page numbering and header/footer-per-page resolution, driven
// by hand-built PageLayout fixtures (pagination itself is untouched by
// this spec: header/footer content never affects where a page breaks).
import { describe, expect, it } from "vitest";
import { createDocument, type DocumentModel, type Section } from "./documentModel";
import type { PageLayout } from "./pagination";
import {
  computePageNumbers,
  fieldValueText,
  headerFooterForPage,
  isFirstPageOfSection,
  resolveHeaderFooterKind,
} from "./pageNumbering";

function page(sectionIndex: number): PageLayout {
  return { sectionIndex, pageSize: { widthCm: 21, heightCm: 29.7 }, margins: { topCm: 2, rightCm: 2, bottomCm: 2, leftCm: 2 }, items: [] };
}

function docWithSections(sections: Section[]): DocumentModel {
  const model = createDocument("page-numbering-test");
  model.sections = sections;
  return model;
}

describe("computePageNumbers (SPEC-029 AC-20..23)", () => {
  it("counts pages from 1 by default", () => {
    const pages = [page(0), page(0), page(0)];
    const model = docWithSections([{ blocks: [] }]);
    expect(computePageNumbers(pages, model)).toEqual([1, 2, 3]);
  });

  it("continues numbering across a section boundary when the next section has no start (AC-23)", () => {
    const pages = [page(0), page(0), page(1), page(1)];
    const model = docWithSections([{ blocks: [] }, { blocks: [] }]);
    expect(computePageNumbers(pages, model)).toEqual([1, 2, 3, 4]);
  });

  it("restarts numbering at a section's pageNumberStart (AC-22)", () => {
    const pages = [page(0), page(0), page(1), page(1)];
    const model = docWithSections([{ blocks: [] }, { blocks: [], pageNumberStart: 1 }]);
    expect(computePageNumbers(pages, model)).toEqual([1, 2, 1, 2]);
  });

  it("restarts at an arbitrary start value", () => {
    const pages = [page(0), page(1)];
    const model = docWithSections([{ blocks: [] }, { blocks: [], pageNumberStart: 100 }]);
    expect(computePageNumbers(pages, model)).toEqual([1, 100]);
  });
});

describe("isFirstPageOfSection (SPEC-029 AC-11, AC-27)", () => {
  it("is true for page 0 and for the first page after a section change", () => {
    const pages = [page(0), page(0), page(1)];
    expect(isFirstPageOfSection(pages, 0)).toBe(true);
    expect(isFirstPageOfSection(pages, 1)).toBe(false);
    expect(isFirstPageOfSection(pages, 2)).toBe(true);
  });
});

describe("resolveHeaderFooterKind (SPEC-029 AC-11..13)", () => {
  it("resolves 'default' when neither option is set", () => {
    const pages = [page(0), page(0)];
    const model = docWithSections([{ blocks: [] }]);
    expect(resolveHeaderFooterKind(model, pages, 0, 1)).toBe("default");
    expect(resolveHeaderFooterKind(model, pages, 1, 2)).toBe("default");
  });

  it("resolves 'first' on the section's first page when titlePage is set", () => {
    const pages = [page(0), page(0)];
    const model = docWithSections([{ blocks: [], titlePage: true }]);
    expect(resolveHeaderFooterKind(model, pages, 0, 1)).toBe("first");
    expect(resolveHeaderFooterKind(model, pages, 1, 2)).toBe("default");
  });

  it("resolves 'even' on even displayed page numbers when evenAndOddHeaders is set", () => {
    const pages = [page(0), page(0), page(0)];
    const model = docWithSections([{ blocks: [], evenAndOddHeaders: true }]);
    expect(resolveHeaderFooterKind(model, pages, 0, 1)).toBe("default");
    expect(resolveHeaderFooterKind(model, pages, 1, 2)).toBe("even");
    expect(resolveHeaderFooterKind(model, pages, 2, 3)).toBe("default");
  });

  it("prefers 'first' over 'even' on a title page whose displayed number happens to be even", () => {
    const pages = [page(0)];
    const model = docWithSections([
      { blocks: [], titlePage: true, evenAndOddHeaders: true, pageNumberStart: 2 },
    ]);
    expect(resolveHeaderFooterKind(model, pages, 0, 2)).toBe("first");
  });
});

describe("headerFooterForPage (SPEC-029 AC-9)", () => {
  it("returns undefined (not the default content) when the resolved kind has no content of its own", () => {
    const pages = [page(0)];
    const model = docWithSections([
      { blocks: [], titlePage: true, headers: { default: { blocks: [] } } },
    ]);
    expect(headerFooterForPage(model, pages, 0, 1, "header")).toBeUndefined();
  });

  it("returns the resolved kind's content when present", () => {
    const pages = [page(0)];
    const model = docWithSections([
      { blocks: [], titlePage: true, headers: { first: { blocks: [] }, default: { blocks: [] } } },
    ]);
    expect(headerFooterForPage(model, pages, 0, 1, "header")).toBe(model.sections[0].headers?.first);
  });
});

describe("fieldValueText (SPEC-029 AC-16..18, AC-20, AC-21)", () => {
  it("formats the page field as the page's own number", () => {
    expect(fieldValueText("page", "decimal", 3, 10)).toBe("3");
    expect(fieldValueText("page", "lowerRoman", 3, 10)).toBe("iii");
    expect(fieldValueText("page", "lowerAlpha", 3, 10)).toBe("c");
  });

  it("formats the numpages field as the total page count", () => {
    expect(fieldValueText("numpages", "decimal", 3, 10)).toBe("10");
  });
});
