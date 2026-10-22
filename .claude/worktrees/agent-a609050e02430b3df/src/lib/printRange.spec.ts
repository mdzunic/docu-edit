import { describe, expect, it } from "vitest";
import { createDocument } from "./documentModel";
import {
  hasMixedPageSizes,
  isDocumentContentEmpty,
  pageBoxStyle,
  pageSizeRuleName,
  parsePageRange,
  printPageRuleCss,
} from "./printRange";

describe("parsePageRange", () => {
  it("parses single pages, ranges and comma lists, deduplicated and sorted", () => {
    expect(parsePageRange("3", 10).pages).toEqual([3]);
    expect(parsePageRange("2-4", 10).pages).toEqual([2, 3, 4]);
    expect(parsePageRange("5, 1-2, 2", 10).pages).toEqual([1, 2, 5]);
    expect(parsePageRange(" 1 - 3 ", 10).pages).toEqual([1, 2, 3]);
  });

  it("rejects non-numeric input", () => {
    expect(parsePageRange("abc", 10).error).toContain('"abc" is not a page number');
    expect(parsePageRange("1-x", 10).error).toBeTruthy();
    expect(parsePageRange("-2", 10).error).toBeTruthy();
    expect(parsePageRange("1--3", 10).error).toBeTruthy();
  });

  it("rejects empty input", () => {
    expect(parsePageRange("", 10).error).toBeTruthy();
    expect(parsePageRange("   ", 10).error).toBeTruthy();
  });

  it("rejects ranges whose start is after their end", () => {
    expect(parsePageRange("5-2", 10).error).toBe('"5-2" starts after it ends');
  });

  it("rejects pages outside the document", () => {
    expect(parsePageRange("999", 10).error).toBe("Pages go from 1 to 10");
    expect(parsePageRange("0", 10).error).toBeTruthy();
    expect(parsePageRange("1-11", 10).error).toBeTruthy();
    expect(parsePageRange("2", 1).error).toBe("This document only has 1 page");
  });

  it("never returns both pages and an error", () => {
    const ok = parsePageRange("1", 2);
    expect(ok.error).toBeUndefined();
    const bad = parsePageRange("9", 2);
    expect(bad.pages).toBeUndefined();
  });
});

describe("isDocumentContentEmpty", () => {
  it("treats a fresh document (one empty paragraph) as empty", () => {
    expect(isDocumentContentEmpty(createDocument("d1"))).toBe(true);
  });

  it("treats any text as content", () => {
    const model = createDocument("d2");
    model.sections[0].blocks = [{ type: "paragraph", runs: [{ text: "hi", formatting: {} }] }];
    expect(isDocumentContentEmpty(model)).toBe(false);
  });

  it("treats non-paragraph blocks as content", () => {
    const model = createDocument("d3");
    model.sections[0].blocks = [{ type: "break", kind: "page" }];
    expect(isDocumentContentEmpty(model)).toBe(false);
  });

  it("treats a field run as content even with empty literal text", () => {
    const model = createDocument("d4");
    model.sections[0].blocks = [
      { type: "paragraph", runs: [{ text: "", formatting: {}, field: { kind: "page", format: "decimal" } }] },
    ];
    expect(isDocumentContentEmpty(model)).toBe(false);
  });
});

describe("printPageRuleCss", () => {
  it("uses the document default page size (A4) with zero margin", () => {
    expect(printPageRuleCss(createDocument("d5"))).toBe(
      "@page { size: 21cm 29.7cm; margin: 0; }",
    );
  });

  it("uses the first section's own page size when set", () => {
    const model = createDocument("d6");
    model.sections[0].pageSize = { widthCm: 21.59, heightCm: 27.94 };
    expect(printPageRuleCss(model)).toBe("@page { size: 21.59cm 27.94cm; margin: 0; }");
  });

  it("declares no named rule when every section shares one size", () => {
    const model = createDocument("d7");
    model.sections = [
      { blocks: [], pageSize: { widthCm: 21, heightCm: 29.7 } },
      { blocks: [], pageSize: { widthCm: 21, heightCm: 29.7 } },
    ];
    expect(printPageRuleCss(model)).toBe("@page { size: 21cm 29.7cm; margin: 0; }");
  });

  // Sizing every sheet from section 0 printed a landscape section — a wide
  // table's usual home — onto portrait paper and cropped it, while the screen
  // showed it correctly.
  it("gives each distinct size its own named rule when sections disagree", () => {
    const model = createDocument("d8");
    model.sections = [
      { blocks: [], pageSize: { widthCm: 21, heightCm: 29.7 } },
      { blocks: [], pageSize: { widthCm: 29.7, heightCm: 21 } },
      { blocks: [], pageSize: { widthCm: 21, heightCm: 29.7 } },
    ];
    const css = printPageRuleCss(model);
    // The unnamed rule stays: it is what a browser without named-page support
    // falls back to, which is the behaviour this replaces rather than nothing.
    expect(css).toContain("@page { size: 21cm 29.7cm; margin: 0; }");
    expect(css).toContain(`@page ${pageSizeRuleName({ widthCm: 29.7, heightCm: 21 })} { size: 29.7cm 21cm; margin: 0; }`);
    expect(css).toContain(`@page ${pageSizeRuleName({ widthCm: 21, heightCm: 29.7 })} { size: 21cm 29.7cm; margin: 0; }`);
    // One rule per DISTINCT size, not one per section.
    expect(css.split("@page").length - 1).toBe(3);
  });
});

describe("pageSizeRuleName", () => {
  it("is a CSS custom-ident derived from the dimensions", () => {
    expect(pageSizeRuleName({ widthCm: 21, heightCm: 29.7 })).toBe("docupage-21-29-7");
    expect(pageSizeRuleName({ widthCm: 29.7, heightCm: 21 })).toBe("docupage-29-7-21");
  });

  it("separates sizes that differ only by orientation", () => {
    expect(pageSizeRuleName({ widthCm: 21, heightCm: 29.7 })).not.toBe(
      pageSizeRuleName({ widthCm: 29.7, heightCm: 21 }),
    );
  });
});

describe("pageBoxStyle", () => {
  const page = { pageSize: { widthCm: 21, heightCm: 29.7 } };

  // Setting `page` is not free: a CHANGE in its value forces a page break in
  // Chromium even with no rule of that name, which added a blank sheet to
  // every single-size print. So the property appears only where it is needed.
  it("omits the page property unless the document needs named rules", () => {
    expect(pageBoxStyle(page)).toEqual({ width: "21cm", height: "29.7cm" });
  });

  it("names the box when asked, matching its size's rule", () => {
    expect(pageBoxStyle(page, true)).toEqual({
      width: "21cm",
      height: "29.7cm",
      page: pageSizeRuleName(page.pageSize),
    });
  });
});

describe("hasMixedPageSizes", () => {
  it("is false for a document whose sections all agree", () => {
    const model = createDocument("d9");
    model.sections = [
      { blocks: [], pageSize: { widthCm: 21, heightCm: 29.7 } },
      { blocks: [], pageSize: { widthCm: 21, heightCm: 29.7 } },
    ];
    expect(hasMixedPageSizes(model)).toBe(false);
    expect(hasMixedPageSizes(createDocument("d10"))).toBe(false);
  });

  it("is true when a section turns landscape", () => {
    const model = createDocument("d11");
    model.sections = [
      { blocks: [], pageSize: { widthCm: 21, heightCm: 29.7 } },
      { blocks: [], pageSize: { widthCm: 29.7, heightCm: 21 } },
    ];
    expect(hasMixedPageSizes(model)).toBe(true);
  });
});
