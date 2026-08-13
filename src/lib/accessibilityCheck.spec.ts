// SPEC-043: the accessibility checker's detection engine. Framework-free,
// like the module under test — see lib/links.spec.ts for the same
// doc()/paragraph() fixture convention this file follows.
import { describe, expect, it } from "vitest";
import {
  checkAccessibility,
  contrastRatio,
  isLargeText,
  SEVERITY_BY_KIND,
  type AccessibilityIssue,
} from "./accessibilityCheck";
import { createDocument, type Block, type DocumentModel, type Image, type Paragraph, type Shape, type Table } from "./documentModel";

function paragraph(...runs: Array<[string, Record<string, unknown>?]>): Paragraph {
  return {
    type: "paragraph",
    runs: runs.map(([text, formatting]) => ({ text, formatting: formatting ?? {} })),
  };
}

function heading(text: string, level: number, formatting: Record<string, unknown> = {}): Paragraph {
  return { ...paragraph([text, formatting]), headingLevel: level, styleId: `Heading${level}` };
}

function doc(...blocks: Block[]): DocumentModel {
  const model = createDocument("a11y-test");
  model.sections = [{ blocks }];
  return model;
}

function withKind(issues: AccessibilityIssue[], kind: AccessibilityIssue["kind"]): AccessibilityIssue[] {
  return issues.filter((issue) => issue.kind === kind);
}

describe("severity assignment (AC-3)", () => {
  it("matches the spec's Severity reference table exactly", () => {
    expect(SEVERITY_BY_KIND).toEqual({
      "missing-alt-text": "error",
      "table-no-header-row": "error",
      "empty-heading": "error",
      "low-contrast": "error",
      "heading-level-skipped": "warning",
      "link-bare-url": "warning",
      "link-click-here": "warning",
      "colour-alone": "warning",
      "table-merged-cells": "warning",
      "no-heading-structure": "tip",
      "visual-heading": "tip",
    });
  });
});

describe("clean state", () => {
  it("finds nothing in a fresh blank document", () => {
    expect(checkAccessibility(createDocument())).toEqual([]);
  });
});

function image(altText?: string): Image {
  return {
    type: "image",
    mediaId: "m1",
    naturalWidthPx: 100,
    naturalHeightPx: 100,
    widthCm: 2,
    heightCm: 2,
    wrap: "inline",
    altText,
  };
}

describe("images without alt text (AC-4)", () => {
  it("flags an image with no alt text, offering the add-alt-text fix", () => {
    const issues = withKind(checkAccessibility(doc(image(undefined))), "missing-alt-text");
    expect(issues).toHaveLength(1);
    expect(issues[0].severity).toBe("error");
    expect(issues[0].fix).toBe("add-alt-text");
  });

  it("flags whitespace-only alt text the same as missing", () => {
    expect(withKind(checkAccessibility(doc(image("   "))), "missing-alt-text")).toHaveLength(1);
  });

  it("does not flag an image with real alt text", () => {
    expect(withKind(checkAccessibility(doc(image("A chart of quarterly revenue"))), "missing-alt-text")).toHaveLength(0);
  });
});

function shape(altText?: string): Shape {
  return {
    type: "shape",
    geometry: "rect",
    widthCm: 3,
    heightCm: 2,
    wrap: "square",
    altText,
  };
}

// SPEC-051 AC-48: no new detection — Shape carries altText in the same field
// Image does, so checkImageAltText (accessibilityCheck.ts) is reused as-is.
describe("shapes without alt text (SPEC-051 AC-48, no change to SPEC-043's own detection)", () => {
  it("flags a shape with no alt text, offering the add-alt-text fix", () => {
    const issues = withKind(checkAccessibility(doc(shape(undefined))), "missing-alt-text");
    expect(issues).toHaveLength(1);
    expect(issues[0].severity).toBe("error");
    expect(issues[0].fix).toBe("add-alt-text");
  });

  it("flags whitespace-only alt text the same as missing", () => {
    expect(withKind(checkAccessibility(doc(shape("   "))), "missing-alt-text")).toHaveLength(1);
  });

  it("does not flag a shape with real alt text", () => {
    expect(withKind(checkAccessibility(doc(shape("A callout pointing at the total"))), "missing-alt-text")).toHaveLength(0);
  });

  it("an image and a shape missing alt text are both reported, in document order", () => {
    const issues = withKind(checkAccessibility(doc(image(undefined), shape(undefined))), "missing-alt-text");
    expect(issues).toHaveLength(2);
  });
});

describe("tables without a header row (AC-5) and merged cells (AC-6)", () => {
  function table(opts: { headerRow?: boolean; merged?: boolean } = {}): Table {
    const row0 = opts.merged
      ? { cells: [{ blocks: [paragraph(["wide"])], colSpan: 2 }, { blocks: [], covered: true }] }
      : { cells: [{ blocks: [paragraph(["A1"])] }, { blocks: [paragraph(["B1"])] }] };
    return {
      type: "table",
      rows: [{ ...row0, isHeader: opts.headerRow }, { cells: [{ blocks: [paragraph(["A2"])] }, { blocks: [paragraph(["B2"])] }] }],
      columnWidthsCm: [5, 5],
    };
  }

  it("flags a table with no header row, offering the mark-header-row fix", () => {
    const issues = withKind(checkAccessibility(doc(table())), "table-no-header-row");
    expect(issues).toHaveLength(1);
    expect(issues[0].severity).toBe("error");
    expect(issues[0].fix).toBe("mark-header-row");
  });

  it("does not flag a table whose first row is marked as header", () => {
    expect(withKind(checkAccessibility(doc(table({ headerRow: true }))), "table-no-header-row")).toHaveLength(0);
  });

  it("flags a table with merged cells as a Warning with no fix action", () => {
    const issues = withKind(checkAccessibility(doc(table({ merged: true }))), "table-merged-cells");
    expect(issues).toHaveLength(1);
    expect(issues[0].severity).toBe("warning");
    expect(issues[0].fix).toBeUndefined();
  });

  it("does not flag a regular, unmerged table", () => {
    expect(withKind(checkAccessibility(doc(table())), "table-merged-cells")).toHaveLength(0);
  });
});

describe("heading levels skipped (AC-7)", () => {
  it("flags H1 -> H3 with the fix targeting H2", () => {
    const model = doc(heading("Intro", 1), paragraph(["body"]), heading("Details", 3));
    const issues = withKind(checkAccessibility(model), "heading-level-skipped");
    expect(issues).toHaveLength(1);
    expect(issues[0].severity).toBe("warning");
    expect(issues[0].fix).toBe("fix-heading-level");
    expect(issues[0].fixHeadingLevel).toBe(2);
  });

  it("does not flag consecutive levels, including a document's first heading starting below H1", () => {
    const model = doc(heading("Details", 3), heading("Sub", 4));
    expect(withKind(checkAccessibility(model), "heading-level-skipped")).toHaveLength(0);
  });

  it("does not flag a level that goes backwards", () => {
    const model = doc(heading("A", 2), heading("B", 1));
    expect(withKind(checkAccessibility(model), "heading-level-skipped")).toHaveLength(0);
  });
});

describe("empty headings (AC-8)", () => {
  it("flags a heading with no text as an Error", () => {
    const model = doc(heading("", 2));
    const issues = withKind(checkAccessibility(model), "empty-heading");
    expect(issues).toHaveLength(1);
    expect(issues[0].severity).toBe("error");
  });

  it("does not flag a heading with text", () => {
    const model = doc(heading("Overview", 2));
    expect(withKind(checkAccessibility(model), "empty-heading")).toHaveLength(0);
  });
});

describe("link text (AC-9, AC-10)", () => {
  function linkParagraph(text: string): Paragraph {
    return paragraph([text, { linkHref: "https://example.com/report" }]);
  }

  it("flags a bare URL as link text", () => {
    const model = doc(linkParagraph("https://example.com/report"));
    const issues = withKind(checkAccessibility(model), "link-bare-url");
    expect(issues).toHaveLength(1);
    expect(issues[0].severity).toBe("warning");
  });

  it("flags a www.-prefixed bare domain", () => {
    const model = doc(paragraph(["www.example.com", { linkHref: "https://example.com" }]));
    expect(withKind(checkAccessibility(model), "link-bare-url")).toHaveLength(1);
  });

  it('flags "click here" as link text, case-insensitively and with trailing punctuation', () => {
    const model = doc(linkParagraph("Click here."));
    const issues = withKind(checkAccessibility(model), "link-click-here");
    expect(issues).toHaveLength(1);
    expect(issues[0].severity).toBe("warning");
  });

  it("does not flag descriptive link text", () => {
    const model = doc(linkParagraph("Read the full quarterly report"));
    expect(withKind(checkAccessibility(model), "link-bare-url")).toHaveLength(0);
    expect(withKind(checkAccessibility(model), "link-click-here")).toHaveLength(0);
  });

  it("does not flag plain (non-linked) text that merely looks like a URL is absent", () => {
    const model = doc(paragraph(["https://example.com/report"]));
    expect(withKind(checkAccessibility(model), "link-bare-url")).toHaveLength(0);
  });
});

describe("contrast (AC-11, AC-13..16)", () => {
  it("computes the WCAG relative-luminance ratio, rounded to two decimals", () => {
    expect(contrastRatio("#000000", "#ffffff")).toBe(21);
    expect(contrastRatio("#ffffff", "#ffffff")).toBe(1);
    expect(contrastRatio("#000000", "#000000")).toBe(1);
  });

  it("is symmetric regardless of argument order", () => {
    expect(contrastRatio("#828282", "#ffffff")).toBe(contrastRatio("#ffffff", "#828282"));
  });

  it("large text (18pt, or 14pt bold) uses the 3:1 threshold instead of 4.5:1 (AC-15)", () => {
    expect(isLargeText(18, false)).toBe(true);
    expect(isLargeText(17.9, false)).toBe(false);
    expect(isLargeText(14, true)).toBe(true);
    expect(isLargeText(13.9, true)).toBe(false);
    expect(isLargeText(14, false)).toBe(false);
  });

  it("flags low-contrast normal text as an Error", () => {
    const model = doc(paragraph(["hard to read", { color: "#cccccc" }]));
    const issues = withKind(checkAccessibility(model), "low-contrast");
    expect(issues).toHaveLength(1);
    expect(issues[0].severity).toBe("error");
  });

  it("does not flag high-contrast text", () => {
    const model = doc(paragraph(["fine", { color: "#000000" }]));
    expect(withKind(checkAccessibility(model), "low-contrast")).toHaveLength(0);
  });

  it("a ratio that fails the normal threshold but passes the large one is only flagged at normal size", () => {
    // #828282 on white is ~3.84:1 — below 4.5, above 3 — verified directly
    // above before relying on it as a threshold-straddling fixture.
    const ratio = contrastRatio("#828282", "#ffffff");
    expect(ratio).toBeGreaterThan(3);
    expect(ratio).toBeLessThan(4.5);
    const normal = doc(paragraph(["straddle", { color: "#828282", fontSizePt: 12 }]));
    expect(withKind(checkAccessibility(normal), "low-contrast")).toHaveLength(1);
    const large = doc(paragraph(["straddle", { color: "#828282", fontSizePt: 20 }]));
    expect(withKind(checkAccessibility(large), "low-contrast")).toHaveLength(0);
  });

  it("does not flag whitespace-only runs", () => {
    const model = doc(paragraph(["   ", { color: "#ffffff" }]));
    expect(withKind(checkAccessibility(model), "low-contrast")).toHaveLength(0);
  });

  describe("foreground resolution order (AC-17)", () => {
    it("falls back to the resolved named style's colour when the run sets none", () => {
      const model = doc(paragraph(["styled"]));
      model.styles = [{ id: "Custom", name: "Custom", type: "paragraph", runFormatting: { color: "#cccccc" } }];
      model.sections[0].blocks[0] = { ...(model.sections[0].blocks[0] as Paragraph), styleId: "Custom" };
      const issues = withKind(checkAccessibility(model), "low-contrast");
      expect(issues).toHaveLength(1);
    });

    it("falls back to #000000 when neither the run nor its style set a colour", () => {
      const model = doc(paragraph(["plain"]));
      expect(withKind(checkAccessibility(model), "low-contrast")).toHaveLength(0);
    });
  });

  describe("background resolution order (AC-18, AC-19)", () => {
    it("resolves highlight through the HIGHLIGHT_HEX palette ahead of the page default", () => {
      // Yellow highlight behind default-black text is easily readable —
      // the point is that the background used is the highlight, not white.
      const model = doc(paragraph(["marked", { highlight: "yellow" }]));
      expect(withKind(checkAccessibility(model), "low-contrast")).toHaveLength(0);
      // Near-white text on a near-white highlight is the low-contrast case.
      const lowContrast = doc(paragraph(["marked", { color: "#fefefe", highlight: "#ffffff" }]));
      expect(withKind(checkAccessibility(lowContrast), "low-contrast")).toHaveLength(1);
    });

    it("falls back to the run's own shadingFill when highlight is unset", () => {
      const model = doc(paragraph(["shaded", { color: "#fefefe", shadingFill: "#ffffff" }]));
      expect(withKind(checkAccessibility(model), "low-contrast")).toHaveLength(1);
    });

    it("an explicit highlight:false falls through to shadingFill like unset (AC-19)", () => {
      const model = doc(paragraph(["shaded", { color: "#fefefe", highlight: false, shadingFill: "#ffffff" }]));
      expect(withKind(checkAccessibility(model), "low-contrast")).toHaveLength(1);
    });

    it("falls back to the enclosing table cell's shadingFill", () => {
      const table: Table = {
        type: "table",
        rows: [
          {
            cells: [{ blocks: [paragraph(["cell text", { color: "#fefefe" }])], shadingFill: "#ffffff" }],
          },
        ],
        columnWidthsCm: [5],
      };
      expect(withKind(checkAccessibility(doc(table)), "low-contrast")).toHaveLength(1);
    });

    it("falls back to the page's white background when nothing else is set", () => {
      const model = doc(paragraph(["plain", { color: "#fefefe" }]));
      expect(withKind(checkAccessibility(model), "low-contrast")).toHaveLength(1);
    });
  });
});

describe("colour alone (AC-12, AC-23, AC-24)", () => {
  it("flags a run distinguished only by colour", () => {
    const model = doc(paragraph(["Items in "], ["red", { color: "#ff0000" }], [" are mandatory"]));
    const issues = withKind(checkAccessibility(model), "colour-alone");
    expect(issues).toHaveLength(1);
    expect(issues[0].severity).toBe("warning");
  });

  it("does not flag when the run is also bold — bold carries the meaning on its own", () => {
    const model = doc(paragraph(["Items in "], ["red", { color: "#ff0000", bold: true }], [" are mandatory"]));
    expect(withKind(checkAccessibility(model), "colour-alone")).toHaveLength(0);
  });

  it("does not flag when the run differs in size as well as colour", () => {
    const model = doc(paragraph(["Items in "], ["red", { color: "#ff0000", fontSizePt: 20 }], [" are mandatory"]));
    expect(withKind(checkAccessibility(model), "colour-alone")).toHaveLength(0);
  });

  it("flags an unstyled hyperlink with no underline (the carve-out, AC-24)", () => {
    const model = doc(paragraph(["See "], ["the report", { linkHref: "https://example.com" }], [" for details"]));
    expect(withKind(checkAccessibility(model), "colour-alone")).toHaveLength(1);
  });

  it("does not flag a hyperlink that is also underlined", () => {
    const model = doc(
      paragraph(["See "], ["the report", { linkHref: "https://example.com", underline: true }], [" for details"]),
    );
    expect(withKind(checkAccessibility(model), "colour-alone")).toHaveLength(0);
  });

  it("does not flag a single-run paragraph (no other colour to compare against)", () => {
    const model = doc(paragraph(["all one colour", { color: "#ff0000" }]));
    expect(withKind(checkAccessibility(model), "colour-alone")).toHaveLength(0);
  });
});

describe("no heading structure at all (AC-20)", () => {
  it("flags a document with real content and no headings, as a Tip", () => {
    const model = doc(paragraph(["Just some body text."]));
    const issues = withKind(checkAccessibility(model), "no-heading-structure");
    expect(issues).toHaveLength(1);
    expect(issues[0].severity).toBe("tip");
  });

  it("does not flag a document that has at least one heading", () => {
    const model = doc(heading("Title", 1), paragraph(["body"]));
    expect(withKind(checkAccessibility(model), "no-heading-structure")).toHaveLength(0);
  });

  it("does not flag a fresh, still-empty blank document", () => {
    expect(withKind(checkAccessibility(createDocument()), "no-heading-structure")).toHaveLength(0);
  });
});

describe("paragraphs that look like a heading but are not styled as one (AC-21, AC-22)", () => {
  it("flags a short, bold, Normal-style paragraph followed by body text", () => {
    const model = doc(paragraph(["Section overview", { bold: true }]), paragraph(["Some following body text."]));
    const issues = withKind(checkAccessibility(model), "visual-heading");
    expect(issues).toHaveLength(1);
    expect(issues[0].severity).toBe("tip");
  });

  it("flags a short, large paragraph (over 1.2x the default size) followed by body text", () => {
    const model = doc(
      paragraph(["Section overview", { fontSizePt: 14 }]),
      paragraph(["Some following body text."]),
    );
    // Document default is 11pt; 1.2x = 13.2pt, comfortably below 14pt.
    expect(withKind(checkAccessibility(model), "visual-heading")).toHaveLength(1);
  });

  it("does not flag when not followed by a non-empty paragraph", () => {
    const model = doc(paragraph(["Section overview", { bold: true }]), paragraph([""]));
    expect(withKind(checkAccessibility(model), "visual-heading")).toHaveLength(0);
  });

  it("does not flag text over 100 characters", () => {
    const long = "x".repeat(101);
    const model = doc(paragraph([long, { bold: true }]), paragraph(["body"]));
    expect(withKind(checkAccessibility(model), "visual-heading")).toHaveLength(0);
  });

  it("does not flag a paragraph already carrying a heading style", () => {
    const model = doc(heading("Section overview", 2, { bold: true }), paragraph(["body"]));
    expect(withKind(checkAccessibility(model), "visual-heading")).toHaveLength(0);
  });

  it("does not flag a paragraph styled as something other than Normal (e.g. Quote)", () => {
    const styled: Paragraph = { ...paragraph(["Section overview", { bold: true }]), styleId: "Quote" };
    const model = doc(styled, paragraph(["body"]));
    expect(withKind(checkAccessibility(model), "visual-heading")).toHaveLength(0);
  });

  it("does not flag plain prose that is neither bold nor large", () => {
    const model = doc(paragraph(["Just an ordinary short line."]), paragraph(["More text follows."]));
    expect(withKind(checkAccessibility(model), "visual-heading")).toHaveLength(0);
  });
});

describe("issue ordering", () => {
  it("lists issues in document order", () => {
    const model = doc(
      paragraph(["Items in "], ["red", { color: "#ff0000" }], [" only"]),
      paragraph(["www.example.com", { linkHref: "https://example.com" }]),
    );
    const issues = checkAccessibility(model);
    const blocks = issues.map((issue) => issue.at.block);
    expect(blocks).toEqual(blocks.slice().sort((a, b) => a - b));
  });
});
