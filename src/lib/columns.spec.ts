// SPEC-030 — the column layout commands: normalisation, applying to a
// section, wrapping a selection in a continuous section, and ruler boundary
// drag.
import { describe, expect, it } from "vitest";
import {
  applyColumns,
  normalizeColumns,
  sectionColumnsState,
  setColumnBoundary,
} from "./columns";
import { caretSelection } from "./commands";
import {
  createDocument,
  resolveColumnWidthsCm,
  sectionContentWidthCm,
  type DocumentModel,
  type Paragraph,
  type SectionColumns,
} from "./documentModel";
import type { CaretPosition } from "./textEditing";

function paragraph(text: string): Paragraph {
  return { type: "paragraph", runs: text ? [{ text, formatting: {} }] : [] };
}

function caret(block: number, offset: number, section = 0): CaretPosition {
  return { section, block, offset };
}

function fixture(...texts: string[]): DocumentModel {
  const model = createDocument("columns-fixture");
  model.sections = [{ blocks: texts.map(paragraph) }];
  return model;
}

const twoColumns: SectionColumns = { count: 2, equalWidth: true, gapCm: 1.27, separator: false };

describe("normalizeColumns", () => {
  it("clamps the count to 1..6 and drops columns at count 1", () => {
    expect(normalizeColumns({ ...twoColumns, count: 1 })).toBeUndefined();
    expect(normalizeColumns({ ...twoColumns, count: 9 })?.count).toBe(6);
    expect(normalizeColumns({ ...twoColumns, count: 3 })?.count).toBe(3);
  });

  it("keeps unequal widths only when the array length matches the count", () => {
    const kept = normalizeColumns({ count: 2, equalWidth: false, gapCm: 1, separator: true, widthsCm: [5, 8] });
    expect(kept?.equalWidth).toBe(false);
    expect(kept?.widthsCm).toHaveLength(2);
    expect(kept?.separator).toBe(true);
    const mismatched = normalizeColumns({ count: 3, equalWidth: false, gapCm: 1, separator: false, widthsCm: [5, 8] });
    expect(mismatched?.equalWidth).toBe(true);
    expect(mismatched?.widthsCm).toBeUndefined();
  });
});

describe("applyColumns — this section (AC-6)", () => {
  it("sets the caret's section columns and clears preserved sectPr", () => {
    const model = fixture("a", "b");
    model.sections[0].docxSectPr = "<w:sectPr/>";
    const { model: next } = applyColumns(model, caretSelection(caret(0, 0)), twoColumns, "thisSection");
    expect(next.sections).toHaveLength(1);
    expect(next.sections[0].columns?.count).toBe(2);
    expect(next.sections[0].docxSectPr).toBeUndefined();
  });

  it("removes columns when the count is 1", () => {
    const model = fixture("a");
    model.sections[0].columns = { ...twoColumns };
    const { model: next } = applyColumns(
      model,
      caretSelection(caret(0, 0)),
      { ...twoColumns, count: 1 },
      "thisSection",
    );
    expect(next.sections[0].columns).toBeUndefined();
  });
});

describe("applyColumns — selection (AC-7/AC-8)", () => {
  it("wraps the touched paragraphs in a new continuous section (AC-7a)", () => {
    const model = fixture("first", "second", "third", "fourth");
    // Select paragraphs 1 and 2 (partial offsets — must expand to whole paras).
    const selection = { anchor: caret(1, 2), focus: caret(2, 3) };
    const { model: next } = applyColumns(model, selection, twoColumns, "selection");
    // before(0) | middle(1,2 continuous+columns) | after(3 continuous)
    expect(next.sections).toHaveLength(3);
    expect(next.sections[0].blocks.map((b) => (b as Paragraph).runs[0]?.text)).toEqual(["first"]);
    expect(next.sections[1].start).toBe("continuous");
    expect(next.sections[1].columns?.count).toBe(2);
    expect(next.sections[1].blocks.map((b) => (b as Paragraph).runs[0]?.text)).toEqual([
      "second",
      "third",
    ]);
    expect(next.sections[2].start).toBe("continuous");
    expect(next.sections[2].blocks.map((b) => (b as Paragraph).runs[0]?.text)).toEqual(["fourth"]);
  });

  it("does not split any paragraph mid-text (AC-7b)", () => {
    const model = fixture("alpha", "beta");
    const before = model.sections[0].blocks.length;
    const selection = { anchor: caret(0, 2), focus: caret(0, 4) };
    const { model: next } = applyColumns(model, selection, twoColumns, "selection");
    const total = next.sections.reduce((n, s) => n + s.blocks.length, 0);
    expect(total).toBe(before); // no paragraph added or split
  });

  it("wraps just the caret's paragraph for an empty selection (AC-8)", () => {
    const model = fixture("one", "two", "three");
    const { model: next } = applyColumns(model, caretSelection(caret(1, 1)), twoColumns, "selection");
    const wrapped = next.sections.find((s) => s.columns?.count === 2);
    expect(wrapped?.blocks.map((b) => (b as Paragraph).runs[0]?.text)).toEqual(["two"]);
  });

  it("wrapping the whole section keeps it a single section", () => {
    const model = fixture("only");
    const { model: next } = applyColumns(model, caretSelection(caret(0, 0)), twoColumns, "selection");
    expect(next.sections).toHaveLength(1);
    expect(next.sections[0].columns?.count).toBe(2);
  });
});

describe("setColumnBoundary (AC-14)", () => {
  it("resizes the two adjacent columns, preserving the total", () => {
    const model = fixture("a");
    model.sections[0].columns = { ...twoColumns };
    const contentWidth = sectionContentWidthCm(model, 0);
    const { model: next } = setColumnBoundary(model, caretSelection(caret(0, 0)), 0, 1.5);
    const cols = next.sections[0].columns!;
    expect(cols.equalWidth).toBe(false);
    const widths = resolveColumnWidthsCm(cols, contentWidth);
    expect(widths[0]).toBeGreaterThan(widths[1]);
    // Total content width (widths + gutter) is preserved.
    expect(widths[0] + widths[1] + cols.gapCm).toBeCloseTo(contentWidth, 1);
  });

  it("is a no-op in a single-column section", () => {
    const model = fixture("a");
    const { model: next } = setColumnBoundary(model, caretSelection(caret(0, 0)), 0, 1);
    expect(next).toBe(model);
  });
});

describe("sectionColumnsState", () => {
  it("returns a single column for a section with no columns", () => {
    const model = fixture("a");
    expect(sectionColumnsState(model, caretSelection(caret(0, 0))).count).toBe(1);
  });
});
