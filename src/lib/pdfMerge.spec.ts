// SPEC-111: the page organiser's framework-free rules — the "What changed"
// grouping, a page card's rotation and aspect, the source-contiguous runs
// and their file names, report keys, and the estimate line.
import { describe, expect, it } from "vitest";
import {
  cardAspect,
  dropRowText,
  groupDrops,
  mergeEstimateLine,
  rotationLabel,
  runFileNames,
  sourceRuns,
  turned,
  uniqueSourceKeys,
  zipFileName,
} from "./pdfMerge";
import type { PdfDropRecord } from "./core/pdfDocument";

describe("groupDrops", () => {
  it("groups records by kind in the union's order, keeping each group's report order and leaving empty kinds out", () => {
    const records: PdfDropRecord[] = [
      { kind: "annotation", source: "a", page: 2, reason: "first link" },
      { kind: "structureTree", source: "b", reason: "tags" },
      { kind: "annotation", source: "a", page: 3, reason: "second link" },
      { kind: "acroForm", source: "c", reason: "form" },
    ];
    expect(groupDrops(records).map((group) => [group.kind, group.label, group.records.map((record) => record.reason)])).toEqual([
      ["acroForm", "Form fields", ["form"]],
      ["structureTree", "Structure tree (accessibility tags)", ["tags"]],
      ["annotation", "Links and annotations", ["first link", "second link"]],
    ]);
    expect(groupDrops([])).toEqual([]);
  });

  it("writes one row per record: the source, its page where it has one, and the reason", () => {
    expect(dropRowText({ kind: "annotation", source: "contract", page: 4, reason: "A link was removed." })).toBe(
      "“contract”, page 4: A link was removed.",
    );
    expect(dropRowText({ kind: "pageLabel", source: "scan", reason: "Labels were removed." })).toBe("“scan”: Labels were removed.");
  });
});

describe("a page card's rotation and shape", () => {
  it("adds the user's turns to the page's own rotation, modulo 360", () => {
    expect(turned(0, 90)).toBe(90);
    expect(turned(270, 90)).toBe(0);
    expect(turned(0, -90)).toBe(270);
    expect(rotationLabel({ rotate: 90, rotateBy: 180 })).toBe("rotate 270");
    expect(rotationLabel({ rotate: 270, rotateBy: 90 })).toBe("rotate 0");
  });

  it("takes the page's real aspect, its sides swapped by a quarter turn", () => {
    expect(cardAspect({ widthPt: 612, heightPt: 792, rotate: 0, rotateBy: 0 })).toBeCloseTo(612 / 792);
    expect(cardAspect({ widthPt: 612, heightPt: 792, rotate: 90, rotateBy: 0 })).toBeCloseTo(792 / 612);
    expect(cardAspect({ widthPt: 612, heightPt: 792, rotate: 90, rotateBy: 90 })).toBeCloseTo(612 / 792);
    expect(cardAspect({ widthPt: 0, heightPt: 792, rotate: 0, rotateBy: 0 })).toBe(1);
  });
});

describe("separate PDFs", () => {
  it("cuts the plan into runs of consecutive pages from one source", () => {
    expect(sourceRuns(["a", "a", "b", "a", "a", "c"], (key) => key)).toEqual([["a", "a"], ["b"], ["a", "a"], ["c"]]);
    expect(sourceRuns([], (key: string) => key)).toEqual([]);
  });

  it("names each run's file in plan order, safe for a zip", () => {
    expect(runFileNames(["contract", "photo", "contract"])).toEqual(["01 contract.pdf", "02 photo.pdf", "03 contract.pdf"]);
    expect(runFileNames(["a/b\\c", " "])).toEqual(["01 a-b-c.pdf", "02 Untitled.pdf"]);
    expect(runFileNames(Array.from({ length: 120 }, () => "x"))[0]).toBe("001 x.pdf");
    expect(zipFileName(" Merged ")).toBe("Merged.zip");
    expect(zipFileName("")).toBe("Combined PDF.zip");
  });
});

describe("report keys and the estimate", () => {
  it("numbers sources that share a title, so a record names one file", () => {
    expect(uniqueSourceKeys(["scan", "photo", "scan", "scan (2)", ""])).toEqual(["scan", "photo", "scan (2)", "scan (2) (2)", "Untitled"]);
  });

  it("states what each destination will write", () => {
    const counts = { pages: 5, selected: 2, sources: 2, runs: 3 };
    expect(mergeEstimateLine("merge", counts)).toBe("5 pages from 2 files");
    expect(mergeEstimateLine("extract", counts)).toBe("2 pages of 5 selected");
    expect(mergeEstimateLine("split", counts)).toBe("3 PDFs, one per run of pages from the same file · downloads as a .zip");
    expect(mergeEstimateLine("merge", { pages: 1, selected: 0, sources: 1, runs: 1 })).toBe("1 page from 1 file");
  });
});
