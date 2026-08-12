import { describe, expect, it } from "vitest";
import { createDocument, isMergeChipField, type Paragraph } from "./documentModel";
import { equalStructural } from "./equality";
import {
  chipLabel,
  createEachChipRun,
  createEndifChipRun,
  createIfChipRun,
  createMergeChipRun,
  findLiteralChipCandidates,
  isValidChipName,
  parseChipText,
} from "./mergeChips";
import { writeDocx } from "./docx/writer";
import { parseDocx } from "./docx/reader";
import { writeOdt } from "./odf/writer";
import { parseOdt } from "./odf/reader";

function docWithOneOfEachChip(formatting: Record<string, unknown> = {}) {
  const model = createDocument("chips");
  const runs = [
    createMergeChipRun("client_name", formatting),
    createIfChipRun("paid", formatting),
    createEndifChipRun(formatting),
    createEachChipRun("line_items", formatting),
  ];
  model.sections[0].blocks = runs.map((run) => ({ type: "paragraph", runs: [run] }) as Paragraph);
  return model;
}

describe("chipLabel / parseChipText", () => {
  it("round-trips every chip kind through its canonical text", () => {
    expect(chipLabel({ kind: "merge", name: "client_name" })).toBe("{{client_name}}");
    expect(chipLabel({ kind: "if", name: "paid" })).toBe("{{#if paid}}");
    expect(chipLabel({ kind: "endif" })).toBe("{{/if}}");
    expect(chipLabel({ kind: "each", name: "line_items" })).toBe("{{#each line_items}}");
    expect(parseChipText("{{client_name}}")).toEqual({ kind: "merge", name: "client_name" });
    expect(parseChipText("{{#if paid}}")).toEqual({ kind: "if", name: "paid" });
    expect(parseChipText("{{/if}}")).toEqual({ kind: "endif" });
    expect(parseChipText("{{#each line_items}}")).toEqual({ kind: "each", name: "line_items" });
  });

  it("supports dotted names", () => {
    expect(parseChipText("{{customer.name}}")).toEqual({ kind: "merge", name: "customer.name" });
  });

  it("refuses malformed or unrelated text", () => {
    expect(parseChipText("{{client name}}")).toBeNull();
    expect(parseChipText("hello")).toBeNull();
    expect(parseChipText("{{}}")).toBeNull();
    expect(parseChipText("{{#if}}")).toBeNull();
  });

  it("validates chip names against the grammar", () => {
    expect(isValidChipName("client_name")).toBe(true);
    expect(isValidChipName("customer.name")).toBe(true);
    expect(isValidChipName("_x9")).toBe(true);
    expect(isValidChipName("9x")).toBe(false);
    expect(isValidChipName("client name")).toBe(false);
    expect(isValidChipName("")).toBe(false);
  });

  it("finds literal {{...}} candidates without requiring valid chip syntax", () => {
    expect(findLiteralChipCandidates("Dear {{client_name}}, re {{invoice_no}}.")).toEqual([
      "{{client_name}}",
      "{{invoice_no}}",
    ]);
    expect(findLiteralChipCandidates("no chips here")).toEqual([]);
  });
});

describe("mail-merge chip round-trip (AC-16..19)", () => {
  it("survives save and re-open in .docx, one chip of every kind, via equalStructural", async () => {
    // Compared on .sections, the same narrowing docx/roundTrip.spec.ts and
    // odf/roundTrip.spec.ts use: a freshly created (never-parsed) model's
    // defaultStyles.pageSize carries un-quantised cm floats that a
    // cm->twips->cm trip through the writer/reader legitimately perturbs
    // (SPEC-002 territory, unrelated to chip serialisation).
    const model = docWithOneOfEachChip();
    const { bytes } = await writeDocx(model);
    const reopened = await parseDocx(bytes);
    const diff = equalStructural(model.sections, reopened.sections);
    expect(diff).toBe(true);
    const runs = (reopened.sections[0].blocks as Paragraph[]).map((p) => p.runs[0]);
    expect(runs.map((r) => r.field)).toEqual([
      { kind: "merge", name: "client_name" },
      { kind: "if", name: "paid" },
      { kind: "endif" },
      { kind: "each", name: "line_items" },
    ]);
    expect(runs.every((r) => r.text === "1")).toBe(true);
  });

  it("survives save and re-open in .odt, one chip of every kind, via equalStructural", async () => {
    // The .odt reader merges the document's default font into every run it
    // reads (documented in odf/writer.ts's header comment) — matching it
    // here up front keeps this fixture's formatting round-trip-stable so
    // the assertion isolates chip serialisation, not that unrelated ODF
    // behaviour (which applies identically to a plain text run).
    const model = docWithOneOfEachChip({ fontFamily: "Calibri", fontSizePt: 11 });
    const { bytes } = await writeOdt(model);
    const reopened = await parseOdt(bytes);
    expect(equalStructural(model.sections, reopened.sections)).toBe(true);
    const runs = (reopened.sections[0].blocks as Paragraph[]).map((p) => p.runs[0]);
    expect(runs.every((r) => r.field && isMergeChipField(r.field))).toBe(true);
  });
});
