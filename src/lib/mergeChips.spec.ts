import { describe, expect, it } from "vitest";
import { createDocument, isMergeChipField, type Paragraph, type Table } from "./documentModel";
import { equalStructural } from "./equality";
import {
  chipLabel,
  convertLiteralChipsToRuns,
  createEachChipRun,
  createEndifChipRun,
  createIfChipRun,
  createMergeChipRun,
  findLiteralChipCandidates,
  hasConvertibleLiteralChips,
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

function convertTestRun(text: string, extra: Record<string, unknown> = {}) {
  return { text, formatting: {}, ...extra };
}

describe("Convert {{…}} text to merge fields (AC-20)", () => {
  it("converts every well-formed {{...}} match, splitting a run around it", () => {
    const model = createDocument("convert-test");
    model.sections[0].blocks = [{ type: "paragraph", runs: [convertTestRun("Dear {{client_name}}, re {{invoice_no}}.")] }];
    const { model: next, converted } = convertLiteralChipsToRuns(model);
    expect(converted).toBe(2);
    const runs = (next.sections[0].blocks[0] as Paragraph).runs;
    expect(runs.map((r) => r.text)).toEqual(["Dear ", "1", ", re ", "1", "."]);
    expect(runs[1].field).toEqual({ kind: "merge", name: "client_name" });
    expect(runs[3].field).toEqual({ kind: "merge", name: "invoice_no" });
  });

  it("converts #if/#endif/#each syntax too", () => {
    const model = createDocument("convert-test-2");
    model.sections[0].blocks = [
      { type: "paragraph", runs: [convertTestRun("{{#if paid}}")] },
      { type: "paragraph", runs: [convertTestRun("Paid")] },
      { type: "paragraph", runs: [convertTestRun("{{/if}}")] },
      { type: "paragraph", runs: [convertTestRun("{{#each line_items}}")] },
    ];
    const { model: next, converted } = convertLiteralChipsToRuns(model);
    expect(converted).toBe(3);
    const blocks = next.sections[0].blocks as Paragraph[];
    expect(blocks[0].runs[0].field).toEqual({ kind: "if", name: "paid" });
    expect(blocks[2].runs[0].field).toEqual({ kind: "endif" });
    expect(blocks[3].runs[0].field).toEqual({ kind: "each", name: "line_items" });
  });

  it("leaves malformed {{...}} text alone (a space in the name is not valid grammar)", () => {
    const model = createDocument("convert-test-3");
    model.sections[0].blocks = [{ type: "paragraph", runs: [convertTestRun("{{client name}}")] }];
    const { model: next, converted } = convertLiteralChipsToRuns(model);
    expect(converted).toBe(0);
    expect((next.sections[0].blocks[0] as Paragraph).runs[0].text).toBe("{{client name}}");
  });

  it("preserves formatting and styleId on the split text pieces and the new chip", () => {
    const model = createDocument("convert-test-4");
    model.sections[0].blocks = [
      { type: "paragraph", runs: [{ text: "Hi {{name}}!", formatting: { bold: true }, styleId: "Emphasis" }] },
    ];
    const { model: next } = convertLiteralChipsToRuns(model);
    const runs = (next.sections[0].blocks[0] as Paragraph).runs;
    expect(runs[0].formatting).toEqual({ bold: true });
    expect(runs[0].styleId).toBe("Emphasis");
    expect(runs[1].formatting).toEqual({ bold: true });
    expect(runs[1].styleId).toBe("Emphasis");
  });

  it("never touches an existing chip's placeholder text", () => {
    const model = createDocument("convert-test-5");
    model.sections[0].blocks = [{ type: "paragraph", runs: [createMergeChipRun("already_a_chip")] }];
    const { model: next, converted } = convertLiteralChipsToRuns(model);
    expect(converted).toBe(0);
    expect((next.sections[0].blocks[0] as Paragraph).runs[0].field).toEqual({ kind: "merge", name: "already_a_chip" });
  });

  it("converts matches inside table cells, headers, footers and notes", () => {
    const model = createDocument("convert-test-6");
    const table: Table = {
      type: "table",
      columnWidthsCm: [5],
      rows: [{ cells: [{ blocks: [{ type: "paragraph", runs: [convertTestRun("{{sku}}")] }] }] }],
    };
    model.sections[0].blocks = [table];
    model.sections[0].headers = { default: { blocks: [{ type: "paragraph", runs: [convertTestRun("{{header_field}}")] }] } };
    model.sections[0].footers = { default: { blocks: [{ type: "paragraph", runs: [convertTestRun("{{footer_field}}")] }] } };
    model.notes = [{ noteId: "n1", kind: "footnote", blocks: [{ type: "paragraph", runs: [convertTestRun("{{note_field}}")] }] }];

    const { model: next, converted } = convertLiteralChipsToRuns(model);
    expect(converted).toBe(4);
    expect(((next.sections[0].blocks[0] as Table).rows[0].cells[0].blocks[0] as Paragraph).runs[0].field).toEqual({
      kind: "merge",
      name: "sku",
    });
    expect((next.sections[0].headers!.default!.blocks[0] as Paragraph).runs[0].field).toEqual({ kind: "merge", name: "header_field" });
    expect((next.sections[0].footers!.default!.blocks[0] as Paragraph).runs[0].field).toEqual({ kind: "merge", name: "footer_field" });
    expect((next.notes![0].blocks[0] as Paragraph).runs[0].field).toEqual({ kind: "merge", name: "note_field" });
  });

  it("does not mutate the input model", () => {
    const model = createDocument("convert-test-7");
    model.sections[0].blocks = [{ type: "paragraph", runs: [convertTestRun("{{name}}")] }];
    convertLiteralChipsToRuns(model);
    expect((model.sections[0].blocks[0] as Paragraph).runs[0].field).toBeUndefined();
  });

  it("hasConvertibleLiteralChips agrees with what the converter actually converts", () => {
    const empty = createDocument("convert-check-1");
    expect(hasConvertibleLiteralChips(empty)).toBe(false);

    const withMalformed = createDocument("convert-check-2");
    withMalformed.sections[0].blocks = [{ type: "paragraph", runs: [convertTestRun("{{client name}}")] }];
    expect(hasConvertibleLiteralChips(withMalformed)).toBe(false);

    const withValid = createDocument("convert-check-3");
    withValid.sections[0].blocks = [{ type: "paragraph", runs: [convertTestRun("{{client_name}}")] }];
    expect(hasConvertibleLiteralChips(withValid)).toBe(true);
  });
});
