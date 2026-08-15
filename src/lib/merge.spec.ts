import { describe, expect, it } from "vitest";
import { createDocument, type Paragraph, type Table, type TableRow } from "./documentModel";
import { createEachChipRun, createEndifChipRun, createIfChipRun, createMergeChipRun } from "./mergeChips";
import {
  buildFieldMappingReport,
  MergeTemplateError,
  mergeZipFileName,
  namedArrayFieldList,
  namedFieldList,
  parseCsvRows,
  parseCsvSource,
  parseJsonSource,
  resolveRecordToModel,
  runMerge,
  slugify,
  validateTemplate,
  zipEntryName,
  type ParsedDataSource,
} from "./merge";
import { unzip } from "./core/zip";
import { parseDocx } from "./docx/reader";
import { parseOdt } from "./odf/reader";

function textRun(text: string) {
  return { text, formatting: {} };
}

function paragraph(...runs: Paragraph["runs"]): Paragraph {
  return { type: "paragraph", runs };
}

function docWith(blocks: Array<Paragraph | Table>) {
  const model = createDocument("merge-test");
  model.sections[0].blocks = blocks;
  return model;
}

// --- CSV parsing ----------------------------------------------------------

describe("parseCsvRows", () => {
  it("splits simple rows", () => {
    expect(parseCsvRows("a,b\n1,2\n")).toEqual([
      ["a", "b"],
      ["1", "2"],
    ]);
  });

  it("handles quoted fields with embedded commas, newlines and escaped quotes", () => {
    const text = 'name,note\n"Acme, Inc.","Line1\nLine2"\n"She said ""hi"""\n';
    expect(parseCsvRows(text)).toEqual([
      ["name", "note"],
      ["Acme, Inc.", "Line1\nLine2"],
      ['She said "hi"'],
    ]);
  });

  it("does not drop a row whose only cell is empty", () => {
    expect(parseCsvRows("a\n\n")).toEqual([["a"], [""]]);
  });

  it("does not add a phantom trailing row for a trailing newline", () => {
    expect(parseCsvRows("a,b\n1,2\n")).toHaveLength(2);
  });

  it("captures a final row with no trailing newline", () => {
    expect(parseCsvRows("a\nb")).toEqual([["a"], ["b"]]);
  });
});

describe("parseCsvSource", () => {
  it("builds one record per row after the header, every value a string", () => {
    const source = parseCsvSource("name,paid\nAcme,true\nBeta,\n");
    expect(source.kind).toBe("csv");
    expect(source.columns).toEqual(["name", "paid"]);
    expect(source.arrayColumns).toEqual([]);
    expect(source.records).toEqual([
      { name: "Acme", paid: "true" },
      { name: "Beta", paid: "" },
    ]);
  });
});

describe("parseJsonSource", () => {
  it("parses a top-level array of objects and finds array columns", () => {
    const source = parseJsonSource(
      JSON.stringify([{ name: "Acme", line_items: [{ sku: "A1" }, { sku: "A2" }] }, { name: "Beta", line_items: [] }]),
    );
    expect(source.kind).toBe("json");
    expect(source.columns.sort()).toEqual(["line_items", "name"]);
    expect(source.arrayColumns).toEqual(["line_items"]);
    expect(source.records).toHaveLength(2);
  });

  it("refuses non-array top-level JSON", () => {
    expect(() => parseJsonSource(JSON.stringify({ name: "Acme" }))).toThrow();
  });

  it("refuses invalid JSON text", () => {
    expect(() => parseJsonSource("{not json")).toThrow();
  });
});

// --- named field list -------------------------------------------------------

describe("namedFieldList", () => {
  it("lists source columns first, then names already used by a chip, deduplicated", () => {
    const model = docWith([paragraph(createMergeChipRun("client_name")), paragraph(createMergeChipRun("extra_note"))]);
    const source: ParsedDataSource = { kind: "csv", columns: ["client_name", "invoice_no"], arrayColumns: [], records: [] };
    const list = namedFieldList(source, model);
    expect(list.map((f) => f.name)).toEqual(["client_name", "invoice_no", "extra_note"]);
  });

  it("works with no attached source — a template can be written before its data exists", () => {
    const model = docWith([paragraph(createMergeChipRun("client_name"))]);
    expect(namedFieldList(null, model).map((f) => f.name)).toEqual(["client_name"]);
  });

  it("narrows to array fields for the repeat-row list, from a JSON source or used each chips", () => {
    const model = docWith([paragraph(createEachChipRun("used_array"))]);
    const source: ParsedDataSource = { kind: "json", columns: ["line_items", "name"], arrayColumns: ["line_items"], records: [] };
    expect(namedArrayFieldList(source, model).map((f) => f.name)).toEqual(["line_items", "used_array"]);
  });
});

describe("buildFieldMappingReport", () => {
  it("reports mapped/unmapped columns and unmapped chips", () => {
    const model = docWith([paragraph(createMergeChipRun("client_name")), paragraph(createMergeChipRun("typo_field"))]);
    const source: ParsedDataSource = { kind: "csv", columns: ["client_name", "Client Name"], arrayColumns: [], records: [] };
    const report = buildFieldMappingReport(source, model);
    expect(report.columns).toEqual([
      { name: "client_name", mapped: true },
      { name: "Client Name", mapped: false },
    ]);
    expect(report.unmappedChips).toEqual(["typo_field"]);
  });
});

// --- template validation ----------------------------------------------------

describe("validateTemplate", () => {
  it("passes a well-formed template", () => {
    const model = docWith([paragraph(createIfChipRun("paid")), paragraph(textRun("Thanks!")), paragraph(createEndifChipRun())]);
    expect(validateTemplate(model)).toEqual([]);
  });

  it("flags an unclosed {{#if}}", () => {
    const model = docWith([paragraph(createIfChipRun("paid"))]);
    const errors = validateTemplate(model);
    expect(errors).toHaveLength(1);
    expect(errors[0].kind).toBe("unpaired-if");
  });

  it("flags a stray {{/if}} with no open {{#if}}", () => {
    const model = docWith([paragraph(createEndifChipRun())]);
    const errors = validateTemplate(model);
    expect(errors).toHaveLength(1);
    expect(errors[0].kind).toBe("unpaired-endif");
  });

  it("flags a pair that starts inside a table cell and ends outside it as crossed", () => {
    const table: Table = {
      type: "table",
      columnWidthsCm: [5],
      rows: [{ cells: [{ blocks: [paragraph(createIfChipRun("paid"))] }] }],
    };
    const model = docWith([table, paragraph(createEndifChipRun())]);
    const errors = validateTemplate(model);
    expect(errors.some((e) => e.kind === "crossed")).toBe(true);
  });

  it("flags two {{#each}} chips in one row", () => {
    const table: Table = {
      type: "table",
      columnWidthsCm: [5, 5],
      rows: [
        {
          cells: [{ blocks: [paragraph(createEachChipRun("a"))] }, { blocks: [paragraph(createEachChipRun("b"))] }],
        },
      ],
    };
    const model = docWith([table]);
    const errors = validateTemplate(model);
    expect(errors.some((e) => e.kind === "nested-each")).toBe(true);
  });

  it("flags an {{#if}} chip in a header as outside the body", () => {
    const model = createDocument("merge-header");
    model.sections[0].headers = { default: { blocks: [paragraph(createIfChipRun("paid"))] } };
    const errors = validateTemplate(model);
    expect(errors.some((e) => e.kind === "outside-body")).toBe(true);
  });

  it("blocks runMerge before producing any document", async () => {
    const model = docWith([paragraph(createIfChipRun("paid"))]);
    const source: ParsedDataSource = { kind: "csv", columns: [], arrayColumns: [], records: [{}] };
    await expect(runMerge(model, source, { format: "docx", delivery: "zip" })).rejects.toBeInstanceOf(MergeTemplateError);
  });
});

// --- conditional resolution --------------------------------------------------

describe("conditional block resolution", () => {
  it("block form: true keeps the paragraphs between, drops only the delimiters", () => {
    const model = docWith([paragraph(createIfChipRun("paid")), paragraph(textRun("Thanks!")), paragraph(createEndifChipRun())]);
    const resolved = resolveRecordToModel(model, { paid: true }, "json", 0);
    const texts = (resolved.sections[0].blocks as Paragraph[]).map((p) => p.runs.map((r) => r.text).join(""));
    expect(texts).toEqual(["Thanks!"]);
  });

  it("block form: false drops the delimiters and everything between (no blank lines left)", () => {
    const model = docWith([
      paragraph(textRun("Before")),
      paragraph(createIfChipRun("paid")),
      paragraph(textRun("Thanks!")),
      paragraph(createEndifChipRun()),
      paragraph(textRun("After")),
    ]);
    const resolved = resolveRecordToModel(model, { paid: false }, "json", 0);
    const texts = (resolved.sections[0].blocks as Paragraph[]).map((p) => p.runs.map((r) => r.text).join(""));
    expect(texts).toEqual(["Before", "After"]);
  });

  it("inline form: true removes only the two chip runs", () => {
    const model = docWith([paragraph(textRun("Status: "), createIfChipRun("paid"), textRun("PAID"), createEndifChipRun(), textRun("."))]);
    const resolved = resolveRecordToModel(model, { paid: true }, "json", 0);
    const texts = (resolved.sections[0].blocks as Paragraph[])[0].runs.map((r) => r.text).join("");
    expect(texts).toBe("Status: PAID.");
  });

  it("inline form: false removes the whole range inclusive", () => {
    const model = docWith([paragraph(textRun("Status: "), createIfChipRun("paid"), textRun("PAID"), createEndifChipRun(), textRun("."))]);
    const resolved = resolveRecordToModel(model, { paid: false }, "json", 0);
    const texts = (resolved.sections[0].blocks as Paragraph[])[0].runs.map((r) => r.text).join("");
    expect(texts).toBe("Status: .");
  });

  it("inline form spanning two paragraphs joins them when false", () => {
    const model = docWith([
      paragraph(textRun("Start "), createIfChipRun("paid")),
      paragraph(createEndifChipRun(), textRun(" End")),
    ]);
    const resolved = resolveRecordToModel(model, { paid: false }, "json", 0);
    const blocks = resolved.sections[0].blocks as Paragraph[];
    expect(blocks).toHaveLength(1);
    expect(blocks[0].runs.map((r) => r.text).join("")).toBe("Start  End");
  });

  it("nesting to any depth resolves independently", () => {
    const model = docWith([
      paragraph(createIfChipRun("outer")),
      paragraph(createIfChipRun("inner")),
      paragraph(textRun("Deep")),
      paragraph(createEndifChipRun()),
      paragraph(createEndifChipRun()),
    ]);
    const resolvedBothTrue = resolveRecordToModel(model, { outer: true, inner: true }, "json", 0);
    expect((resolvedBothTrue.sections[0].blocks as Paragraph[]).map((p) => p.runs.map((r) => r.text).join(""))).toEqual(["Deep"]);
    const resolvedOuterTrueInnerFalse = resolveRecordToModel(model, { outer: true, inner: false }, "json", 0);
    expect((resolvedOuterTrueInnerFalse.sections[0].blocks as Paragraph[]).map((p) => p.runs.map((r) => r.text).join(""))).toEqual([]);
  });

  it("two sibling inline conditionals sharing one paragraph both resolve", () => {
    // "A {{#if x}}X{{/if}} B {{#if y}}Y{{/if}} C", x false, y true.
    const model = docWith([
      paragraph(
        textRun("A "),
        createIfChipRun("x"),
        textRun("X"),
        createEndifChipRun(),
        textRun(" B "),
        createIfChipRun("y"),
        textRun("Y"),
        createEndifChipRun(),
        textRun(" C"),
      ),
    ]);
    const resolved = resolveRecordToModel(model, { x: false, y: true }, "json", 0);
    const blocks = resolved.sections[0].blocks as Paragraph[];
    expect(blocks).toHaveLength(1);
    expect(blocks[0].runs.map((r) => r.text).join("")).toBe("A  B Y C");
    // No if/endif chip runs may survive into the resolved output.
    expect(blocks[0].runs.some((r) => r.field && (r.field.kind === "if" || r.field.kind === "endif"))).toBe(false);
  });

  it("three sibling inline conditionals in one paragraph all resolve, in either truth combination", () => {
    const model = docWith([
      paragraph(
        createIfChipRun("a"),
        textRun("A"),
        createEndifChipRun(),
        createIfChipRun("b"),
        textRun("B"),
        createEndifChipRun(),
        createIfChipRun("c"),
        textRun("C"),
        createEndifChipRun(),
      ),
    ]);
    const resolved = resolveRecordToModel(model, { a: true, b: false, c: true }, "json", 0);
    const blocks = resolved.sections[0].blocks as Paragraph[];
    expect(blocks).toHaveLength(1);
    expect(blocks[0].runs.map((r) => r.text).join("")).toBe("AC");
    expect(blocks[0].runs.some((r) => r.field && (r.field.kind === "if" || r.field.kind === "endif"))).toBe(false);
  });

  it("a sibling conditional carried into the closing fragment of a block-spanning conditional resolves", () => {
    const model = docWith([
      paragraph(textRun("Start "), createIfChipRun("outer")),
      paragraph(createEndifChipRun(), textRun(" mid "), createIfChipRun("inner"), textRun("Y"), createEndifChipRun(), textRun(" end")),
    ]);
    const resolved = resolveRecordToModel(model, { outer: true, inner: false }, "json", 0);
    const blocks = resolved.sections[0].blocks as Paragraph[];
    expect(blocks.map((p) => p.runs.map((r) => r.text).join(""))).toEqual(["Start ", " mid  end"]);
    expect(blocks.every((p) => p.runs.every((r) => !r.field || (r.field.kind !== "if" && r.field.kind !== "endif")))).toBe(true);
  });
});

describe("truthiness table", () => {
  it.each([
    [undefined, false],
    [null, false],
    [false, false],
    [0, false],
    ["", false],
    [[], false],
    [{}, false],
    ["no", true],
    [1, true],
    ["yes", true],
  ])("json %o -> %s", (value, expected) => {
    const model = docWith([paragraph(createIfChipRun("x")), paragraph(textRun("Y")), paragraph(createEndifChipRun())]);
    const resolved = resolveRecordToModel(model, { x: value }, "json", 0);
    const kept = (resolved.sections[0].blocks as Paragraph[]).length > 0;
    expect(kept).toBe(expected);
  });

  it.each([
    ["", false],
    ["false", false],
    ["FALSE", false],
    ["no", false],
    ["NO", false],
    ["0", false],
    ["yes", true],
    ["1", true],
  ])("csv %s -> %s", (value, expected) => {
    const model = docWith([paragraph(createIfChipRun("x")), paragraph(textRun("Y")), paragraph(createEndifChipRun())]);
    const resolved = resolveRecordToModel(model, { x: value }, "csv", 0);
    const kept = (resolved.sections[0].blocks as Paragraph[]).length > 0;
    expect(kept).toBe(expected);
  });

  it("a missing field a condition tests is false and not reported", () => {
    const report = { missingFields: [] as unknown[], arrayFieldErrors: [] as unknown[] };
    const model = docWith([paragraph(createIfChipRun("nope")), paragraph(textRun("Y")), paragraph(createEndifChipRun())]);
    const resolved = resolveRecordToModel(model, {}, "json", 0, report as never);
    expect((resolved.sections[0].blocks as Paragraph[])).toEqual([]);
    expect(report.missingFields).toEqual([]);
  });
});

// --- merge chip resolution + gaps -------------------------------------------

describe("merge chip resolution", () => {
  it("replaces a resolved chip with plain text", () => {
    const model = docWith([paragraph(textRun("Dear "), createMergeChipRun("client_name"), textRun(","))]);
    const resolved = resolveRecordToModel(model, { client_name: "Acme" }, "json", 0);
    expect((resolved.sections[0].blocks as Paragraph[])[0].runs.map((r) => r.text).join("")).toBe("Dear Acme,");
  });

  it("resolves dotted names through nested JSON objects", () => {
    const model = docWith([paragraph(createMergeChipRun("customer.name"))]);
    const resolved = resolveRecordToModel(model, { customer: { name: "Acme" } }, "json", 0);
    expect((resolved.sections[0].blocks as Paragraph[])[0].runs[0].text).toBe("Acme");
  });

  it("prefers an exact column name over dotted traversal (CSV)", () => {
    const model = docWith([paragraph(createMergeChipRun("customer.name"))]);
    const resolved = resolveRecordToModel(model, { "customer.name": "Acme literal column" }, "csv", 0);
    expect((resolved.sections[0].blocks as Paragraph[])[0].runs[0].text).toBe("Acme literal column");
  });

  it("leaves a clearly marked, highlighted gap for a missing field and reports it", () => {
    const report = { missingFields: [] as Array<{ recordIndex: number; field: string }>, arrayFieldErrors: [] as unknown[] };
    const model = docWith([paragraph(createMergeChipRun("client_name"))]);
    const resolved = resolveRecordToModel(model, {}, "json", 2, report as never);
    const run = (resolved.sections[0].blocks as Paragraph[])[0].runs[0];
    expect(run.text).toBe("[client_name]");
    expect(run.formatting.highlight).toBe("yellow");
    expect(report.missingFields).toEqual([{ recordIndex: 2, field: "client_name" }]);
  });

  it("never evaluates merge-field syntax present in the data — inserted as plain text", () => {
    const model = docWith([paragraph(createMergeChipRun("payload"))]);
    const resolved = resolveRecordToModel(model, { payload: "{{#if x}}{{/if}} <script>alert(1)</script>" }, "json", 0);
    const run = (resolved.sections[0].blocks as Paragraph[])[0].runs[0];
    expect(run.text).toBe("{{#if x}}{{/if}} <script>alert(1)</script>");
    expect(run.field).toBeUndefined();
  });
});

// --- repeating table rows ----------------------------------------------------

function rowWithEach(fieldChip: ReturnType<typeof createEachChipRun>, extraRuns: Paragraph["runs"] = []): TableRow {
  return { cells: [{ blocks: [paragraph(fieldChip, ...extraRuns)] }, { blocks: [paragraph(createMergeChipRun("sku"))] }] };
}

describe("repeating table rows", () => {
  it("expands one copy per array element, each chip removed, element scope resolves first", () => {
    const table: Table = { type: "table", columnWidthsCm: [5, 5], rows: [rowWithEach(createEachChipRun("line_items"))] };
    const model = docWith([table]);
    const resolved = resolveRecordToModel(model, { line_items: [{ sku: "A1" }, { sku: "A2" }] }, "json", 0);
    const resolvedTable = resolved.sections[0].blocks[0] as Table;
    expect(resolvedTable.rows).toHaveLength(2);
    expect((resolvedTable.rows[0].cells[1].blocks[0] as Paragraph).runs[0].text).toBe("A1");
    expect((resolvedTable.rows[1].cells[1].blocks[0] as Paragraph).runs[0].text).toBe("A2");
    // the each chip itself is gone
    expect(resolvedTable.rows[0].cells[0].blocks[0].runs.every((r) => !r.field)).toBe(true);
  });

  it("falls back to the record root when a name isn't on the current element", () => {
    const table: Table = {
      type: "table",
      columnWidthsCm: [5, 5],
      rows: [{ cells: [{ blocks: [paragraph(createEachChipRun("line_items"))] }, { blocks: [paragraph(createMergeChipRun("invoice_no"))] }] }],
    };
    const model = docWith([table]);
    const resolved = resolveRecordToModel(model, { invoice_no: "INV-1", line_items: [{ sku: "A1" }] }, "json", 0);
    const resolvedTable = resolved.sections[0].blocks[0] as Table;
    expect((resolvedTable.rows[0].cells[1].blocks[0] as Paragraph).runs[0].text).toBe("INV-1");
  });

  it("removes the row for an empty array, and the whole table when that leaves no rows", () => {
    const table: Table = { type: "table", columnWidthsCm: [5, 5], rows: [rowWithEach(createEachChipRun("line_items"))] };
    const model = docWith([table]);
    const resolved = resolveRecordToModel(model, { line_items: [] }, "json", 0);
    expect(resolved.sections[0].blocks).toEqual([]);
  });

  it("keeps a header row intact and out of the repeat", () => {
    const table: Table = {
      type: "table",
      columnWidthsCm: [5, 5],
      rows: [
        { isHeader: true, cells: [{ blocks: [paragraph(textRun("SKU"))] }, { blocks: [paragraph(textRun("Name"))] }] },
        rowWithEach(createEachChipRun("line_items")),
      ],
    };
    const model = docWith([table]);
    const resolved = resolveRecordToModel(model, { line_items: [{ sku: "A1" }, { sku: "A2" }] }, "json", 0);
    const resolvedTable = resolved.sections[0].blocks[0] as Table;
    expect(resolvedTable.rows).toHaveLength(3);
    expect(resolvedTable.rows[0].isHeader).toBe(true);
  });

  it("a missing/non-array field is a data error: the row is emitted once with gaps, reported, run continues", () => {
    const report = { missingFields: [] as Array<{ recordIndex: number; field: string }>, arrayFieldErrors: [] as Array<{ recordIndex: number; field: string }> };
    const table: Table = { type: "table", columnWidthsCm: [5, 5], rows: [rowWithEach(createEachChipRun("line_items"))] };
    const model = docWith([table]);
    const resolved = resolveRecordToModel(model, {}, "json", 5, report as never);
    const resolvedTable = resolved.sections[0].blocks[0] as Table;
    expect(resolvedTable.rows).toHaveLength(1);
    expect((resolvedTable.rows[0].cells[1].blocks[0] as Paragraph).runs[0].text).toBe("[sku]");
    expect(report.arrayFieldErrors).toEqual([{ recordIndex: 5, field: "line_items" }]);
  });

  it("CSV sources never have arrays, so {{#each}} is always the not-an-array data error", () => {
    const report = { missingFields: [] as unknown[], arrayFieldErrors: [] as Array<{ recordIndex: number; field: string }> };
    const table: Table = { type: "table", columnWidthsCm: [5, 5], rows: [rowWithEach(createEachChipRun("line_items"))] };
    const model = docWith([table]);
    resolveRecordToModel(model, { line_items: "not an array" }, "csv", 0, report as never);
    expect(report.arrayFieldErrors).toEqual([{ recordIndex: 0, field: "line_items" }]);
  });
});

// --- zip naming / output -----------------------------------------------------

describe("zip naming", () => {
  it("mergeZipFileName appends -merge.zip to the sanitised title", () => {
    expect(mergeZipFileName("Invoices")).toBe("Invoices-merge.zip");
    expect(mergeZipFileName("  ")).toMatch(/-merge\.zip$/);
  });

  it("slugify keeps safe characters, collapses the rest, trims and truncates", () => {
    expect(slugify("Acme, Inc. / Ltd!")).toBe("Acme-Inc.-Ltd");
    expect(slugify("a".repeat(100))).toHaveLength(60);
  });

  it("zipEntryName zero-pads the record number to the record count's width", () => {
    expect(zipEntryName(0, 120, "acme", "docx")).toBe("001-acme.docx");
    expect(zipEntryName(0, 5, null, "docx")).toBe("1.docx");
    expect(zipEntryName(9, 12, "beta", "odt")).toBe("10-beta.odt");
  });

  it("names stay unique by construction even when the name field repeats", () => {
    const names = [0, 1, 2].map((i) => zipEntryName(i, 3, "same", "docx"));
    expect(new Set(names).size).toBe(3);
  });
});

describe("runMerge", () => {
  it("produces one document per record as a zip, deliverable and re-openable, data errors never failing the run", async () => {
    const model = docWith([paragraph(textRun("Dear "), createMergeChipRun("client_name"), textRun(", total "), createMergeChipRun("total"))]);
    const source: ParsedDataSource = {
      kind: "json",
      columns: ["client_name", "total"],
      arrayColumns: [],
      records: [{ client_name: "Acme" }, { client_name: "Beta", total: 42 }],
    };
    const result = await runMerge(model, source, { format: "docx", delivery: "zip" });
    expect(result.delivery).toBe("zip");
    expect(result.fileName).toMatch(/-merge\.zip$/);
    expect(result.report.recordCount).toBe(2);
    expect(result.report.missingFields).toEqual([{ recordIndex: 0, field: "total" }]);
    const files = await unzip(result.bytes);
    expect([...files.keys()].sort()).toEqual(["1.docx", "2.docx"]);
    const first = await parseDocx(files.get("1.docx")!);
    const text = (first.sections[0].blocks[0] as Paragraph).runs.map((r) => r.text).join("");
    expect(text).toBe("Dear Acme, total [total]");
  });

  it("produces a single multi-page document with a page break between records", async () => {
    const model = docWith([paragraph(textRun("Hello "), createMergeChipRun("client_name"))]);
    const source: ParsedDataSource = { kind: "json", columns: ["client_name"], arrayColumns: [], records: [{ client_name: "Acme" }, { client_name: "Beta" }] };
    const result = await runMerge(model, source, { format: "odt", delivery: "single" });
    expect(result.delivery).toBe("single");
    const reopened = await parseOdt(result.bytes);
    const blocks = reopened.sections[0].blocks;
    expect(blocks.some((b) => b.type === "break")).toBe(true);
    const texts = blocks.filter((b): b is Paragraph => b.type === "paragraph").map((p) => p.runs.map((r) => r.text).join(""));
    expect(texts).toEqual(["Hello Acme", "Hello Beta"]);
  });
});
