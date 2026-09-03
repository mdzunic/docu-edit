// SPEC-076 AC-76: guards the committed table fixtures under `e2e/fixtures/` —
// a four-page table with a repeating header, a banded table, a table sorted by
// two columns, and a tab-separated list, each in .docx and .odt. Every one
// opens, renders, saves and reopens `equalStructural`.
//
// This spec is READ-ONLY on e2e/fixtures — generation lives in
// scripts/build-spec076-fixtures.mjs, so a unit run can never rewrite a file
// the e2e suite is uploading (the same reasoning as fixtures.borders.spec.ts).
import { promises as fs } from "node:fs";
import path from "node:path";
import { beforeAll, describe, expect, it } from "vitest";
import { parseDocx } from "./docx/reader";
import { writeDocx } from "./docx/writer";
import { parseOdt } from "./odf/reader";
import { writeOdt } from "./odf/writer";
import { equalStructural, firstStructuralDifference } from "./equality";
import { resolveCellFormatting, tableStyleFor } from "./tableStyles";
import { tableSliceRows } from "./tables";
import { tableToText } from "./tableConvert";
import { documentBodyStats } from "./wordCount";
import {
  BANDED_STYLE_ID,
  buildTableFixtures,
  TAB_SEPARATED_LINES,
  TALL_TABLE_BODY_ROWS,
  tableFixtureModels,
} from "./fixtures.tables";
import type { DocumentModel, Table } from "./documentModel";

const FIXTURE_DIR = path.resolve(__dirname, "../../e2e/fixtures");

const REGENERATE_HINT =
  "The committed fixture no longer matches src/lib/fixtures.tables.ts. " +
  "Regenerate it with `npx vite-node scripts/build-spec076-fixtures.mjs` and commit the result.";

const committed = new Map<string, Uint8Array>();

beforeAll(async () => {
  for (const { name } of await buildTableFixtures()) {
    committed.set(name, new Uint8Array(await fs.readFile(path.join(FIXTURE_DIR, name))));
  }
});

/** The parsed package as a DocumentModel — everything the reader returned
 * plus the `format` the app layer stamps on an opened document, so a
 * same-format save takes the preserving path the app itself takes. */
function asModel(key: string, parsed: object, format: "docx" | "odt"): DocumentModel {
  return { id: key, title: key, format, ...parsed } as DocumentModel;
}

function tableIn(model: DocumentModel): Table {
  const table = model.sections[0].blocks.find((block) => block.type === "table");
  if (!table || table.type !== "table") {
    throw new Error("expected a table in the fixture");
  }
  return table;
}

describe("SPEC-076 table fixtures (AC-75, AC-76)", () => {
  it("commits eight packages — four documents in both formats", async () => {
    const built = await buildTableFixtures();
    expect(built.map((entry) => entry.name).sort()).toEqual([
      "tables-banded.docx",
      "tables-banded.odt",
      "tables-repeating-header.docx",
      "tables-repeating-header.odt",
      "tables-sorted.docx",
      "tables-sorted.odt",
      "tables-tab-separated.docx",
      "tables-tab-separated.odt",
    ]);
    for (const { name } of built) {
      expect(committed.get(name), `${name} is missing from e2e/fixtures. ${REGENERATE_HINT}`).toBeDefined();
    }
  });

  for (const fixture of tableFixtureModels()) {
    it(`${fixture.key}.docx opens, saves and reopens equalStructural`, async () => {
      const opened = asModel(fixture.key, await parseDocx(committed.get(`${fixture.key}.docx`)!), "docx");
      const { bytes } = await writeDocx(opened);
      const reopened = asModel(fixture.key, await parseDocx(bytes), "docx");
      expect(
        firstStructuralDifference(reopened.sections, opened.sections),
        `${fixture.key}.docx changed on a same-format round trip. ${REGENERATE_HINT}`,
      ).toBeNull();
    });

    it(`${fixture.key}.odt opens, saves and reopens equalStructural`, async () => {
      const opened = asModel(fixture.key, await parseOdt(committed.get(`${fixture.key}.odt`)!), "odt");
      const { bytes } = await writeOdt(opened);
      const reopened = asModel(fixture.key, await parseOdt(bytes), "odt");
      expect(
        firstStructuralDifference(reopened.sections, opened.sections),
        `${fixture.key}.odt changed on a same-format round trip. ${REGENERATE_HINT}`,
      ).toBeNull();
    });
  }
});

describe("what each fixture is for (AC-76's 'renders')", () => {
  it("the repeating-header fixture holds one header row and enough body to run over four pages", async () => {
    const opened = asModel("tall", await parseDocx(committed.get("tables-repeating-header.docx")!), "docx");
    const table = tableIn(opened);
    expect(table.rows[0].isHeader).toBe(true);
    expect(table.rows).toHaveLength(TALL_TABLE_BODY_ROWS + 1);
    // The layout rule the canvas paints with: a continuation redraws the run
    // above its first body row, and the model still holds exactly one copy.
    const slice = tableSliceRows(table, 40, 60);
    expect(slice[0]).toMatchObject({ rowIndex: 0, repeated: true });
    expect(slice[1]).toMatchObject({ rowIndex: 40, repeated: false });
    // Word count walks the model, so the repeat cannot inflate it (AC-10):
    // the header's three words are counted once, however many pages redraw
    // them. Body rows contribute "Item n" (2) + region (1) + units (1).
    const lead = opened.sections[0].blocks[0];
    const leadWords = lead.type === "paragraph" ? lead.runs.map((run) => run.text).join("").split(" ").length : 0;
    expect(documentBodyStats(opened).words).toBe(leadWords + 3 + TALL_TABLE_BODY_ROWS * 4);
  });

  it("the banded fixture resolves alternating shading without any cell carrying it (AC-24, AC-26)", async () => {
    const opened = asModel("banded", await parseDocx(committed.get("tables-banded.docx")!), "docx");
    const table = tableIn(opened);
    expect(table.styleId).toBe(BANDED_STYLE_ID);
    expect(table.look).toEqual({ firstRow: true, bandedRows: true });
    const style = tableStyleFor(opened.styles, table);
    expect(style?.banding?.bandedRows?.shading).toBeDefined();
    // Row 0 is the header region; banding counts from row 1.
    const banded = resolveCellFormatting(table, style, 1, 0);
    const plain = resolveCellFormatting(table, style, 2, 0);
    expect(banded.shadingFill).toBe(style?.banding?.bandedRows?.shading);
    expect(plain.shadingFill).not.toBe(banded.shadingFill);
    // Nothing was baked onto a cell.
    expect(table.rows[1].cells[0].shadingFill).toBeUndefined();
  });

  it("the sorted fixture is in region-ascending, units-descending order (AC-46, AC-47)", async () => {
    const opened = asModel("sorted", await parseOdt(committed.get("tables-sorted.odt")!), "odt");
    const table = tableIn(opened);
    const rows = table.rows.slice(1).map((row) => row.cells.map((cell) => cell.blocks[0].runs[0]?.text ?? ""));
    expect(rows.map((row) => row[0])).toEqual(["East", "North", "North", "South", "South"]);
    expect(rows.slice(1, 3).map((row) => row[1])).toEqual(["40", "40"]);
    // The tie between the two North/40 rows keeps the ORIGINAL order (Grace
    // came before Edsger in the unsorted model) — the sort is stable.
    expect(rows.slice(1, 3).map((row) => row[2])).toEqual(["Grace", "Edsger"]);
    expect(rows.slice(3).map((row) => row[1])).toEqual(["12", "7"]);
  });

  it("the tab-separated fixture converts to a table and back to the same paragraphs (AC-65)", async () => {
    const opened = asModel("tabs", await parseDocx(committed.get("tables-tab-separated.docx")!), "docx");
    const lines = opened.sections[0].blocks
      .filter((block) => block.type === "paragraph")
      .map((block) => (block.type === "paragraph" ? block.runs.map((run) => run.text).join("") : ""));
    for (const line of TAB_SEPARATED_LINES) {
      expect(lines).toContain(line);
    }
  });

  it("a banded table saved as .odt reopens with its look and banding intact (AC-68, AC-69)", async () => {
    const opened = asModel("banded", await parseOdt(committed.get("tables-banded.odt")!), "odt");
    const table = tableIn(opened);
    expect(table.look).toEqual({ firstRow: true, bandedRows: true });
    const style = tableStyleFor(opened.styles, table);
    expect(style?.banding?.firstRow?.bold).toBe(true);
    expect(equalStructural(style?.banding?.bandedRows?.size, 1)).toBe(true);
  });

  it("flattening the sorted fixture's table yields one paragraph per row (AC-62)", async () => {
    const opened = asModel("sorted", await parseDocx(committed.get("tables-sorted.docx")!), "docx");
    const paragraphs = tableToText(tableIn(opened), "\t");
    expect(paragraphs).toHaveLength(6);
    expect(paragraphs[0].runs.map((run) => run.text).join("")).toBe("Region\tUnits\tOwner");
  });
});
