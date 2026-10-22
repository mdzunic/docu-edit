// SPEC-130's ODS data-validation fixture (spec130-validation.ods, built by
// fixtures/build.ts): the committed e2e file is the builder's bytes and a
// whole package, Calc's shapes are really in content.xml (the prefix's
// `table:content-validations`, the cells' names, a repeated run of
// attribute-only empty cells), and it reads as SPEC130_VALIDATION_ODS
// records — the notes rule as ONE rectangle with no Cell stored for it, the
// orphan kept verbatim, the error-macro rule enforcing nothing; untouched it
// writes back byte-identical, edited the block keeps every definition.
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { equalBytes } from "../core/equality";
import { unzip } from "../core/zip";
import { parseRange, type GridRange } from "./address";
import { buildSpec130ValidationOdsBytes, SPEC130_VALIDATION_ODS } from "./fixtures/build";
import { packageIntegrity } from "./fixtures/packageIntegrity";
import { setCellEntry } from "./cells";
import { readSpreadsheetBytes } from "./open";
import { writeOds } from "./ods/writer";
import type { DataValidation } from "./validation";

const E2E_FIXTURES = join(__dirname, "../../../e2e/fixtures");
const decoder = new TextDecoder();
const O = SPEC130_VALIDATION_ODS;
const range = (a1: string): GridRange => parseRange(a1)!;

/** A rule as the fixture records it: no id or source. */
function shape(rule: DataValidation): Omit<DataValidation, "id" | "sourceXml"> {
  const { id: _id, sourceXml: _source, ...rest } = rule;
  return rest;
}

const expected = O.rules.map(({ ranges, ...rest }) => ({ ...rest, ranges: ranges.map(range) }));

describe("spec130-validation.ods", () => {
  const bytes = buildSpec130ValidationOdsBytes();

  it("is the builder's bytes, a whole package, with Calc's content validations in the prefix and the names on the cells", async () => {
    expect(equalBytes(new Uint8Array(readFileSync(join(E2E_FIXTURES, "spec130-validation.ods"))), await bytes)).toBe(true);
    expect(await packageIntegrity(await bytes)).toEqual([]);
    const content = decoder.decode((await unzip(await bytes)).get("content.xml"));
    expect(content).toContain(`<table:calculation-settings table:automatic-find-labels="false" table:use-regular-expressions="false" table:use-wildcards="true"/><table:content-validations>`);
    expect(content).toContain(`<table:content-validations>${O.definitions.join("")}${O.orphan}</table:content-validations><table:table table:name="Form"`);
    // The notes rule's rows 12–1000: one repeated row, its C cell attribute-only.
    expect(content).toContain(
      `<table:table-row table:style-name="ro1" table:number-rows-repeated="${O.notesRun.rows}"><table:table-cell table:number-columns-repeated="2"/>` +
        `<table:table-cell table:content-validation-name="val11"/><table:table-cell table:number-columns-repeated="16381"/></table:table-row>`,
    );
    expect(content).toContain(`table:display-list="sort-ascending"`);
    expect(content).toContain(`<table:error-macro table:execute="true"/>`);
    expect(content).not.toContain(`table:content-validation-name="val12"`);
  });

  it("reads every rule of Form as the fixture records it, in definition order; Lists has none", async () => {
    const model = await readSpreadsheetBytes(await bytes, "ods", "t");
    expect(model.sheets.map((sheet) => sheet.name)).toEqual(O.tableNames);
    const [form, lists] = model.sheets;
    expect(form!.validations.map(shape)).toEqual(expected);
    expect(lists!.validations).toEqual([]);
    // Each rule's sourceXml is its definition verbatim.
    expect(form!.validations.map((rule) => rule.sourceXml)).toEqual([...O.definitions]);
  });

  it("the notes rule over a repeated run is ONE rectangle and stores no Cell in C2:C1000", async () => {
    const model = await readSpreadsheetBytes(await bytes, "ods", "t");
    const form = model.sheets[0]!;
    const notes = form.validations.find((rule) => rule.odsName === "val11")!;
    expect(notes.ranges).toEqual([range("C2:C1000")]);
    const stored = form.rows.flatMap((row) => row.cells.filter((cell) => cell.col === 2 && row.index >= 1 && row.index <= 999));
    expect(stored).toEqual([]);
    // Nothing past the labelled rows is materialised.
    expect(Math.max(...form.rows.map((row) => row.index))).toBeLessThan(11);
    expect(form.rows.some((row) => row.cells.some((cell) => cell.preservedAttrs !== undefined))).toBe(false);
  });

  it("keeps the orphan verbatim and the error macro's rule enforces nothing", async () => {
    const model = await readSpreadsheetBytes(await bytes, "ods", "t");
    expect(model.odsOrphanValidations).toEqual([O.orphan]);
    const macro = model.sheets[0]!.validations.find((rule) => rule.odsName === "val10")!;
    expect(macro.error.show).toBe(false);
    expect(macro.sourceXml).toContain(`<table:error-macro table:execute="true"/><office:event-listeners>`);
  });

  it("untouched, writes every part back byte-identical", async () => {
    const source = await unzip(await bytes);
    const out = await unzip(await writeOds(await readSpreadsheetBytes(await bytes, "ods", "t")));
    expect([...out.keys()].toSorted()).toEqual([...source.keys()].toSorted());
    for (const [name, data] of source) {
      expect(equalBytes(out.get(name)!, data), name).toBe(true);
    }
  });

  it("edited, regenerates the block with every definition and the orphan kept verbatim, and reads back the same rules", async () => {
    const model = await readSpreadsheetBytes(await bytes, "ods", "t");
    setCellEntry(model, 0, { row: 3, col: 1 }, { value: 2.5 });
    const written = await writeOds(model);
    const content = decoder.decode((await unzip(written)).get("content.xml"));
    expect(content).toContain(`<table:content-validations>${O.definitions.join("")}${O.orphan}</table:content-validations>`);
    expect(content).toContain(
      `<table:table-row table:style-name="ro1" table:number-rows-repeated="${O.notesRun.rows}"><table:table-cell table:number-columns-repeated="2"/>` +
        `<table:table-cell table:content-validation-name="val11"/><table:table-cell table:number-columns-repeated="16381"/></table:table-row>`,
    );
    const reread = await readSpreadsheetBytes(written, "ods", "t2");
    expect(reread.sheets[0]!.validations.map(shape)).toEqual(expected);
    expect(reread.odsOrphanValidations).toEqual([O.orphan]);
  });
});
