import { describe, expect, it } from "vitest";
import { readPresentationBytes } from "../open";
import { buildOdpTablesFixtureBytes, buildPptxTablesFixtureBytes, EXPECTED_FIXTURE_TABLE } from "./build";
import type { PresentationModel } from "../model";

function assertFixture(model: PresentationModel, expectBand: boolean): void {
  const table = model.slides[0].frames.find((f) => f.kind === "table");
  expect(table).toBeDefined();
  if (table && table.kind === "table") {
    expect(table.rows).toHaveLength(EXPECTED_FIXTURE_TABLE.rows);
    expect(table.columnWidthsCm).toHaveLength(EXPECTED_FIXTURE_TABLE.cols);
    expect(table.firstRow).toBe(true);
    // bandRow is a pptx look flag (a:tblPr@bandRow); ODF has no structural
    // banded-rows attribute, so it does not survive an odp round trip.
    if (expectBand) {
      expect(table.bandRow).toBe(true);
    }
    const header = table.rows[0].cells.map((c) => c.paragraphs[0]?.runs[0]?.text);
    expect(header).toEqual(EXPECTED_FIXTURE_TABLE.headerText);
    expect(table.rows[1].cells[0].paragraphs[0].runs[0].text).toBe(EXPECTED_FIXTURE_TABLE.bodyFirstCell);
  }
  // The deliberately out-of-subset table stays a foreign "Table" frame (AC-17c).
  const foreign = model.slides[0].frames.find((f) => f.kind === "foreign");
  expect(foreign?.kind === "foreign" && foreign.label).toBe("Table");
}

describe("SPEC-091 table fixtures (AC-17)", () => {
  it("pptx: opens to the modelled table + a preserved out-of-subset table", async () => {
    const model = await readPresentationBytes(await buildPptxTablesFixtureBytes(), "pptx", "Tables");
    assertFixture(model, true);
  });

  it("odp: opens to the modelled table + a preserved out-of-subset table", async () => {
    const model = await readPresentationBytes(await buildOdpTablesFixtureBytes(), "odp", "Tables");
    assertFixture(model, false);
  });
});
