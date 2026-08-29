// SPEC-091 table round-trip fidelity + byte stability, exercised through the
// mini-deck's slide-4 table (a real graphicFrame / table:table part).
import { describe, expect, it } from "vitest";
import { equalBytes } from "../core/equality";
import { unzip } from "../core/zip";
import { buildOdpFixtureBytes, buildPptxFixtureBytes } from "./fixtures/build";
import { readPresentationBytes } from "./open";
import { writeOdp } from "./odp/writer";
import { writePptx } from "./pptx/writer";
import { markSlideEdited } from "./text";
import { insertTableColumn, setCellFill } from "./tables";
import type { PresentationModel, TableFrame } from "./model";

function tableOf(model: PresentationModel): TableFrame {
  const frame = model.slides[3].frames.find((f) => f.kind === "table");
  if (!frame || frame.kind !== "table") {
    throw new Error("no table on slide 4");
  }
  return frame;
}

describe("SPEC-091 table round trip (pptx)", () => {
  it("no-edit download re-emits every part byte-identical, table included (AC-13a)", async () => {
    const sourceBytes = await buildPptxFixtureBytes();
    const model = await readPresentationBytes(sourceBytes, "pptx", "Deck");
    expect(tableOf(model)).toBeDefined();
    const source = await unzip(sourceBytes);
    const written = await unzip((await writePptx(model)).bytes);
    for (const [name, data] of source) {
      expect(equalBytes(written.get(name)!, data), `part ${name}`).toBe(true);
    }
  });

  it("an edit on slide 4 leaves every other part byte-identical, and modelled fields survive (AC-13b, AC-14)", async () => {
    const sourceBytes = await buildPptxFixtureBytes();
    const source = await unzip(sourceBytes);
    const model = await readPresentationBytes(sourceBytes, "pptx", "Deck");
    const table = tableOf(model);
    table.rows[0].cells[0].paragraphs = [{ runs: [{ text: "Edited" }] }];
    insertTableColumn(table, 0, "right");
    setCellFill(table, { section: 0, block: 0, anchor: { row: 0, col: 1 }, focus: { row: 0, col: 1 } }, { type: "solid", color: "#ff0000" });
    markSlideEdited(model.slides[3], "pptx");

    const { bytes } = await writePptx(model);
    const written = await unzip(bytes);
    for (const [name, data] of source) {
      if (name === "ppt/slides/slide4.xml") continue;
      expect(equalBytes(written.get(name)!, data), `part ${name} must stay byte-verbatim`).toBe(true);
    }

    const reopened = await readPresentationBytes(bytes, "pptx", "Deck");
    const rt = tableOf(reopened);
    expect(rt.columnWidthsCm).toHaveLength(2);
    expect(rt.rows[0].cells[0].paragraphs[0].runs[0].text).toBe("Edited");
    expect(rt.rows[0].cells[1].fill).toEqual({ type: "solid", color: "#ff0000" });

    // Byte stability (AC-13c): rewriting the reopened model re-emits identical
    // slide-part bytes.
    const twice = await unzip((await writePptx(reopened)).bytes);
    expect(equalBytes(twice.get("ppt/slides/slide4.xml")!, written.get("ppt/slides/slide4.xml")!)).toBe(true);
  });
});

describe("SPEC-091 table round trip (odp)", () => {
  it("no-edit download re-emits every part byte-identical, table included (AC-13a)", async () => {
    const sourceBytes = await buildOdpFixtureBytes();
    const model = await readPresentationBytes(sourceBytes, "odp", "Deck");
    expect(tableOf(model)).toBeDefined();
    const source = await unzip(sourceBytes);
    const written = await unzip((await writeOdp(model)).bytes);
    for (const [name, data] of source) {
      expect(equalBytes(written.get(name)!, data), `part ${name}`).toBe(true);
    }
  });

  it("an edit round-trips the modelled grid, text and fill (AC-14)", async () => {
    const sourceBytes = await buildOdpFixtureBytes();
    const model = await readPresentationBytes(sourceBytes, "odp", "Deck");
    const table = tableOf(model);
    table.rows[0].cells[0].paragraphs = [{ runs: [{ text: "Edited" }] }];
    insertTableColumn(table, 0, "right");
    setCellFill(table, { section: 0, block: 0, anchor: { row: 0, col: 1 }, focus: { row: 0, col: 1 } }, { type: "solid", color: "#ff0000" });
    markSlideEdited(model.slides[3], "odp");

    const { bytes } = await writeOdp(model);
    const reopened = await readPresentationBytes(bytes, "odp", "Deck");
    const rt = tableOf(reopened);
    expect(rt.columnWidthsCm).toHaveLength(2);
    expect(rt.rows[0].cells[0].paragraphs[0].runs[0].text).toBe("Edited");
    expect(rt.rows[0].cells[1].fill).toEqual({ type: "solid", color: "#ff0000" });

    // Byte stability: rewriting the reopened model re-emits identical content.xml.
    const first = await unzip(bytes);
    const twice = await unzip((await writeOdp(reopened)).bytes);
    expect(equalBytes(twice.get("content.xml")!, first.get("content.xml")!)).toBe(true);
  });
});
