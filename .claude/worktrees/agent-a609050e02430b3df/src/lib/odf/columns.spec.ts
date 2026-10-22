// SPEC-030 — .odt column round-trip: style:columns with fo:column-count,
// fo:column-gap, style:column rel-widths for unequal columns, and a
// style:column-sep for "line between columns".
import { describe, expect, it } from "vitest";
import { parseOdt } from "./reader";
import { buildStylesXml, writeOdt } from "./writer";
import { createDocument, type DocumentModel, type Section } from "../documentModel";

function modelWith(section: Section): DocumentModel {
  const model = createDocument("spec-030-odt");
  model.sections = [section];
  return model;
}

async function reopen(model: DocumentModel) {
  const { bytes } = await writeOdt(model);
  return parseOdt(bytes);
}

describe("odt columns round-trip (SPEC-030)", () => {
  it("writes style:columns with fo:column-count and fo:column-gap", () => {
    const model = modelWith({
      blocks: [{ type: "paragraph", runs: [{ text: "col", formatting: {} }] }],
      columns: { count: 2, equalWidth: true, gapCm: 1.25, separator: true },
    });
    const xml = buildStylesXml(model);
    expect(xml).toContain('<style:columns fo:column-count="2"');
    expect(xml).toContain("fo:column-gap=");
    expect(xml).toContain("style:column-sep");
  });

  it("writes style:column rel-widths for unequal columns", () => {
    const model = modelWith({
      blocks: [{ type: "paragraph", runs: [{ text: "col", formatting: {} }] }],
      columns: { count: 2, equalWidth: false, gapCm: 1, separator: false, widthsCm: [5, 10] },
    });
    const xml = buildStylesXml(model);
    expect((xml.match(/<style:column /g) ?? []).length).toBe(2);
    expect(xml).toContain("style:rel-width=");
  });

  it("round-trips equal columns", async () => {
    const model = modelWith({
      blocks: [{ type: "paragraph", runs: [{ text: "col", formatting: {} }] }],
      columns: { count: 3, equalWidth: true, gapCm: 1.25, separator: true },
    });
    const reopened = await reopen(model);
    const cols = reopened.sections[0].columns;
    expect(cols?.count).toBe(3);
    expect(cols?.equalWidth).toBe(true);
    expect(cols?.separator).toBe(true);
  });

  it("round-trips unequal columns' proportions", async () => {
    const model = modelWith({
      blocks: [{ type: "paragraph", runs: [{ text: "col", formatting: {} }] }],
      columns: { count: 2, equalWidth: false, gapCm: 1, separator: false, widthsCm: [5, 10] },
    });
    const reopened = await reopen(model);
    const cols = reopened.sections[0].columns;
    expect(cols?.count).toBe(2);
    expect(cols?.equalWidth).toBe(false);
    expect((cols!.widthsCm![1] ?? 0) / (cols!.widthsCm![0] ?? 1)).toBeCloseTo(2, 1);
  });
});
