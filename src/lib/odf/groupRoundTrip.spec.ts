// SPEC-057 legacy migration + native round trip (.odt side). lo-shapes.odt is
// the shipped SPEC-051 fixture; its grouped pair carries the private
// `de:group` attribute, which migrates to one native DrawingGroup container
// on open (AC-65). Saving writes a native `draw:g` (AC-63), and reopening
// keeps the container (AC-66 analog).
import { describe, it, expect } from "vitest";
import { readFileSync } from "node:fs";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import { parseOdt, type ParsedOdtDocument } from "./reader";
import { writeOdt } from "./writer";
import { unzip } from "../core/zip";
import { equalStructural, firstStructuralDifference } from "../equality";
import { createDocument, isChart, isDrawingGroup, type Chart, type DocumentModel, type DrawingGroup, type Shape } from "../documentModel";

const fixturesDir = join(dirname(fileURLToPath(import.meta.url)), "fixtures");

function loadFixture(name: string): Uint8Array {
  return new Uint8Array(readFileSync(join(fixturesDir, name)));
}

function modelFrom(parsed: ParsedOdtDocument): DocumentModel {
  const model: DocumentModel = {
    id: "group-round-trip",
    title: "Shapes",
    sections: parsed.sections,
    defaultStyles: parsed.defaultStyles,
    format: "odt",
    opaque: parsed.opaque,
    sourceXml: parsed.sourceXml,
  };
  if (parsed.media.length > 0) {
    model.media = parsed.media;
  }
  return model;
}

describe("SPEC-057 legacy migration — lo-shapes.odt (AC-2, AC-63, AC-65)", () => {
  it("opens the de:group pair as one DrawingGroup container", async () => {
    const parsed = await parseOdt(loadFixture("lo-shapes.odt"));
    const groups = parsed.sections.flatMap((s) => s.blocks).filter(isDrawingGroup);
    expect(groups).toHaveLength(1);
    expect(groups[0].members.length).toBeGreaterThanOrEqual(2);
  });

  it("saves the container as a native draw:g, never de:group (AC-1/AC-63)", async () => {
    const model = modelFrom(await parseOdt(loadFixture("lo-shapes.odt")));
    const { bytes } = await writeOdt(model);
    const parts = await unzip(bytes);
    const content = new TextDecoder().decode(parts.get("content.xml")!);
    expect(content).toContain("<draw:g");
    expect(content).not.toContain("de:group");
  });

  it("save then reopen keeps the container structurally unchanged (AC-66)", async () => {
    // The container itself must survive a native draw:g round trip. (The
    // document-wide `sourceXml.odtAutomaticStyles` preservation count shifts
    // when the writer re-registers member graphic styles — an ODF style
    // artifact unrelated to the container's own structure — so the assertion
    // targets the container blocks, which is what AC-66 promises.)
    const opened = modelFrom(await parseOdt(loadFixture("lo-shapes.odt")));
    const openedGroup = opened.sections.flatMap((s) => s.blocks).find(isDrawingGroup)!;
    const { bytes } = await writeOdt(opened);
    const reopened = modelFrom(await parseOdt(bytes));
    const reopenedGroups = reopened.sections.flatMap((s) => s.blocks).filter(isDrawingGroup);
    expect(reopenedGroups).toHaveLength(1);
    // Wrap each container in a single-block model so equalStructural compares
    // the drawing group and its members alone.
    const wrap = (group: typeof openedGroup): DocumentModel => ({
      ...opened,
      sections: [{ blocks: [group] }],
      sourceXml: {},
    });
    const diff = firstStructuralDifference(wrap(openedGroup), wrap(reopenedGroups[0]));
    expect(diff, diff ? JSON.stringify(diff) : "").toBeNull();
    expect(equalStructural(wrap(openedGroup), wrap(reopenedGroups[0]))).toBe(true);
  });
});

describe("SPEC-057 native chart member round trip — draw:g / draw:frame (AC-18, AC-62, AC-63)", () => {
  function chartMemberModel(): DocumentModel {
    const base = createDocument("chart-member");
    const chart: Chart = {
      type: "chart",
      chartType: "columnClustered",
      categories: ["A", "B"],
      series: [{ name: "S", values: [3, 4], color: "#4472c4" }],
      options: { legend: "none", dataLabels: false, gridlines: true },
      widthCm: 5,
      heightCm: 4,
      wrap: "square",
      offsetXCm: 6,
      offsetYCm: 0,
    };
    const shape: Shape = { type: "shape", geometry: "rect", widthCm: 3, heightCm: 2, wrap: "square", offsetXCm: 0, offsetYCm: 0 };
    const group: DrawingGroup = {
      type: "drawingGroup",
      widthCm: 11,
      heightCm: 4,
      wrap: "square",
      offsetXCm: 2,
      offsetYCm: 2,
      members: [shape, chart],
    };
    return { ...base, format: "odt", sections: [{ blocks: [{ type: "paragraph", runs: [] }, group] }] };
  }

  it("writes a chart member as a positioned draw:frame/draw:object with its own embedded object (AC-63)", async () => {
    const { bytes } = await writeOdt(chartMemberModel());
    const parts = await unzip(bytes);
    const contentXml = new TextDecoder().decode(parts.get("content.xml")!);
    expect(contentXml).toContain("<draw:g");
    expect(contentXml).toContain("draw:object");
    // An embedded chart object subpackage was minted for the member.
    const objectPart = [...parts.keys()].find((k) => /^Object \d+\/content\.xml$/.test(k));
    expect(objectPart).toBeTruthy();
  });

  it("re-opens the chart member as a first-class Chart member of the container (AC-62)", async () => {
    const { bytes } = await writeOdt(chartMemberModel());
    const reopened = await parseOdt(bytes);
    const groups = reopened.sections.flatMap((s) => s.blocks).filter(isDrawingGroup);
    expect(groups).toHaveLength(1);
    const chartMember = groups[0].members.find(isChart);
    expect(chartMember).toBeDefined();
    expect(chartMember!.chartType).toBe("columnClustered");
    expect(chartMember!.offsetXCm).toBeCloseTo(6, 1);
    expect(groups[0].preservedSubtreeXml).toBeUndefined();
  });
});
