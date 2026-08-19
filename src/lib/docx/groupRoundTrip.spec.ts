// SPEC-057 legacy migration + native round trip (.docx side). word-shapes.docx
// is the shipped SPEC-051 fixture; its grouped pair carries the private
// `de:group` attribute, which now migrates to one native DrawingGroup
// container on open (AC-65), and a save-then-reopen keeps the container
// (AC-66).
import { describe, it, expect } from "vitest";
import { readFileSync } from "node:fs";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import { parseDocx, type ParsedDocxDocument } from "./reader";
import { writeDocx } from "./writer";
import { equalStructural, firstStructuralDifference } from "../equality";
import { createDocument, isChart, isDrawingGroup, type Chart, type DocumentModel, type DrawingGroup, type Shape } from "../documentModel";

const fixturesDir = join(dirname(fileURLToPath(import.meta.url)), "fixtures");

function loadFixture(name: string): Uint8Array {
  return new Uint8Array(readFileSync(join(fixturesDir, name)));
}

function modelFrom(parsed: ParsedDocxDocument): DocumentModel {
  return {
    id: "group-round-trip",
    title: "Shapes",
    sections: parsed.sections,
    defaultStyles: parsed.defaultStyles,
    format: "docx",
    opaque: parsed.opaque,
    sourceXml: parsed.sourceXml,
    media: parsed.media,
  };
}

describe("SPEC-057 legacy migration — word-shapes.docx (AC-2, AC-65)", () => {
  it("opens the de:group pair as one DrawingGroup container", async () => {
    const parsed = await parseDocx(loadFixture("word-shapes.docx"));
    const blocks = parsed.sections.flatMap((section) => section.blocks);
    const groups = blocks.filter(isDrawingGroup);
    expect(groups).toHaveLength(1);
    expect(groups[0].members.length).toBeGreaterThanOrEqual(2);
    // No stray legacy ordinal survives migration.
    const anyOrdinal = blocks.some((block) => (block as { groupOrdinal?: number }).groupOrdinal !== undefined);
    expect(anyOrdinal).toBe(false);
  });

  it("no longer emits the private de:group attribute for the container's members on save (AC-1)", async () => {
    const parsed = await parseDocx(loadFixture("word-shapes.docx"));
    const model = modelFrom(parsed);
    const { bytes } = await writeDocx(model);
    const { unzip } = await import("../core/zip");
    const parts = await unzip(bytes);
    const documentXml = new TextDecoder().decode(parts.get("word/document.xml")!);
    // The container writes as a native wpg:wgp, never the private de:group.
    expect(documentXml).toContain("wordprocessingGroup");
    expect(documentXml).not.toContain("de:group");
  });

  it("save then reopen keeps the container — the migration test (AC-66)", async () => {
    const opened = modelFrom(await parseDocx(loadFixture("word-shapes.docx")));
    const { bytes } = await writeDocx(opened);
    const reopened = modelFrom(await parseDocx(bytes));
    const reopenedGroups = reopened.sections.flatMap((s) => s.blocks).filter(isDrawingGroup);
    expect(reopenedGroups).toHaveLength(1);
    const diff = firstStructuralDifference(opened, reopened);
    expect(diff, diff ? JSON.stringify(diff) : "").toBeNull();
    expect(equalStructural(opened, reopened)).toBe(true);
  });
});

describe("SPEC-057 native chart member round trip — wpg:graphicFrame (AC-18, AC-55, AC-59)", () => {
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
    return { ...base, format: "docx", sections: [{ blocks: [{ type: "paragraph", runs: [] }, group] }] };
  }

  it("writes a chart member as a wpg:graphicFrame with its own chart part (AC-59)", async () => {
    const { bytes } = await writeDocx(chartMemberModel());
    const { unzip } = await import("../core/zip");
    const parts = await unzip(bytes);
    const documentXml = new TextDecoder().decode(parts.get("word/document.xml")!);
    expect(documentXml).toContain("wordprocessingGroup");
    expect(documentXml).toContain("graphicFrame");
    // A chart part was minted for the member.
    const chartPart = [...parts.keys()].find((k) => /^word\/charts\/chart\d+\.xml$/.test(k));
    expect(chartPart).toBeTruthy();
  });

  it("re-opens the chart member as a first-class Chart member of the container (AC-55)", async () => {
    const { bytes } = await writeDocx(chartMemberModel());
    const reopened = await parseDocx(bytes);
    const groups = reopened.sections.flatMap((s) => s.blocks).filter(isDrawingGroup);
    expect(groups).toHaveLength(1);
    const chartMember = groups[0].members.find(isChart);
    expect(chartMember).toBeDefined();
    expect(chartMember!.chartType).toBe("columnClustered");
    expect(chartMember!.categories).toEqual(["A", "B"]);
    expect(chartMember!.offsetXCm).toBeCloseTo(6, 1);
    // The container is not sealed — a modelled chart member does not seal it.
    expect(groups[0].preservedSubtreeXml).toBeUndefined();
  });
});
