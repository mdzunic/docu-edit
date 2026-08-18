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
import { isDrawingGroup, type DocumentModel } from "../documentModel";

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
