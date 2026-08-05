// SPEC-008 — round-trip preservation of unsupported content (.docx side).
// The fixture word-charts-and-controls.docx carries the three things the
// spec's test AC names: a DrawingML chart (part + inline w:drawing run), a
// block-level w:sdt content control and a customXml part.
import { describe, it, expect } from "vitest";
import { readFileSync } from "node:fs";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import { parseDocx, type ParsedDocxDocument } from "./reader";
import { writeDocx } from "./writer";
import { writeOdt } from "../odf/writer";
import { unzip } from "./zip";
import { canonicalizeXml, equalBytes, firstStructuralDifference } from "../equality";
import { base64ToBytes } from "../opaque";
import type { DocumentModel, Paragraph } from "../documentModel";
import { asParagraph } from "../documentModel";

const fixturesDir = join(dirname(fileURLToPath(import.meta.url)), "fixtures");

function loadFixture(name: string): Uint8Array {
  return new Uint8Array(readFileSync(join(fixturesDir, name)));
}

/** Unlike SPEC-006's modelFrom, this carries the retained source package —
 * the model an opened document actually holds (see openDocument.ts). */
function modelFrom(parsed: ParsedDocxDocument, title = "Annual Metrics"): DocumentModel {
  return {
    id: "test-doc",
    title,
    sections: parsed.sections,
    defaultStyles: parsed.defaultStyles,
    format: "docx",
    opaque: parsed.opaque,
    sourceXml: parsed.sourceXml,
  };
}

const decoder = new TextDecoder();

describe("parseDocx — opaque capture (AC-1, AC-3, AC-4)", () => {
  const fixture = loadFixture("word-charts-and-controls.docx");

  it("retains every unmodelled package part for the lifetime of the document", async () => {
    const parsed = await parseDocx(fixture);
    const names = parsed.opaque.map((part) => part.name);
    expect(names).toContain("word/charts/chart1.xml");
    expect(names).toContain("customXml/item1.xml");
    expect(names).toContain("customXml/itemProps1.xml");
    expect(names).toContain("word/styles.xml");
    expect(names).toContain("[Content_Types].xml");
    // The main document part and core properties are rebuilt on save, not
    // retained.
    expect(names).not.toContain("word/document.xml");
    expect(names).not.toContain("docProps/core.xml");
    // Retained bytes are the original part bytes.
    const original = await unzip(fixture);
    for (const part of parsed.opaque) {
      expect(equalBytes(base64ToBytes(part.data), original.get(part.name)!)).toBe(true);
    }
  });

  it("captures the chart's w:drawing run as an opaque inline anchored between the runs", async () => {
    const parsed = await parseDocx(fixture);
    const chartParagraph = asParagraph(parsed.sections[0].blocks[2]);
    expect(chartParagraph.runs.map((run) => run.text)).toEqual(["Quarterly results: "]);
    expect(chartParagraph.opaqueInlines).toHaveLength(1);
    expect(chartParagraph.opaqueInlines![0].runIndex).toBe(1);
    expect(chartParagraph.opaqueInlines![0].xml).toContain("drawing");
    expect(chartParagraph.opaqueInlines![0].xml).toContain("chart");
  });

  it("captures the block-level w:sdt content control as an opaque block anchored in the body", async () => {
    const parsed = await parseDocx(fixture);
    // Blocks: heading, intro, chart paragraph, closing — the sdt sits
    // before the closing paragraph.
    expect(parsed.sections[0].blocks).toHaveLength(4);
    expect(parsed.sections[0].opaqueBlocks).toHaveLength(1);
    const control = parsed.sections[0].opaqueBlocks![0];
    expect(control.blockIndex).toBe(3);
    expect(control.xml).toContain("sdt");
    expect(control.xml).toContain("Controlled block text");
  });

  it("captures the original body-level w:sectPr", async () => {
    const parsed = await parseDocx(fixture);
    expect(parsed.sourceXml.docxSectPr).toContain("pgSz");
  });
});

describe("writeDocx — same-format round trip (AC-2, AC-5, AC-6, AC-7)", () => {
  const fixture = loadFixture("word-charts-and-controls.docx");

  it("re-opens equalStructural after an unedited save", async () => {
    const opened = await parseDocx(fixture);
    const { bytes, dropped } = await writeDocx(modelFrom(opened));
    expect(dropped).toEqual([]);
    const reopened = await parseDocx(bytes);
    expect(firstStructuralDifference(opened, reopened)).toBeNull();
  });

  it("writes every preserved opaque part back byte-identical (equalBytes)", async () => {
    const opened = await parseDocx(fixture);
    const { bytes } = await writeDocx(modelFrom(opened));
    const savedParts = await unzip(bytes);
    const originalParts = await unzip(fixture);
    expect(opened.opaque.length).toBeGreaterThan(0);
    for (const part of opened.opaque) {
      expect(savedParts.has(part.name), `${part.name} should be written back`).toBe(true);
      expect(
        equalBytes(savedParts.get(part.name)!, originalParts.get(part.name)!),
        `${part.name} should be byte-identical`,
      ).toBe(true);
    }
  });

  it("re-emits opaque nodes in place: drawing inside its paragraph, sdt before the closing paragraph", async () => {
    const opened = await parseDocx(fixture);
    const { bytes } = await writeDocx(modelFrom(opened));
    const documentXml = decoder.decode((await unzip(bytes)).get("word/document.xml")!);

    const chartParagraphStart = documentXml.indexOf("Quarterly results:");
    const drawing = documentXml.indexOf("<w:drawing");
    const sdt = documentXml.indexOf("<w:sdt");
    const closing = documentXml.indexOf("Closing remarks.");
    expect(chartParagraphStart).toBeGreaterThan(-1);
    // The drawing run follows its paragraph's text run; the content control
    // sits between the chart paragraph and the closing paragraph.
    expect(drawing).toBeGreaterThan(chartParagraphStart);
    expect(sdt).toBeGreaterThan(drawing);
    expect(closing).toBeGreaterThan(sdt);
  });

  it("emits the original w:sectPr verbatim", async () => {
    const opened = await parseDocx(fixture);
    const { bytes } = await writeDocx(modelFrom(opened));
    const documentXml = decoder.decode((await unzip(bytes)).get("word/document.xml")!);
    expect(documentXml).toContain('<w:pgSz w:w="11906" w:h="16838"/>');
  });

  it("adds docProps/core.xml (patching content types and rels) when the source had none", async () => {
    // word-authored.docx ships without docProps/core.xml; the save must
    // still carry the edited title, so the part is added and declared.
    const opened = await parseDocx(loadFixture("word-authored.docx"));
    const { bytes } = await writeDocx(modelFrom(opened, "Renamed"));
    const parts = await unzip(bytes);
    expect(decoder.decode(parts.get("docProps/core.xml")!)).toContain("<dc:title>Renamed</dc:title>");
    expect(decoder.decode(parts.get("[Content_Types].xml")!)).toContain('PartName="/docProps/core.xml"');
    expect(decoder.decode(parts.get("_rels/.rels")!)).toContain("core-properties");
  });
});

describe("writeDocx — editing one paragraph leaves opaque content untouched (AC-8, AC-14)", () => {
  const fixture = loadFixture("word-charts-and-controls.docx");

  async function editIntroAndSave() {
    const opened = await parseDocx(fixture);
    // modelFrom shares the parsed tree by reference; clone so `opened`
    // stays the pristine as-opened model the assertions compare against.
    const model = modelFrom(structuredClone(opened));
    const intro = asParagraph(model.sections[0].blocks[1]);
    expect(intro.runs[0].text).toBe("Editable intro paragraph.");
    intro.runs[0].text = "Edited intro paragraph!";
    const { bytes } = await writeDocx(model);
    return { opened, bytes };
  }

  it("keeps the chart part and the custom XML part byte-identical", async () => {
    const { bytes } = await editIntroAndSave();
    const savedParts = await unzip(bytes);
    const originalParts = await unzip(fixture);
    for (const name of ["word/charts/chart1.xml", "customXml/item1.xml", "customXml/itemProps1.xml"]) {
      expect(equalBytes(savedParts.get(name)!, originalParts.get(name)!), name).toBe(true);
    }
  });

  it("keeps the content control and the chart drawing intact, and only the edited paragraph differs", async () => {
    const { opened, bytes } = await editIntroAndSave();
    const reopened = await parseDocx(bytes);

    // The edit round-trips…
    expect((reopened.sections[0].blocks[1] as Paragraph).runs[0].text).toBe("Edited intro paragraph!");

    // …every opaque node is structurally identical to what was opened…
    expect(canonicalizeXml(reopened.sections[0].opaqueBlocks![0].xml)).toBe(
      canonicalizeXml(opened.sections[0].opaqueBlocks![0].xml),
    );
    expect(canonicalizeXml(asParagraph(reopened.sections[0].blocks[2]).opaqueInlines![0].xml)).toBe(
      canonicalizeXml(asParagraph(opened.sections[0].blocks[2]).opaqueInlines![0].xml),
    );

    // …and so is every other block: the edited paragraph is the only
    // structural difference between the opened and re-opened models.
    const editedBack = structuredClone(reopened);
    (editedBack.sections[0].blocks[1] as Paragraph).runs[0].text = "Editable intro paragraph.";
    expect(firstStructuralDifference(opened, editedBack)).toBeNull();
  });
});

describe("writeOdt — converting a .docx model reports what it drops (SPEC-008 writer API)", () => {
  const fixture = loadFixture("word-charts-and-controls.docx");

  it("lists the chart, the content control and the custom XML part with locations", async () => {
    const opened = await parseDocx(fixture);
    const { bytes, dropped } = await writeOdt(modelFrom(opened));

    const chart = dropped.find((d) => d.kind === "chart" && d.location === "paragraph 3");
    expect(chart, "inline chart at paragraph 3").toBeDefined();
    const control = dropped.find((d) => d.kind === "contentControl");
    expect(control?.location).toBe("before paragraph 4");
    expect(dropped.some((d) => d.kind === "chart" && d.location.includes("word/charts/chart1.xml"))).toBe(true);
    expect(dropped.some((d) => d.kind === "customXml" && d.location.includes("customXml/item1.xml"))).toBe(true);
    for (const record of dropped) {
      expect(record.reason.length).toBeGreaterThan(0);
    }

    // Packaging infrastructure is not reported as dropped content.
    expect(dropped.some((d) => d.location.includes("[Content_Types].xml"))).toBe(false);
    expect(dropped.some((d) => d.location.includes("word/styles.xml"))).toBe(false);

    // The converted package carries no WordprocessingML opaque content.
    const parts = await unzip(bytes);
    expect(parts.has("word/charts/chart1.xml")).toBe(false);
    const contentXml = decoder.decode(parts.get("content.xml")!);
    expect(contentXml).not.toContain("w:sdt");
    expect(contentXml).not.toContain("w:drawing");
    expect(contentXml).toContain("Quarterly results:");
  });

  it("reports nothing for a blank document", async () => {
    const { dropped } = await writeOdt({
      id: "blank",
      title: "Blank",
      sections: [{ blocks: [{ type: "paragraph", runs: [] }] }],
      defaultStyles: (await parseDocx(fixture)).defaultStyles,
    });
    expect(dropped).toEqual([]);
  });
});
