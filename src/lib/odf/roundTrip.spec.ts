// SPEC-008 — round-trip preservation of unsupported content (.odt side).
// The fixture lo-embedded.odt carries an embedded image (Pictures/ part +
// inline draw:frame) and a table:table block, with the graphic/table
// automatic styles those nodes reference.
import { describe, it, expect } from "vitest";
import { readFileSync } from "node:fs";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import { parseOdt, type ParsedOdtDocument } from "./reader";
import { writeOdt } from "./writer";
import { writeDocx } from "../docx/writer";
import { unzip } from "../docx/zip";
import { canonicalizeXml, equalBytes, firstStructuralDifference } from "../equality";
import { base64ToBytes } from "../opaque";
import type { DocumentModel } from "../documentModel";

const fixturesDir = join(dirname(fileURLToPath(import.meta.url)), "fixtures");

function loadFixture(name: string): Uint8Array {
  return new Uint8Array(readFileSync(join(fixturesDir, name)));
}

/** Unlike SPEC-007's modelFrom, this carries the retained source package —
 * the model an opened document actually holds (see openDocument.ts). */
function modelFrom(parsed: ParsedOdtDocument, title = "Field Report"): DocumentModel {
  return {
    id: "test-doc",
    title,
    sections: parsed.sections,
    defaultStyles: parsed.defaultStyles,
    format: "odt",
    opaque: parsed.opaque,
    sourceXml: parsed.sourceXml,
  };
}

const decoder = new TextDecoder();

describe("parseOdt — opaque capture (AC-1, AC-3, AC-4)", () => {
  const fixture = loadFixture("lo-embedded.odt");

  it("retains every unmodelled package part, including binary images", async () => {
    const parsed = await parseOdt(fixture);
    const names = parsed.opaque.map((part) => part.name);
    expect(names).toContain("Pictures/image1.png");
    expect(names).toContain("styles.xml");
    expect(names).toContain("mimetype");
    expect(names).toContain("META-INF/manifest.xml");
    expect(names).toContain("settings.xml");
    // content.xml and meta.xml are rebuilt on save, not retained.
    expect(names).not.toContain("content.xml");
    expect(names).not.toContain("meta.xml");
    const original = await unzip(fixture);
    for (const part of parsed.opaque) {
      expect(equalBytes(base64ToBytes(part.data), original.get(part.name)!)).toBe(true);
    }
  });

  it("captures the draw:frame as an opaque inline anchored between the runs", async () => {
    const parsed = await parseOdt(fixture);
    const photoParagraph = parsed.sections[0].blocks[2];
    expect(photoParagraph.runs.map((run) => run.text)).toEqual(["Photo: ", "shows the site."]);
    expect(photoParagraph.opaqueInlines).toHaveLength(1);
    expect(photoParagraph.opaqueInlines![0].runIndex).toBe(1);
    expect(photoParagraph.opaqueInlines![0].xml).toContain("frame");
    expect(photoParagraph.opaqueInlines![0].xml).toContain("Pictures/image1.png");
  });

  it("captures the table as an opaque block anchored in the body", async () => {
    const parsed = await parseOdt(fixture);
    // Blocks: heading, intro, photo paragraph, closing — the table sits
    // before the closing paragraph.
    expect(parsed.sections[0].blocks).toHaveLength(4);
    expect(parsed.sections[0].opaqueBlocks).toHaveLength(1);
    const table = parsed.sections[0].opaqueBlocks![0];
    expect(table.blockIndex).toBe(3);
    expect(table.xml).toContain("table");
    expect(table.xml).toContain("Cell A");
  });

  it("keeps font-face declarations and the non-paragraph/text automatic styles opaque nodes reference", async () => {
    const parsed = await parseOdt(fixture);
    expect(parsed.sourceXml.odtFontFaceDecls).toContain("Liberation Serif");
    const styles = parsed.sourceXml.odtAutomaticStyles ?? [];
    expect(styles.some((def) => def.includes('"fr1"'))).toBe(true);
    expect(styles.some((def) => def.includes('"Table1"'))).toBe(true);
    // Paragraph-family styles are resolved into the model and regenerated
    // flat on save, so they are not captured.
    expect(styles.some((def) => def.includes('"P1"'))).toBe(false);
  });
});

describe("writeOdt — same-format round trip (AC-2, AC-5, AC-6, AC-7)", () => {
  const fixture = loadFixture("lo-embedded.odt");

  it("re-opens equalStructural after an unedited save", async () => {
    const opened = await parseOdt(fixture);
    const { bytes, dropped } = await writeOdt(modelFrom(opened));
    expect(dropped).toEqual([]);
    const reopened = await parseOdt(bytes);
    expect(firstStructuralDifference(opened, reopened)).toBeNull();
  });

  it("writes every preserved opaque part back byte-identical (equalBytes)", async () => {
    const opened = await parseOdt(fixture);
    const { bytes } = await writeOdt(modelFrom(opened));
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

  it("keeps the package a conformant ODF ZIP: stored mimetype entry first", async () => {
    const { bytes } = await writeOdt(modelFrom(await parseOdt(fixture)));
    const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    expect(view.getUint32(0, true)).toBe(0x04034b50);
    expect(view.getUint16(8, true)).toBe(0); // compression method: STORE
    const nameLength = view.getUint16(26, true);
    expect(decoder.decode(bytes.subarray(30, 30 + nameLength))).toBe("mimetype");
  });

  it("re-emits opaque nodes in place with the preserved styles they reference", async () => {
    const opened = await parseOdt(fixture);
    const { bytes } = await writeOdt(modelFrom(opened));
    const contentXml = decoder.decode((await unzip(bytes)).get("content.xml")!);

    const photo = contentXml.indexOf("Photo:");
    const frame = contentXml.indexOf("<draw:frame");
    const tail = contentXml.indexOf("shows the site.");
    const table = contentXml.indexOf("<table:table");
    const closing = contentXml.indexOf("Closing summary.");
    expect(photo).toBeGreaterThan(-1);
    expect(frame).toBeGreaterThan(photo);
    expect(tail).toBeGreaterThan(frame);
    expect(table).toBeGreaterThan(tail);
    expect(closing).toBeGreaterThan(table);

    // The graphic/table styles and font declarations survive alongside.
    expect(contentXml).toContain('style:name="fr1"');
    expect(contentXml).toContain('style:name="Table1"');
    expect(contentXml).toContain("office:font-face-decls");
  });
});

describe("writeOdt — editing one paragraph leaves opaque content untouched (AC-8)", () => {
  const fixture = loadFixture("lo-embedded.odt");

  it("keeps the image bytes and the table intact; only the edited paragraph differs", async () => {
    const opened = await parseOdt(fixture);
    // modelFrom shares the parsed tree by reference; clone so `opened`
    // stays the pristine as-opened model the assertions compare against.
    const model = modelFrom(structuredClone(opened));
    const intro = model.sections[0].blocks[1];
    expect(intro.runs[0].text).toBe("Editable intro paragraph.");
    intro.runs[0].text = "Edited intro paragraph!";
    const { bytes } = await writeOdt(model);

    const savedParts = await unzip(bytes);
    const originalParts = await unzip(fixture);
    expect(
      equalBytes(savedParts.get("Pictures/image1.png")!, originalParts.get("Pictures/image1.png")!),
    ).toBe(true);

    const reopened = await parseOdt(bytes);
    expect(reopened.sections[0].blocks[1].runs[0].text).toBe("Edited intro paragraph!");
    expect(canonicalizeXml(reopened.sections[0].opaqueBlocks![0].xml)).toBe(
      canonicalizeXml(opened.sections[0].opaqueBlocks![0].xml),
    );
    expect(canonicalizeXml(reopened.sections[0].blocks[2].opaqueInlines![0].xml)).toBe(
      canonicalizeXml(opened.sections[0].blocks[2].opaqueInlines![0].xml),
    );

    const editedBack = structuredClone(reopened);
    editedBack.sections[0].blocks[1].runs[0].text = "Editable intro paragraph.";
    expect(firstStructuralDifference(opened, editedBack)).toBeNull();
  });
});

describe("writeDocx — converting a .odt model reports what it drops (SPEC-008 writer API)", () => {
  const fixture = loadFixture("lo-embedded.odt");

  it("lists the frame, the table and the image part with locations, skipping packaging parts", async () => {
    const opened = await parseOdt(fixture);
    const { bytes, dropped } = await writeDocx(modelFrom(opened));

    expect(dropped.some((d) => d.kind === "opaqueInline" && d.location === "paragraph 3")).toBe(true);
    expect(dropped.some((d) => d.kind === "opaqueBlock" && d.location === "before paragraph 4")).toBe(true);
    expect(dropped.some((d) => d.kind === "opaquePart" && d.location.includes("Pictures/image1.png"))).toBe(true);
    // Packaging infrastructure is not reported as dropped content.
    expect(dropped.some((d) => d.location.includes("mimetype"))).toBe(false);
    expect(dropped.some((d) => d.location.includes("styles.xml"))).toBe(false);
    expect(dropped.some((d) => d.location.includes("META-INF"))).toBe(false);

    // The converted package carries no ODF opaque content.
    const parts = await unzip(bytes);
    expect(parts.has("Pictures/image1.png")).toBe(false);
    const documentXml = decoder.decode(parts.get("word/document.xml")!);
    expect(documentXml).not.toContain("draw:frame");
    expect(documentXml).not.toContain("table:table");
    expect(documentXml).toContain("Photo:");
  });
});
