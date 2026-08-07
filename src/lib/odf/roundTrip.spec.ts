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
import { equalBytes, equalStructural, firstStructuralDifference } from "../equality";
import { base64ToBytes } from "../opaque";
import type { DocumentModel } from "../documentModel";
import { asParagraph, createDocument } from "../documentModel";

const fixturesDir = join(dirname(fileURLToPath(import.meta.url)), "fixtures");

function loadFixture(name: string): Uint8Array {
  return new Uint8Array(readFileSync(join(fixturesDir, name)));
}

/** Unlike SPEC-007's modelFrom, this carries the retained source package —
 * the model an opened document actually holds (see openDocument.ts). */
function modelFrom(parsed: ParsedOdtDocument, title = "Field Report"): DocumentModel {
  const model: DocumentModel = {
    id: "test-doc",
    title,
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

const decoder = new TextDecoder();

describe("parseOdt — opaque capture (AC-1, AC-3, AC-4)", () => {
  const fixture = loadFixture("lo-embedded.odt");

  it("retains every unmodelled package part", async () => {
    const parsed = await parseOdt(fixture);
    const names = parsed.opaque.map((part) => part.name);
    expect(names).toContain("styles.xml");
    expect(names).toContain("mimetype");
    expect(names).toContain("META-INF/manifest.xml");
    expect(names).toContain("settings.xml");
    // content.xml and meta.xml are rebuilt on save, not retained; the image
    // (Pictures/image1.png) is modelled since SPEC-024 and travels in
    // `media` instead, not `opaque` (see the "models the draw:frame" test).
    expect(names).not.toContain("content.xml");
    expect(names).not.toContain("meta.xml");
    expect(names).not.toContain("Pictures/image1.png");
    const original = await unzip(fixture);
    for (const part of parsed.opaque) {
      expect(equalBytes(base64ToBytes(part.data), original.get(part.name)!)).toBe(true);
    }
  });

  it("models the draw:frame as an Image block, splitting the paragraph around it (SPEC-024)", async () => {
    const parsed = await parseOdt(fixture);
    // The paragraph "Photo: <frame> shows the site." splits into three
    // blocks — the image addresses its own block, the same way SPEC-022's
    // table does, rather than an opaque inline anchored between runs.
    const before = asParagraph(parsed.sections[0].blocks[2]);
    expect(before.runs.map((run) => run.text)).toEqual(["Photo: "]);
    const image = parsed.sections[0].blocks[3];
    expect(image.type).toBe("image");
    if (image.type !== "image") {
      return;
    }
    expect(image.wrap).toBe("inline");
    expect(image.widthCm).toBeCloseTo(2, 5);
    expect(image.heightCm).toBeCloseTo(2, 5);
    expect(image.mediaId).toBe("Pictures/image1.png");
    const after = asParagraph(parsed.sections[0].blocks[4]);
    expect(after.runs.map((run) => run.text)).toEqual(["shows the site."]);
    // The media bytes are modelled, not left opaque.
    expect(parsed.media).toHaveLength(1);
    expect(parsed.media[0].mediaId).toBe("Pictures/image1.png");
    expect(parsed.media[0].contentType).toBe("image/png");
    const original = await unzip(fixture);
    expect(equalBytes(base64ToBytes(parsed.media[0].data), original.get("Pictures/image1.png")!)).toBe(true);
  });

  it("models the table as a Table block anchored in the body (SPEC-022)", async () => {
    const parsed = await parseOdt(fixture);
    // Blocks: heading, intro, "Photo: ", image, "shows the site.", table,
    // closing — the table sits before the closing paragraph, as a modelled
    // block since SPEC-022 (previously preserved as an opaque node); the
    // photo paragraph now splits around a modelled Image block (SPEC-024).
    expect(parsed.sections[0].blocks).toHaveLength(7);
    expect(parsed.sections[0].opaqueBlocks ?? []).toHaveLength(0);
    const table = parsed.sections[0].blocks[5];
    expect(table.type).toBe("table");
    if (table.type !== "table") {
      return;
    }
    expect(table.rows).toHaveLength(1);
    // The fixture declares its two columns with number-columns-repeated.
    expect(table.columnWidthsCm).toHaveLength(2);
    expect(table.odtStyleName).toBe("Table1");
    expect(
      table.rows.map((row) =>
        row.cells.map((cell) => cell.blocks.map((p) => p.runs.map((r) => r.text).join("")).join("")),
      ),
    ).toEqual([["Cell A", "Cell B"]]);
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
    const intro = asParagraph(model.sections[0].blocks[1]);
    expect(intro.runs[0].text).toBe("Editable intro paragraph.");
    intro.runs[0].text = "Edited intro paragraph!";
    const { bytes } = await writeOdt(model);

    const savedParts = await unzip(bytes);
    const originalParts = await unzip(fixture);
    expect(
      equalBytes(savedParts.get("Pictures/image1.png")!, originalParts.get("Pictures/image1.png")!),
    ).toBe(true);

    const reopened = await parseOdt(bytes);
    expect(asParagraph(reopened.sections[0].blocks[1]).runs[0].text).toBe("Edited intro paragraph!");
    // The table block (modelled since SPEC-022) survives structurally
    // untouched — same rows, columns, widths and style references.
    expect(firstStructuralDifference(opened.sections[0].blocks[5], reopened.sections[0].blocks[5])).toBeNull();
    // The image block (modelled since SPEC-024) survives the same way.
    expect(firstStructuralDifference(opened.sections[0].blocks[3], reopened.sections[0].blocks[3])).toBeNull();

    const editedBack = structuredClone(reopened);
    asParagraph(editedBack.sections[0].blocks[1]).runs[0].text = "Editable intro paragraph.";
    expect(firstStructuralDifference(opened, editedBack)).toBeNull();
  });
});

describe("writeOdt — special characters round-trip (SPEC-026 AC-46)", () => {
  // The canonical fixture: "a b c", NON-BREAKING HYPHEN (U+2011), "d", SOFT
  // HYPHEN (U+00AD), "e", MATHEMATICAL ITALIC SMALL X (U+1D465), "f", GRINNING
  // FACE (U+1F600), "g" — every character representable in ODF.
  const SPECIAL_FIXTURE = "a b c‑d­e\u{1D465}f\u{1F600}g";

  function specialModel(text: string): DocumentModel {
    const model = createDocument("spec-026");
    model.sections[0].blocks = [{ type: "paragraph", runs: [{ text, formatting: {} }] }];
    return model;
  }

  it("survives open -> save -> open equalStructural", async () => {
    // The first serialise+parse yields the canonical opened model; the fixture
    // text must survive it intact…
    const opened = await parseOdt((await writeOdt(specialModel(SPECIAL_FIXTURE))).bytes);
    expect(asParagraph(opened.sections[0].blocks[0]).runs.map((run) => run.text).join("")).toBe(
      SPECIAL_FIXTURE,
    );
    // …and a further save+open is structurally identical.
    const reopened = await parseOdt((await writeOdt(modelFrom(opened))).bytes);
    expect(equalStructural(opened.sections, reopened.sections)).toBe(true);
  });
});

describe("writeOdt — sections through a same-format save (SPEC-028 AC-33)", () => {
  const modified = new Date("2026-08-03T12:34:56Z");

  it("patches the generated master page into the PRESERVED styles.xml; the next save is byte-stable", async () => {
    // Split an opened fixture into two sections — the same-format save must
    // keep every preserved part (SPEC-008) while adding the SectionMaster2
    // definition content.xml now references.
    const opened = await parseOdt(loadFixture("libreoffice-writer.odt"));
    const model = modelFrom(opened, "Sectioned");
    const blocks = model.sections[0].blocks;
    model.sections = [{ blocks: blocks.slice(0, 4) }, { blocks: blocks.slice(4) }];

    const first = await writeOdt(model, { modified });
    const firstParts = await unzip(first.bytes);
    const stylesXml = new TextDecoder().decode(firstParts.get("styles.xml")!);
    // Patched, not rebuilt: the source's named styles survive alongside the
    // new pair.
    expect(stylesXml).toContain('style:name="Heading_20_1"');
    expect(stylesXml).toContain('style:name="SectionMaster2"');
    expect(stylesXml).toContain('style:name="pm2"');
    expect(new TextDecoder().decode(firstParts.get("content.xml")!)).toContain(
      'style:master-page-name="SectionMaster2"',
    );

    const reopened = await parseOdt(first.bytes);
    expect(reopened.sections).toHaveLength(2);
    // The section had no geometry of its own, so the master pinned the then-
    // effective document defaults — explicit on re-open.
    expect(reopened.sections[1].pageSize).toEqual(reopened.defaultStyles.pageSize);
    expect(reopened.sections[1].margins).toEqual(reopened.defaultStyles.margins);

    // Second save: the preserved styles.xml already defines SectionMaster2
    // exactly, so the patch no-ops and the part stays byte-identical.
    const second = await writeOdt(modelFrom(reopened, "Sectioned"), { modified });
    const secondParts = await unzip(second.bytes);
    expect(equalBytes(secondParts.get("styles.xml")!, firstParts.get("styles.xml")!)).toBe(true);
    const reopenedAgain = await parseOdt(second.bytes);
    expect(firstStructuralDifference(reopened.sections, reopenedAgain.sections)).toBeNull();
  });

  it("never displaces a FOREIGN page-layout that shares the generated name (pm2 — SPEC-008)", async () => {
    // OpenOffice/older-LibreOffice number automatic page-layouts pm1, pm2, …
    // so a foreign doc's second section is literally "pm2", carrying content
    // this writer does not model. A same-format save must keep that
    // definition byte-for-byte and give the factory's own section a
    // non-colliding layout name instead.
    const { zip } = await import("../docx/zipWriter");
    const makeContent = () => `<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" office:version="1.2">
 <office:automatic-styles><style:style style:name="P1" style:family="paragraph" style:master-page-name="A5Master"/></office:automatic-styles>
 <office:body><office:text><text:p>first</text:p><text:p text:style-name="P1">second</text:p></office:text></office:body>
</office:document-content>`;
    const makeStyles = () => `<?xml version="1.0" encoding="UTF-8"?>
<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" office:version="1.2">
 <office:automatic-styles>
  <style:page-layout style:name="pm1"><style:page-layout-properties fo:page-width="21cm" fo:page-height="29.7cm"/></style:page-layout>
  <style:page-layout style:name="pm2"><style:page-layout-properties fo:page-width="14.8cm" fo:page-height="21cm"><style:background-image/></style:page-layout-properties></style:page-layout>
 </office:automatic-styles>
 <office:master-styles>
  <style:master-page style:name="Standard" style:page-layout-name="pm1"/>
  <style:master-page style:name="A5Master" style:page-layout-name="pm2"/>
 </office:master-styles>
</office:document-styles>`;
    const opened = await parseOdt(
      await zip([
        { name: "mimetype", data: new TextEncoder().encode("application/vnd.oasis.opendocument.text"), store: true },
        { name: "content.xml", data: new TextEncoder().encode(makeContent()) },
        { name: "styles.xml", data: new TextEncoder().encode(makeStyles()) },
      ]),
    );
    const model = modelFrom(opened, "Collision");
    expect(model.sections).toHaveLength(2);

    const { bytes } = await writeOdt(model, { modified });
    const parts = await unzip(bytes);
    const stylesXml = decoder.decode(parts.get("styles.xml")!);
    // The foreign pm2 — with its unmodeled background-image — survives whole.
    expect(stylesXml).toContain(
      '<style:page-layout style:name="pm2"><style:page-layout-properties fo:page-width="14.8cm"' +
        ' fo:page-height="21cm"><style:background-image/></style:page-layout-properties></style:page-layout>',
    );
    // The factory's own section gets a non-colliding layout name…
    expect(stylesXml).toContain('style:name="pm2-2"');
    expect(stylesXml).toContain('<style:master-page style:name="SectionMaster2" style:page-layout-name="pm2-2"/>');
    // …and content.xml references the factory's master page consistently.
    expect(decoder.decode(parts.get("content.xml")!)).toContain('style:master-page-name="SectionMaster2"');
  });

  it("replaces a stale generated master page when the section's geometry was edited", async () => {
    const opened = await parseOdt(loadFixture("libreoffice-writer.odt"));
    const model = modelFrom(opened, "Sectioned");
    const blocks = model.sections[0].blocks;
    model.sections = [{ blocks: blocks.slice(0, 4) }, { blocks: blocks.slice(4) }];
    const reopened = await parseOdt((await writeOdt(model, { modified })).bytes);

    const edited = modelFrom(reopened, "Sectioned");
    edited.sections[1].pageSize = { widthCm: 14.8, heightCm: 21 };
    const { bytes } = await writeOdt(edited, { modified });
    const stylesXml = new TextDecoder().decode((await unzip(bytes)).get("styles.xml")!);
    expect(stylesXml).toContain('fo:page-width="14.8cm"');

    const reopenedEdited = await parseOdt(bytes);
    expect(reopenedEdited.sections[1].pageSize).toEqual({ widthCm: 14.8, heightCm: 21 });
  });
});

describe("writeDocx — converting a .odt model reports what it drops (SPEC-008 writer API)", () => {
  const fixture = loadFixture("lo-embedded.odt");

  it("lists the table with a location, skipping packaging parts and the image (converts losslessly)", async () => {
    const opened = await parseOdt(fixture);
    const { bytes, dropped } = await writeDocx(modelFrom(opened));

    // The table converts as a modelled block since SPEC-022 — it is no
    // longer reported as dropped content.
    expect(dropped.some((d) => d.kind === "opaqueBlock")).toBe(false);
    // The image converts as a modelled block since SPEC-024 (a PNG has a
    // faithful .docx representation, so nothing is dropped for it either).
    expect(dropped.some((d) => d.kind === "opaqueInline")).toBe(false);
    expect(dropped.some((d) => d.kind === "svgImage")).toBe(false);
    expect(dropped.some((d) => d.kind === "opaquePart" && d.location.includes("Pictures/image1.png"))).toBe(false);
    // Packaging infrastructure is not reported as dropped content.
    expect(dropped.some((d) => d.location.includes("mimetype"))).toBe(false);
    expect(dropped.some((d) => d.location.includes("styles.xml"))).toBe(false);
    expect(dropped.some((d) => d.location.includes("META-INF"))).toBe(false);

    // The converted package carries no ODF opaque content, but does carry
    // the converted image as a .docx media part.
    const parts = await unzip(bytes);
    expect(parts.has("Pictures/image1.png")).toBe(false);
    expect(parts.has("word/media/image1.png")).toBe(true);
    const documentXml = decoder.decode(parts.get("word/document.xml")!);
    expect(documentXml).not.toContain("draw:frame");
    expect(documentXml).not.toContain("table:table");
    expect(documentXml).toContain("<w:drawing");
    // The modelled table converts to WordprocessingML (SPEC-022 AC-24).
    expect(documentXml).toContain("<w:tbl>");
    expect(documentXml).toContain("Cell A");
    expect(documentXml).toContain("Photo:");
  });
});
