import { describe, it, expect } from "vitest";
import { readFileSync } from "node:fs";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import {
  openDocumentFile,
  openDocxFile,
  titleFromFileName,
  DocxOpenError,
  FlatOdfError,
  LegacyDocError,
} from "./openDocument";
import { FLAT_ODF_MESSAGE } from "./odf/reader";

const fixturesDir = join(dirname(fileURLToPath(import.meta.url)), "docx", "fixtures");
const odfFixturesDir = join(dirname(fileURLToPath(import.meta.url)), "odf", "fixtures");

function loadFixtureFile(name: string): File {
  const bytes = readFileSync(join(fixturesDir, name));
  return new File([bytes], name);
}

describe("titleFromFileName", () => {
  it("strips the .docx extension", () => {
    expect(titleFromFileName("Quarterly Report.docx")).toBe("Quarterly Report");
  });

  it("only strips the trailing extension, not dots inside the name", () => {
    expect(titleFromFileName("Meeting.Notes.v2.docx")).toBe("Meeting.Notes.v2");
  });

  it("is case-insensitive about the extension", () => {
    expect(titleFromFileName("Report.DOCX")).toBe("Report");
  });
});

describe("openDocxFile", () => {
  it("opens a real .docx file into a DocumentModel with format 'docx' (AC-2, AC-30)", async () => {
    const file = loadFixtureFile("word-authored.docx");
    const doc = await openDocxFile(file);
    expect(doc.format).toBe("docx");
    expect(doc.sections[0].blocks).toHaveLength(4);
  });

  it("titles the document from the file name without extension (AC-29)", async () => {
    const bytes = readFileSync(join(fixturesDir, "word-authored.docx"));
    const file = new File([bytes], "My Report.docx");
    const doc = await openDocxFile(file);
    expect(doc.title).toBe("My Report");
  });

  it("gives each opened document a distinct id", async () => {
    const file = loadFixtureFile("word-authored.docx");
    const first = await openDocxFile(file);
    const second = await openDocxFile(file);
    expect(first.id).not.toBe(second.id);
  });

  it("rejects a non-ZIP file with DocxOpenError (AC-31)", async () => {
    const file = new File([new TextEncoder().encode("not a docx")], "broken.docx");
    await expect(openDocxFile(file)).rejects.toThrow(DocxOpenError);
  });

  it("rejects a legacy .doc file with LegacyDocError (AC-34)", async () => {
    const magic = new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1, 0, 0]);
    const file = new File([magic], "old.docx");
    await expect(openDocxFile(file)).rejects.toThrow(LegacyDocError);
  });
});

describe("titleFromFileName — .odt", () => {
  it("strips the .odt extension, case-insensitively", () => {
    expect(titleFromFileName("Meeting Minutes.odt")).toBe("Meeting Minutes");
    expect(titleFromFileName("Report.ODT")).toBe("Report");
  });
});

describe("openDocumentFile", () => {
  const FLAT_ODF = `<?xml version="1.0" encoding="UTF-8"?>
<office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" office:mimetype="application/vnd.oasis.opendocument.text" office:version="1.3"/>`;

  function loadOdfFixtureFile(name: string, as: string = name): File {
    const bytes = readFileSync(join(odfFixturesDir, name));
    return new File([bytes], as);
  }

  it("opens a real .odt file into a DocumentModel with format 'odt' (AC-1, AC-30)", async () => {
    const doc = await openDocumentFile(loadOdfFixtureFile("libreoffice-writer.odt"));
    expect(doc.format).toBe("odt");
    expect(doc.title).toBe("libreoffice-writer");
    expect(doc.sections[0].blocks.length).toBeGreaterThan(0);
  });

  it("detects ODF content behind a .docx name — mimetype entry wins over extension (AC-2)", async () => {
    const doc = await openDocumentFile(loadOdfFixtureFile("libreoffice-writer.odt", "renamed.docx"));
    expect(doc.format).toBe("odt");
  });

  it("detects OOXML content behind a .odt name — mimetype entry wins over extension (AC-2)", async () => {
    const bytes = readFileSync(join(fixturesDir, "word-authored.docx"));
    const doc = await openDocumentFile(new File([bytes], "renamed.odt"));
    expect(doc.format).toBe("docx");
  });

  it("rejects a flat XML ODF (.fodt) file with the specific message (AC-29)", async () => {
    const file = new File([new TextEncoder().encode(FLAT_ODF)], "flat.fodt");
    await expect(openDocumentFile(file)).rejects.toThrow(FlatOdfError);
    await expect(openDocumentFile(file)).rejects.toThrow(FLAT_ODF_MESSAGE);
  });

  it("still rejects non-ZIP and legacy .doc files via the .docx path", async () => {
    const notZip = new File([new TextEncoder().encode("not a document")], "broken.odt");
    await expect(openDocumentFile(notZip)).rejects.toThrow(DocxOpenError);

    const magic = new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1, 0, 0]);
    await expect(openDocumentFile(new File([magic], "old.doc"))).rejects.toThrow(LegacyDocError);
  });
});
