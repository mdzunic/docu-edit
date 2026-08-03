import { describe, it, expect } from "vitest";
import { readFileSync } from "node:fs";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import { openDocxFile, titleFromFileName, DocxOpenError, LegacyDocError } from "./openDocument";

const fixturesDir = join(dirname(fileURLToPath(import.meta.url)), "docx", "fixtures");

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
