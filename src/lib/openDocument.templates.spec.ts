// SPEC-046 AC-9..11 — opening .dotx/.ott files as templates: format
// detection widened for .ott (an ODF package whose mimetype declares the
// text-template variant), and openTemplateFile drops the retained package
// so a later save never resurfaces the source's "this is a template"
// package declarations on what is now an ordinary document.
import { describe, it, expect } from "vitest";
import { readFileSync } from "node:fs";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import { detectOdf, OTT_MIMETYPE } from "./odf/reader";
import { hasTemplateFileExtension, openTemplateFile, titleFromFileName } from "./openDocument";
import { unzip } from "./docx/zip";
import { zip } from "./docx/zipWriter";

const docxFixturesDir = join(dirname(fileURLToPath(import.meta.url)), "docx", "fixtures");
const odfFixturesDir = join(dirname(fileURLToPath(import.meta.url)), "odf", "fixtures");

function loadFixtureBytes(dir: string, name: string): Uint8Array {
  return new Uint8Array(readFileSync(join(dir, name)));
}

/** Repackages a real .odt fixture with the .ott mimetype — this environment
 * has no genuine .ott export to load (see odf/reader.spec.ts's fixture doc
 * comment for the same constraint on the .odt fixtures themselves), but an
 * .ott is byte-for-byte the same package shape as .odt bar this one entry. */
async function asOttBytes(odtBytes: Uint8Array): Promise<Uint8Array> {
  const parts = await unzip(odtBytes);
  const entries = [...parts.entries()].map(([name, data]) =>
    name === "mimetype" ? { name, data: new TextEncoder().encode(OTT_MIMETYPE), store: true } : { name, data },
  );
  return zip(entries);
}

describe("detectOdf recognises .ott alongside .odt (AC-10)", () => {
  it("still detects a plain .odt as a package", async () => {
    const odt = loadFixtureBytes(odfFixturesDir, "libreoffice-writer.odt");
    expect(await detectOdf(odt)).toBe("package");
  });

  it("detects an .ott (text-template mimetype) as a package too", async () => {
    const odt = loadFixtureBytes(odfFixturesDir, "libreoffice-writer.odt");
    const ott = await asOttBytes(odt);
    expect(await detectOdf(ott)).toBe("package");
  });
});

describe("hasTemplateFileExtension", () => {
  it("accepts .dotx and .ott, case-insensitively", () => {
    expect(hasTemplateFileExtension("Invoice.dotx")).toBe(true);
    expect(hasTemplateFileExtension("Invoice.DOTX")).toBe(true);
    expect(hasTemplateFileExtension("Invoice.ott")).toBe(true);
  });

  it("rejects plain .docx/.odt (those are documents, not templates)", () => {
    expect(hasTemplateFileExtension("Invoice.docx")).toBe(false);
    expect(hasTemplateFileExtension("Invoice.odt")).toBe(false);
  });
});

describe("titleFromFileName strips template extensions too", () => {
  it("strips .dotx and .ott", () => {
    expect(titleFromFileName("Invoice.dotx")).toBe("Invoice");
    expect(titleFromFileName("Invoice.ott")).toBe("Invoice");
  });
});

describe("openTemplateFile (AC-9, AC-11)", () => {
  it("opens a .dotx (an OPC docx package) as a new document with a fresh id", async () => {
    const bytes = loadFixtureBytes(docxFixturesDir, "word-authored.docx");
    const file = new File([bytes], "Letter.dotx");
    const model = await openTemplateFile(file);
    expect(model.format).toBe("docx");
    expect(model.title).toBe("Letter");
    expect(model.sections[0].blocks.length).toBeGreaterThan(0);
    expect(model.id).toBeTruthy();
  });

  it("opens an .ott as a new document with a fresh id", async () => {
    const odt = loadFixtureBytes(odfFixturesDir, "libreoffice-writer.odt");
    const ott = await asOttBytes(odt);
    const file = new File([ott], "Memo.ott");
    const model = await openTemplateFile(file);
    expect(model.format).toBe("odt");
    expect(model.title).toBe("Memo");
    expect(model.sections[0].blocks.length).toBeGreaterThan(0);
  });

  it("two opens of the same template file produce independent documents with different ids", async () => {
    const bytes = loadFixtureBytes(docxFixturesDir, "word-authored.docx");
    const file = new File([bytes], "Letter.dotx");
    const first = await openTemplateFile(file);
    const second = await openTemplateFile(file);
    expect(first.id).not.toBe(second.id);
  });

  it("drops the retained package so a later save never re-declares the source as a template", async () => {
    const bytes = loadFixtureBytes(docxFixturesDir, "word-authored.docx");
    const file = new File([bytes], "Letter.dotx");
    const model = await openTemplateFile(file);
    expect(model.opaque).toBeUndefined();
    expect(model.sourceXml).toBeUndefined();
  });
});
