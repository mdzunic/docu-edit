import { describe, it, expect } from "vitest";
import { readFileSync } from "node:fs";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import { parseDocx } from "./reader";
import { parseOdt } from "../odf/reader";
import { writeDocx } from "./writer";
import { validateDocx } from "./validate";
import { unzip } from "../core/zip";
import { zip } from "../core/zipWriter";
import { createDocument, type DocumentModel } from "../documentModel";

const here = dirname(fileURLToPath(import.meta.url));

function loadFixture(...segments: string[]): Uint8Array {
  return new Uint8Array(readFileSync(join(here, ...segments)));
}

async function writtenFixture(name: string): Promise<Uint8Array> {
  const parsed = await parseDocx(loadFixture("fixtures", name));
  const model: DocumentModel = {
    id: "test-doc",
    title: "Fixture",
    sections: parsed.sections,
    defaultStyles: parsed.defaultStyles,
    format: "docx",
  };
  return (await writeDocx(model)).bytes;
}

const decoder = new TextDecoder();
const encoder = new TextEncoder();

/** Re-packages writer output with one part replaced (or removed, when
 * `replacement` is null), preserving part order — for proving the validator
 * actually catches broken packages rather than rubber-stamping. */
async function tamper(
  bytes: Uint8Array,
  partName: string,
  replacement: string | null | ((current: string) => string),
): Promise<Uint8Array> {
  const parts = await unzip(bytes);
  const entries: { name: string; data: Uint8Array }[] = [];
  for (const [name, data] of parts) {
    if (name !== partName) {
      entries.push({ name, data });
    } else if (replacement !== null) {
      const text = typeof replacement === "function" ? replacement(decoder.decode(data)) : replacement;
      entries.push({ name, data: encoder.encode(text) });
    }
  }
  return zip(entries);
}

describe("validateDocx — accepts everything the writer produces (AC-19, AC-22)", () => {
  it.each(["word-authored.docx", "libreoffice-authored.docx", "google-docs-exported.docx"])(
    "round-tripped %s validates with no violations",
    async (fixtureName) => {
      expect(await validateDocx(await writtenFixture(fixtureName))).toEqual([]);
    },
  );

  it("a blank document validates with no violations", async () => {
    expect(await validateDocx((await writeDocx(createDocument("id-1"))).bytes)).toEqual([]);
  });

  it("a document converted from .odt validates with no violations (AC-23)", async () => {
    const parsed = await parseOdt(loadFixture("..", "odf", "fixtures", "libreoffice-writer.odt"));
    const model: DocumentModel = {
      id: "odt-doc",
      title: "Project Plan",
      sections: parsed.sections,
      defaultStyles: parsed.defaultStyles,
      format: "odt",
    };
    expect(await validateDocx((await writeDocx(model)).bytes)).toEqual([]);
  });
});

describe("validateDocx — rejects broken packages", () => {
  it("flags bytes that are not a ZIP archive", async () => {
    const violations = await validateDocx(new TextEncoder().encode("not a zip at all, sorry"));
    expect(violations).toHaveLength(1);
    expect(violations[0]).toContain("not a valid ZIP");
  });

  it("flags a missing required part", async () => {
    const broken = await tamper(await writtenFixture("word-authored.docx"), "word/styles.xml", null);
    const violations = await validateDocx(broken);
    expect(violations.some((v) => v.includes("missing required part word/styles.xml"))).toBe(true);
  });

  it("flags a part that is not well-formed XML", async () => {
    const broken = await tamper(
      await writtenFixture("word-authored.docx"),
      "word/document.xml",
      "<w:document><unclosed>",
    );
    const violations = await validateDocx(broken);
    expect(violations.some((v) => v.includes("not well-formed XML"))).toBe(true);
  });

  it("flags rPr children out of the schema's sequence order (AC-22)", async () => {
    const broken = await tamper(
      await writtenFixture("word-authored.docx"),
      "word/document.xml",
      (xml) => xml.replace("<w:rPr><w:rFonts", '<w:rPr><w:sz w:val="22"/><w:rFonts'),
    );
    const violations = await validateDocx(broken);
    expect(violations.some((v) => v.includes("out of schema sequence order"))).toBe(true);
  });

  it("flags an invalid w:jc value", async () => {
    const broken = await tamper(
      await writtenFixture("word-authored.docx"),
      "word/document.xml",
      (xml) => xml.replace('<w:jc w:val="center"/>', '<w:jc w:val="middle"/>'),
    );
    const violations = await validateDocx(broken);
    expect(violations.some((v) => v.includes("not an ST_Jc value"))).toBe(true);
  });

  it("flags an out-of-range w:outlineLvl", async () => {
    const broken = await tamper(
      await writtenFixture("word-authored.docx"),
      "word/document.xml",
      (xml) => xml.replace('<w:outlineLvl w:val="0"/>', '<w:outlineLvl w:val="12"/>'),
    );
    const violations = await validateDocx(broken);
    expect(violations.some((v) => v.includes("outline level"))).toBe(true);
  });

  it("flags a w:pgMar missing its required attributes", async () => {
    const broken = await tamper(
      await writtenFixture("word-authored.docx"),
      "word/document.xml",
      (xml) => xml.replace(/ w:header="\d+"/, ""),
    );
    const violations = await validateDocx(broken);
    expect(violations.some((v) => v.includes("missing required attribute w:header"))).toBe(true);
  });

  it("flags a body whose only child is w:sectPr (the empty-file repair case, AC-24)", async () => {
    const broken = await tamper(
      await writtenFixture("word-authored.docx"),
      "word/document.xml",
      (xml) => xml.replace(/<w:p>.*<\/w:p>/, "").replace("<w:p/>", ""),
    );
    const violations = await validateDocx(broken);
    expect(violations.some((v) => v.includes("at least one w:p"))).toBe(true);
  });

  it("flags a missing dcterms:modified in core.xml (AC-15)", async () => {
    const broken = await tamper(
      await writtenFixture("word-authored.docx"),
      "docProps/core.xml",
      (xml) => xml.replace(/<dcterms:modified.*<\/dcterms:modified>/, ""),
    );
    const violations = await validateDocx(broken);
    expect(violations.some((v) => v.includes("missing dcterms:modified"))).toBe(true);
  });

  it("flags an undeclared content type", async () => {
    // The .rels parts have no Override, so dropping the rels Default leaves
    // them untyped (an Override removal alone would fall back to the xml
    // Default and stay legal).
    const broken = await tamper(
      await writtenFixture("word-authored.docx"),
      "[Content_Types].xml",
      (xml) => xml.replace(/<Default Extension="rels"[^>]*\/>/, ""),
    );
    const violations = await validateDocx(broken);
    expect(violations.some((v) => v.includes("no content type declared for _rels/.rels"))).toBe(
      true,
    );
  });

  it("flags a main part typed as something other than WordprocessingML", async () => {
    const broken = await tamper(
      await writtenFixture("word-authored.docx"),
      "[Content_Types].xml",
      (xml) => xml.replace(/<Override PartName="\/word\/document\.xml"[^>]*\/>/, ""),
    );
    const violations = await validateDocx(broken);
    expect(violations.some((v) => v.includes("not the WordprocessingML main part type"))).toBe(
      true,
    );
  });

  it("flags a missing officeDocument relationship (AC-5)", async () => {
    const broken = await tamper(
      await writtenFixture("word-authored.docx"),
      "_rels/.rels",
      (xml) => xml.replace(/<Relationship Id="rId1"[^>]*\/>/, ""),
    );
    const violations = await validateDocx(broken);
    expect(violations.some((v) => v.includes("missing officeDocument relationship"))).toBe(true);
  });
});
