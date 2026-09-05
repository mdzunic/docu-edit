import { describe, it, expect } from "vitest";
import { readFileSync, readdirSync } from "node:fs";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import { parseDocx } from "./reader";
import { parseOdt } from "../odf/reader";
import { readDocxBytes } from "../openDocument";
import { writeDocx } from "./writer";
import { validateDocx } from "./validate";
import { unzip } from "../core/zip";
import { zip } from "../core/zipWriter";
import { createDocument, type DocumentModel } from "../documentModel";

const here = dirname(fileURLToPath(import.meta.url));

function loadFixture(...segments: string[]): Uint8Array {
  return new Uint8Array(readFileSync(join(here, ...segments)));
}

/** A fixture opened the way the app opens it — the full model, retained
 * source package included — and saved back. Building a bare model here
 * instead (sections + defaultStyles only) skipped SPEC-008's preservation
 * path entirely, which is where every package-level violation this
 * validator now catches actually comes from. */
async function writtenFixture(name: string): Promise<Uint8Array> {
  return (await writeDocx(await readDocxBytes(loadFixture("fixtures", name), "Fixture"))).bytes;
}

/** Every `.docx` fixture in the directory, so a new one is covered the day
 * it lands rather than whenever someone remembers to extend a list. */
const DOCX_FIXTURES = readdirSync(join(here, "fixtures"))
  .filter((name) => name.endsWith(".docx"))
  .sort();

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
  it.each(DOCX_FIXTURES)("round-tripped %s validates with no violations", async (fixtureName) => {
    expect(await validateDocx(await writtenFixture(fixtureName))).toEqual([]);
  });

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

// The OPC packaging invariants an append-only relationship/content-type
// patch can break. None of them stop OUR reader re-opening the file, so a
// model-equality round-trip test cannot see them — only a package-level
// check can, which is the whole reason these live here.
describe("validateDocx — OPC packaging invariants", () => {
  it("flags a duplicate Override for one part (ECMA-376 Part 2 §10.1.2.3)", async () => {
    const broken = await tamper(
      await writtenFixture("word-authored.docx"),
      "[Content_Types].xml",
      (xml) =>
        xml.replace(
          "</Types>",
          '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>',
        ),
    );
    expect(await validateDocx(broken)).toContain(
      "[Content_Types].xml: duplicate Override for part /word/styles.xml",
    );
  });

  it("flags a duplicate Default for one extension (§10.1.2.2)", async () => {
    const broken = await tamper(
      await writtenFixture("word-authored.docx"),
      "[Content_Types].xml",
      (xml) => xml.replace("</Types>", '<Default Extension="xml" ContentType="application/xml"/></Types>'),
    );
    expect(await validateDocx(broken)).toContain('[Content_Types].xml: duplicate Default for extension "xml"');
  });

  it("flags an Override naming a part the package does not contain", async () => {
    const broken = await tamper(
      await writtenFixture("word-authored.docx"),
      "[Content_Types].xml",
      (xml) =>
        xml.replace(
          "</Types>",
          '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>',
        ),
    );
    expect(await validateDocx(broken)).toContain(
      "[Content_Types].xml: Override names missing part /word/header1.xml",
    );
  });

  it("flags a duplicate relationship Id within one .rels part", async () => {
    const broken = await tamper(
      await writtenFixture("word-authored.docx"),
      "word/_rels/document.xml.rels",
      (xml) =>
        xml.replace(
          "</Relationships>",
          '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>',
        ),
    );
    expect(await validateDocx(broken)).toContain(
      'word/_rels/document.xml.rels: duplicate relationship Id "rId1"',
    );
  });

  it("flags a relationship targeting a part the package does not contain", async () => {
    const broken = await tamper(
      await writtenFixture("word-authored.docx"),
      "word/_rels/document.xml.rels",
      (xml) =>
        xml.replace(
          "</Relationships>",
          '<Relationship Id="rIdGone" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/></Relationships>',
        ),
    );
    expect(await validateDocx(broken)).toContain(
      'word/_rels/document.xml.rels: relationship "rIdGone" targets missing part word/media/image1.png',
    );
  });

  it("accepts an external relationship target, which names no part at all", async () => {
    const withLink = await tamper(
      await writtenFixture("word-authored.docx"),
      "word/_rels/document.xml.rels",
      (xml) =>
        xml.replace(
          "</Relationships>",
          '<Relationship Id="rIdExt" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com" TargetMode="External"/></Relationships>',
        ),
    );
    expect(await validateDocx(withLink)).toEqual([]);
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
