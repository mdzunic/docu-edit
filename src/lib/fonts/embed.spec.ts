import { readFileSync } from "node:fs";
import { resolve } from "node:path";
import { describe, expect, it } from "vitest";
import { unzip } from "../core/zip";
import type { DocumentModel } from "../documentModel";
import { createDefaultStyles } from "../documentModel";
import { parseOdt } from "../odf/reader";
import { writeOdt } from "../odf/writer";
import { SVG_NAMESPACE, XLINK_NAMESPACE } from "../odf/xml";
import { extractDocxEmbeddedFonts, extractOdtEmbeddedFonts } from "./embeddedFonts";
import { embedFontsIntoDocx } from "./embedDocx";
import { embedFontsIntoOdt } from "./embedOdt";
import { collectFontUsage, formatEmbedSize, planFontEmbedding } from "./embedPlan";
import type { EmbedFace } from "./embedPlan";

function bundled(file: string): Uint8Array {
  return new Uint8Array(readFileSync(resolve(__dirname, "../../../public/fonts", file)));
}

function fixture(rel: string): Uint8Array {
  return new Uint8Array(readFileSync(resolve(__dirname, rel)));
}

function modelUsing(family: string, text: string, bold = false): DocumentModel {
  return {
    id: "d",
    title: "t",
    defaultStyles: { ...createDefaultStyles(), fontFamily: family },
    sections: [
      {
        blocks: [
          { type: "paragraph", formatting: {}, runs: [{ text, formatting: { fontFamily: family, bold } }] },
        ],
      },
    ],
  } as unknown as DocumentModel;
}

describe("collectFontUsage", () => {
  it("groups by family and style, unioning code points", () => {
    const usage = collectFontUsage(modelUsing("Calibri", "AB", true));
    // The default (Calibri regular, empty text) and the bold run.
    const bold = usage.find((u) => u.bold);
    expect(bold?.family).toBe("Calibri");
    expect(bold?.codepoints.has("A".codePointAt(0)!)).toBe(true);
  });
});

describe("planFontEmbedding", () => {
  it("subsets embeddable faces and skips + names restricted ones (AC-15/17/18)", () => {
    const model = modelUsing("Calibri", "Hello");
    const carlito = bundled("Carlito-Regular.ttf");
    const plan = planFontEmbedding(model, (family) => {
      if (family === "Restricted") {
        return { bytes: carlito, embeddable: false };
      }
      return { bytes: carlito, embeddable: true };
    });
    expect(plan.faces.length).toBeGreaterThan(0);
    expect(plan.faces[0].bytes.length).toBeLessThan(carlito.length); // subset
    expect(plan.totalBytes).toBeGreaterThan(0);
    expect(plan.skipped).toEqual([]);

    const withRestricted = planFontEmbedding(modelUsing("Restricted", "Hi"), (family) => ({
      bytes: carlito,
      embeddable: family !== "Restricted",
    }));
    expect(withRestricted.skipped).toContain("Restricted");
  });
});

describe("formatEmbedSize", () => {
  it("renders a human size", () => {
    expect(formatEmbedSize(512)).toBe("512 B");
    expect(formatEmbedSize(2048)).toBe("2.0 KB");
    expect(formatEmbedSize(5 * 1024 * 1024)).toBe("5.0 MB");
  });
});

const faces: EmbedFace[] = [
  { family: "MyFont", bold: false, italic: false, bytes: bundled("Carlito-Regular.ttf") },
  { family: "MyFont", bold: true, italic: false, bytes: bundled("Carlito-Bold.ttf") },
];

describe("embedFontsIntoDocx", () => {
  it("embeds faces that round-trip back through the extractor (AC-15)", async () => {
    const embedded = await embedFontsIntoDocx(fixture("../docx/fixtures/word-authored.docx"), faces);
    const parts = await unzip(embedded);
    // Declarations wired up.
    expect(new TextDecoder().decode(parts.get("[Content_Types].xml")!)).toContain('Extension="odttf"');
    expect(new TextDecoder().decode(parts.get("word/_rels/document.xml.rels")!)).toContain("/fontTable");
    // The extractor reads them straight back, de-obfuscated.
    const extracted = extractDocxEmbeddedFonts(parts);
    expect(extracted).toHaveLength(2);
    expect(extracted.map((f) => f.family)).toEqual(["MyFont", "MyFont"]);
    const regular = extracted.find((f) => !f.bold)!;
    expect(Array.from(regular.bytes)).toEqual(Array.from(faces[0].bytes));
  });

  it("leaves the bytes unchanged when there is nothing to embed", async () => {
    const original = fixture("../docx/fixtures/word-authored.docx");
    expect(await embedFontsIntoDocx(original, [])).toBe(original);
  });
});

describe("embedFontsIntoOdt", () => {
  it("embeds faces that round-trip back through the extractor (AC-15)", async () => {
    const embedded = await embedFontsIntoOdt(fixture("../odf/fixtures/libreoffice-writer.odt"), faces);
    const parts = await unzip(embedded);
    const content = new TextDecoder().decode(parts.get("content.xml")!);
    expect(content).toContain("svg:font-family=\"MyFont\"");
    const manifest = new TextDecoder().decode(parts.get("META-INF/manifest.xml")!);
    expect(manifest).toContain("Fonts/");
    // Serialize the font-face-decls element with its namespaces in scope, the
    // way the reader captures sourceXml.odtFontFaceDecls, then read it back.
    const doc = new DOMParser().parseFromString(content, "application/xml");
    const decls = doc.getElementsByTagNameNS(
      "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
      "font-face-decls",
    )[0];
    const extracted = extractOdtEmbeddedFonts(new XMLSerializer().serializeToString(decls), parts);
    expect(extracted.length).toBeGreaterThanOrEqual(2);
    const boldFace = extracted.find((f) => f.bold);
    expect(boldFace?.family).toBe("MyFont");
  });

  it("stays namespace-well-formed when embedding into writeOdt output (AC-15)", async () => {
    // The real save path: a document with no preserved font-face-decls,
    // re-serialised by writeOdt — whose content.xml root declares only
    // office/style/text/fo, NOT svg/xlink — then embedded. This is where an
    // undeclared svg:/xlink: prefix on the injected <style:font-face> corrupts
    // the package (a raw LibreOffice fixture can't catch it: its own root
    // already declares svg/xlink, so the injected prefixes inherit).
    const opened = await parseOdt(fixture("../odf/fixtures/libreoffice-writer.odt"));
    const model: DocumentModel = {
      id: "d",
      title: "t",
      sections: opened.sections,
      defaultStyles: opened.defaultStyles,
      format: "odt",
      opaque: opened.opaque,
      sourceXml: { ...opened.sourceXml, odtFontFaceDecls: undefined },
    };
    const { bytes } = await writeOdt(model);
    const embedded = await embedFontsIntoOdt(bytes, faces);

    // The app must be able to reopen its own output: the reader parses
    // content.xml with a strict, namespace-aware DOMParser and throws
    // OdtOpenError on ill-formed markup (QA's exact symptom), and the embedded
    // faces must round-trip back out.
    const reopened = await parseOdt(embedded);
    const reExtracted = extractOdtEmbeddedFonts(
      reopened.sourceXml.odtFontFaceDecls ?? "",
      await unzip(embedded),
    );
    expect(reExtracted.map((f) => f.family)).toContain("MyFont");

    // The injected face must bind svg:/xlink: in its own scope regardless of the
    // root — assert on the parsed element so this holds in any DOMParser, not
    // only the strict browser one that surfaced the corruption for QA.
    const content = new TextDecoder().decode((await unzip(embedded)).get("content.xml")!);
    const face = new DOMParser()
      .parseFromString(content, "application/xml")
      .getElementsByTagNameNS("urn:oasis:names:tc:opendocument:xmlns:style:1.0", "font-face")[0];
    expect(face.lookupNamespaceURI("svg")).toBe(SVG_NAMESPACE);
    expect(face.lookupNamespaceURI("xlink")).toBe(XLINK_NAMESPACE);
  });
});
