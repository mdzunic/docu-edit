import { describe, it, expect } from "vitest";
import { readFileSync } from "node:fs";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import { detectOdf, parseOdt, type ParsedOdtDocument } from "./reader";
import { parseDocx, type ParsedDocxDocument } from "../docx/reader";
import { writeDocx } from "../docx/writer";
import {
  buildMetaXml,
  ODT_MIMETYPE,
  odtFileName,
  UNTITLED_ODT_TITLE,
  writeOdt,
} from "./writer";
import { unzip } from "../core/zip";
import { parseXml } from "../core/xml";
import { OFFICE_NAMESPACE, STYLE_NAMESPACE, TEXT_NAMESPACE } from "../core/odfXml";
import { asParagraph, createDocument, type Block, type DocumentModel, type Paragraph } from "../documentModel";
import { firstStructuralDifference } from "../equality";

const here = dirname(fileURLToPath(import.meta.url));
const odfFixturesDir = join(here, "fixtures");
const docxFixturesDir = join(here, "..", "docx", "fixtures");

function loadFixture(dir: string, name: string): Uint8Array {
  return new Uint8Array(readFileSync(join(dir, name)));
}

function modelFrom(parsed: ParsedOdtDocument, title = "Fixture"): DocumentModel {
  return {
    id: "test-doc",
    title,
    sections: parsed.sections,
    defaultStyles: parsed.defaultStyles,
    format: "odt",
  };
}

function modelFromDocx(parsed: ParsedDocxDocument, title = "Fixture"): DocumentModel {
  return {
    id: "test-doc",
    title,
    sections: parsed.sections,
    defaultStyles: parsed.defaultStyles,
    format: "docx",
  };
}

const decoder = new TextDecoder();

async function partText(bytes: Uint8Array, name: string): Promise<string> {
  const parts = await unzip(bytes);
  const raw = parts.get(name);
  if (!raw) {
    throw new Error(`part ${name} missing from written package`);
  }
  return decoder.decode(raw);
}

function paragraphTexts(model: { sections: { blocks: Block[] }[] }): string[] {
  return model.sections.flatMap((section) =>
    section.blocks.map((block) => asParagraph(block).runs.map((run) => run.text).join("")),
  );
}

function headingLevels(model: { sections: { blocks: Block[] }[] }): (number | undefined)[] {
  return model.sections.flatMap((section) => section.blocks.map((block) => asParagraph(block).headingLevel));
}

describe("writeOdt — package structure (AC-4..AC-8)", () => {
  it("writes an ODF package containing every required part", async () => {
    const bytes = (await writeOdt(createDocument("id-1"))).bytes;
    const parts = await unzip(bytes);
    for (const required of [
      "mimetype",
      "content.xml",
      "styles.xml",
      "meta.xml",
      "settings.xml",
      "META-INF/manifest.xml",
    ]) {
      expect(parts.has(required), `expected ${required} in the package`).toBe(true);
    }
  });

  it("every written XML part is well-formed", async () => {
    const bytes = (await writeOdt(createDocument("id-1"))).bytes;
    const parts = await unzip(bytes);
    for (const [name, raw] of parts) {
      if (name === "mimetype") {
        continue;
      }
      expect(() => parseXml(decoder.decode(raw)), `${name} should parse`).not.toThrow();
    }
  });

  it("the manifest lists the package root with the ODT media type and every XML part", async () => {
    const manifest = parseXml(
      await partText((await writeOdt(createDocument("id-1"))).bytes, "META-INF/manifest.xml"),
    );
    const ns = "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0";
    const entries = Array.from(manifest.getElementsByTagNameNS(ns, "file-entry"));
    const byPath = new Map(
      entries.map((entry) => [
        entry.getAttributeNS(ns, "full-path"),
        entry.getAttributeNS(ns, "media-type"),
      ]),
    );
    expect(byPath.get("/")).toBe(ODT_MIMETYPE);
    for (const path of ["content.xml", "styles.xml", "meta.xml", "settings.xml"]) {
      expect(byPath.get(path), `manifest should list ${path}`).toBe("text/xml");
    }
  });

  it("the reader's own detection accepts the written package", async () => {
    expect(await detectOdf((await writeOdt(createDocument("id-1"))).bytes)).toBe("package");
  });
});

describe("writeOdt — mimetype entry (AC-3, AC-9, AC-10)", () => {
  it("stores the mimetype uncompressed as the first ZIP entry, containing exactly the ODT media type", async () => {
    const bytes = (await writeOdt(createDocument("id-1"))).bytes;
    const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    // The first local file header sits at offset 0 (AC-3).
    expect(view.getUint32(0, true)).toBe(0x04034b50);
    expect(view.getUint16(8, true), "compression method should be STORE").toBe(0);
    const compressedSize = view.getUint32(18, true);
    const uncompressedSize = view.getUint32(22, true);
    expect(compressedSize).toBe(uncompressedSize);
    const nameLength = view.getUint16(26, true);
    const extraLength = view.getUint16(28, true);
    expect(decoder.decode(bytes.subarray(30, 30 + nameLength))).toBe("mimetype");
    const dataStart = 30 + nameLength + extraLength;
    expect(decoder.decode(bytes.subarray(dataStart, dataStart + compressedSize))).toBe(
      "application/vnd.oasis.opendocument.text",
    );
  });
});

describe("writeOdt — meta.xml (AC-19, AC-20)", () => {
  const modified = new Date("2026-08-03T12:34:56Z");

  it("carries the document title", async () => {
    const model = { ...createDocument("id-1"), title: "Quarterly Report" };
    const meta = parseXml(await partText((await writeOdt(model, { modified })).bytes, "meta.xml"));
    const title = meta.getElementsByTagNameNS("http://purl.org/dc/elements/1.1/", "title")[0];
    expect(title.textContent).toBe("Quarterly Report");
  });

  it("carries a dc:date modification timestamp", async () => {
    const meta = parseXml(
      await partText((await writeOdt(createDocument("id-1"), { modified })).bytes, "meta.xml"),
    );
    const date = meta.getElementsByTagNameNS("http://purl.org/dc/elements/1.1/", "date")[0];
    expect(date.textContent).toBe("2026-08-03T12:34:56Z");
  });

  it("escapes XML metacharacters in the title", () => {
    const xml = buildMetaXml({ title: 'Q3 <Report> & "Friends"' }, modified);
    const doc = parseXml(xml);
    expect(
      doc.getElementsByTagNameNS("http://purl.org/dc/elements/1.1/", "title")[0].textContent,
    ).toBe('Q3 <Report> & "Friends"');
  });

  it("falls back to 'Document 1' when the model title is blank (spec decision)", async () => {
    const model = { ...createDocument("id-1"), title: "   " };
    const meta = await partText((await writeOdt(model, { modified })).bytes, "meta.xml");
    expect(meta).toContain(`<dc:title>${UNTITLED_ODT_TITLE}</dc:title>`);
  });
});

describe.each(["libreoffice-writer.odt", "openoffice-era.odt"])(
  "writeOdt — %s round trip (AC-11..AC-15, AC-23)",
  (fixtureName) => {
    it("open -> save -> re-open reproduces the model exactly", async () => {
      const first = await parseOdt(loadFixture(odfFixturesDir, fixtureName));
      const written = (await writeOdt(modelFrom(first))).bytes;
      const second = await parseOdt(written);
      // Text, run formatting, paragraph formatting, style references and
      // heading levels all live inside sections; page layout and default
      // font live in defaultStyles. Together these are the whole parsed
      // model.
      expect(second.sections).toEqual(first.sections);
      expect(second.defaultStyles).toEqual(first.defaultStyles);
    });

    it("a second save of the re-opened model is byte-identical (stable output)", async () => {
      const modified = new Date("2026-08-03T12:34:56Z");
      const first = await parseOdt(loadFixture(odfFixturesDir, fixtureName));
      const written = (await writeOdt(modelFrom(first), { modified })).bytes;
      const reopened = await parseOdt(written);
      const rewritten = (await writeOdt(modelFrom(reopened), { modified })).bytes;
      expect(rewritten).toEqual(written);
    });
  },
);

describe("writeOdt — automatic styles (AC-16, AC-17, AC-18)", () => {
  async function contentDom(model: DocumentModel): Promise<Document> {
    return parseXml(await partText((await writeOdt(model)).bytes, "content.xml"));
  }

  function formattedModel(): DocumentModel {
    const model = createDocument("id-1");
    model.sections[0].blocks = [
      {
        type: "paragraph",
        runs: [{ text: "styled", formatting: { bold: true, color: "#ff0000" } }],
        formatting: { alignment: "center", indentStartCm: 1.27 },
        styleId: "Fancy",
      },
    ];
    return model;
  }

  it("emits run formatting as an automatic text style referenced by name (AC-16)", async () => {
    const dom = await contentDom(formattedModel());
    const span = dom.getElementsByTagNameNS(TEXT_NAMESPACE, "span")[0];
    const styleName = span.getAttributeNS(TEXT_NAMESPACE, "style-name");
    expect(styleName).toBeTruthy();
    const styles = Array.from(dom.getElementsByTagNameNS(STYLE_NAMESPACE, "style"));
    const definition = styles.find(
      (style) =>
        style.getAttributeNS(STYLE_NAMESPACE, "name") === styleName &&
        style.getAttributeNS(STYLE_NAMESPACE, "family") === "text",
    );
    expect(definition, "referenced text style should be defined").toBeTruthy();
    expect(definition!.parentElement!.localName).toBe("automatic-styles");
    const props = definition!.getElementsByTagNameNS(STYLE_NAMESPACE, "text-properties")[0];
    expect(props.getAttributeNS("urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0", "font-weight")).toBe("bold");
  });

  it("emits paragraph formatting as an automatic paragraph style named after the styleId (AC-17)", async () => {
    const dom = await contentDom(formattedModel());
    const paragraph = dom.getElementsByTagNameNS(TEXT_NAMESPACE, "p")[0];
    expect(paragraph.getAttributeNS(TEXT_NAMESPACE, "style-name")).toBe("Fancy");
    const styles = Array.from(dom.getElementsByTagNameNS(STYLE_NAMESPACE, "style"));
    const definition = styles.find(
      (style) =>
        style.getAttributeNS(STYLE_NAMESPACE, "name") === "Fancy" &&
        style.getAttributeNS(STYLE_NAMESPACE, "family") === "paragraph",
    );
    expect(definition, "referenced paragraph style should be defined").toBeTruthy();
    const props = definition!.getElementsByTagNameNS(STYLE_NAMESPACE, "paragraph-properties")[0];
    expect(props.getAttributeNS("urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0", "text-align")).toBe("center");
  });

  it("emits no inline formatting attributes on body elements (AC-18)", async () => {
    const first = await parseOdt(loadFixture(odfFixturesDir, "libreoffice-writer.odt"));
    const dom = await contentDom(modelFrom(first));
    const body = dom.getElementsByTagNameNS(OFFICE_NAMESPACE, "body")[0];
    const allowed = new Set(["style-name", "outline-level", "c", "name"]);
    // Hyperlinks/bookmarks (SPEC-025) put link *targets* inline — text:a's
    // xlink:href/type, text:bookmark's text:name, and the xmlns:xlink
    // declaration on a link-carrying paragraph. Targets are content, not
    // formatting, so they are the one addition to this whitelist; every
    // formatting attribute is still required to live in automatic styles.
    const XLINK_NS = "http://www.w3.org/1999/xlink";
    const XMLNS_NS = "http://www.w3.org/2000/xmlns/";
    for (const el of Array.from(body.getElementsByTagName("*"))) {
      for (const attr of Array.from(el.attributes)) {
        expect(
          (attr.namespaceURI === TEXT_NAMESPACE && allowed.has(attr.localName)) ||
            (attr.namespaceURI === XLINK_NS && (attr.localName === "href" || attr.localName === "type")) ||
            (attr.namespaceURI === XMLNS_NS && attr.localName === "xlink"),
          `unexpected inline attribute ${attr.name} on ${el.tagName}`,
        ).toBe(true);
      }
    }
  });
});

describe("writeOdt — content details", () => {
  it("round-trips tabs and forced line breaks as text:tab/text:line-break", async () => {
    const model = createDocument("id-1");
    model.sections[0].blocks = [
      { type: "paragraph", runs: [{ text: "col1\tcol2\nnext line", formatting: {} }] },
    ];
    const contentXml = await partText((await writeOdt(model)).bytes, "content.xml");
    expect(contentXml).toContain("<text:tab/>");
    expect(contentXml).toContain("<text:line-break/>");
    const reopened = await parseOdt((await writeOdt(model)).bytes);
    expect(paragraphTexts(reopened)).toEqual(["col1\tcol2\nnext line"]);
  });

  it("round-trips XML metacharacters and non-ASCII text", async () => {
    const text = `5 < 6 & 7 > 2, "quoted", 'single' — héllo π`;
    const model = createDocument("id-1");
    model.sections[0].blocks = [{ type: "paragraph", runs: [{ text, formatting: {} }] }];
    const reopened = await parseOdt((await writeOdt(model)).bytes);
    expect(paragraphTexts(reopened)).toEqual([text]);
  });

  it("protects leading, trailing and consecutive spaces with text:s", async () => {
    const model = createDocument("id-1");
    model.sections[0].blocks = [
      { type: "paragraph", runs: [{ text: "  spaced  out  ", formatting: {} }] },
    ];
    const contentXml = await partText((await writeOdt(model)).bytes, "content.xml");
    expect(contentXml).toContain('<text:s text:c="2"/>');
    const reopened = await parseOdt((await writeOdt(model)).bytes);
    expect(paragraphTexts(reopened)).toEqual(["  spaced  out  "]);
  });

  it("keeps consecutive no-break spaces as literal characters, never text:s (AC-56)", async () => {
    const model = createDocument("id-1");
    model.sections[0].blocks = [
      { type: "paragraph", runs: [{ text: "a  b", formatting: {} }] },
    ];
    const contentXml = await partText((await writeOdt(model)).bytes, "content.xml");
    // Two literal U+00A0 characters — ODF never collapses them, so no text:s.
    expect(contentXml).toContain("  ");
    expect(contentXml).not.toContain("<text:s/>");
    expect(contentXml).not.toContain("<text:s ");
  });

  it("protects a space that follows a tab across a run boundary", async () => {
    const model = createDocument("id-1");
    model.sections[0].blocks = [
      {
        type: "paragraph",
        runs: [
          { text: "tab\t", formatting: {} },
          { text: " next", formatting: { bold: true } },
        ],
      },
    ];
    const reopened = await parseOdt((await writeOdt(model)).bytes);
    expect(paragraphTexts(reopened)).toEqual(["tab\t next"]);
  });

  it("round-trips explicitly-disabled run formatting as normal/none (AC-12)", async () => {
    const model = createDocument("id-1");
    model.sections[0].blocks = [
      {
        type: "paragraph",
        runs: [{ text: "plain", formatting: { bold: false, underline: false } }],
      },
    ];
    const reopened = await parseOdt((await writeOdt(model)).bytes);
    const run = asParagraph(reopened.sections[0].blocks[0]).runs[0];
    expect(run.formatting.bold).toBe(false);
    expect(run.formatting.underline).toBe(false);
  });

  it("round-trips a heading level via text:h with text:outline-level", async () => {
    const model = createDocument("id-1");
    model.sections[0].blocks = [
      { type: "paragraph", runs: [{ text: "Deep", formatting: {} }], headingLevel: 3 },
    ];
    const contentXml = await partText((await writeOdt(model)).bytes, "content.xml");
    expect(contentXml).toContain('<text:h text:outline-level="3">');
    const reopened = await parseOdt((await writeOdt(model)).bytes);
    expect(asParagraph(reopened.sections[0].blocks[0]).headingLevel).toBe(3);
  });

  it("marks a landscape page with style:print-orientation and round-trips the page size (AC-15)", async () => {
    const model = createDocument("id-1");
    model.defaultStyles.pageSize = { widthCm: 29.7, heightCm: 21 };
    const stylesXml = await partText((await writeOdt(model)).bytes, "styles.xml");
    expect(stylesXml).toContain('style:print-orientation="landscape"');
    const reopened = await parseOdt((await writeOdt(model)).bytes);
    expect(reopened.defaultStyles.pageSize).toEqual({ widthCm: 29.7, heightCm: 21 });
  });
});

// SPEC-051 regression: shapeGraphicStyleXml/shapeElementXml emit draw:/svg:
// prefixed attributes but never declared those namespaces anywhere in the
// document, so any .odt containing a shape had an unparseable content.xml
// (an "unbound prefix" error) that neither this app nor any ODF reader could
// reopen. Exercised at the full write+reopen level, not just the shape.ts
// unit level, because the bug lived in the *absence* of a namespace
// declaration on the emitted elements — invisible to tests that only assert
// on individual attribute values.
describe("writeOdt — Shape (SPEC-051)", () => {
  it("writes well-formed content.xml for a filled, outlined shape and round-trips it", async () => {
    const model = createDocument("id-1");
    model.sections[0].blocks = [
      {
        type: "shape",
        geometry: "star5",
        widthCm: 4,
        heightCm: 4,
        wrap: "square",
        fill: { kind: "solid", color: "#00cc44" },
        outline: { color: "#cc0000", widthPt: 3, dash: "dash" },
        altText: "A green star",
      },
    ];
    const contentXml = await partText((await writeOdt(model)).bytes, "content.xml");
    // parseXml throws on an unbound-prefix/malformed document — this alone
    // would have caught the bug.
    expect(() => parseXml(contentXml)).not.toThrow();
    const reopened = await parseOdt((await writeOdt(model)).bytes);
    const shape = reopened.sections[0].blocks[0];
    expect(shape.type).toBe("shape");
    if (shape.type !== "shape") {
      return;
    }
    expect(shape.geometry).toBe("star5");
    expect(shape.fill).toEqual({ kind: "solid", color: "#00cc44" });
    expect(shape.outline?.color).toBe("#cc0000");
    expect(shape.outline?.dash).toBe("dash");
    expect(shape.altText).toBe("A green star");
  });

  it("writes well-formed content.xml for a plain default text box (no fill/outline)", async () => {
    const model = createDocument("id-1");
    model.sections[0].blocks = [
      { type: "shape", geometry: "rect", widthCm: 3, heightCm: 2, wrap: "inline" },
    ];
    const contentXml = await partText((await writeOdt(model)).bytes, "content.xml");
    expect(() => parseXml(contentXml)).not.toThrow();
    const reopened = await parseOdt((await writeOdt(model)).bytes);
    expect(reopened.sections[0].blocks[0].type).toBe("shape");
  });

  it("cross-format save of an unsupported geometry reports a shapeGeometry drop and still writes a well-formed, reopenable file (AC-61)", async () => {
    const unsupportedGeometryXml = '<a:prstGeom xmlns:a="urn:example" prst="pentagon"/>';
    const model = createDocument("id-1");
    model.format = "docx";
    model.sections[0].blocks = [
      { type: "shape", geometry: "rect", widthCm: 3, heightCm: 2, wrap: "square", unsupportedGeometryXml },
    ];
    const { bytes, dropped } = await writeOdt(model);
    // AC-61: a cross-format save can't translate a foreign geometry into
    // the target's own vocabulary, so it reports the loss...
    expect(dropped.map((d) => d.kind)).toContain("shapeGeometry");
    const contentXml = await partText(bytes, "content.xml");
    // ...but still writes a well-formed file — the DropRecord is the user
    // warning; a corrupt, unreadable package would be a second, undeclared
    // loss on top of it (this exact bug: SHAPE_XMLNS, fixed above).
    expect(() => parseXml(contentXml)).not.toThrow();
    const reopened = await parseOdt(bytes);
    expect(reopened.sections[0].blocks[0].type).toBe("shape");
  });
});

describe("writeOdt — empty document", () => {
  it("a freshly created blank document produces a valid one-paragraph file", async () => {
    const reopened = await parseOdt((await writeOdt(createDocument("id-1"))).bytes);
    expect(reopened.sections[0].blocks).toHaveLength(1);
    expect(asParagraph(reopened.sections[0].blocks[0]).runs).toEqual([]);
  });

  it("a model with no blocks at all still writes one paragraph", async () => {
    const model = createDocument("id-1");
    model.sections[0].blocks = [];
    const contentXml = await partText((await writeOdt(model)).bytes, "content.xml");
    expect(contentXml).toContain("<text:p/>");
    const reopened = await parseOdt((await writeOdt(model)).bytes);
    expect(reopened.sections[0].blocks).toHaveLength(1);
  });
});

describe("writeOdt — from a document opened as .docx (AC-24, AC-25, AC-26)", () => {
  it("produces a converted ODF package, not a renamed OPC file", async () => {
    const parsed = await parseDocx(loadFixture(docxFixturesDir, "word-authored.docx"));
    const bytes = (await writeOdt(modelFromDocx(parsed))).bytes;
    const parts = await unzip(bytes);
    expect(parts.has("content.xml")).toBe(true);
    expect(parts.has("[Content_Types].xml")).toBe(false);
    expect(parts.has("word/document.xml")).toBe(false);
    const contentXml = decoder.decode(parts.get("content.xml"));
    expect(contentXml).toContain("urn:oasis");
    expect(contentXml).not.toContain("schemas.openxmlformats.org");
  });

  it("preserves all text (AC-24)", async () => {
    const parsed = await parseDocx(loadFixture(docxFixturesDir, "word-authored.docx"));
    const reopened = await parseOdt((await writeOdt(modelFromDocx(parsed))).bytes);
    expect(paragraphTexts(reopened)).toEqual(paragraphTexts(parsed));
  });

  it("preserves headings (AC-25)", async () => {
    const parsed = await parseDocx(loadFixture(docxFixturesDir, "google-docs-exported.docx"));
    const reopened = await parseOdt((await writeOdt(modelFromDocx(parsed))).bytes);
    expect(headingLevels(reopened)).toEqual(headingLevels(parsed));
  });

  it("preserves bold and italic character formatting (AC-26)", async () => {
    const parsed = await parseDocx(loadFixture(docxFixturesDir, "word-authored.docx"));
    const reopened = await parseOdt((await writeOdt(modelFromDocx(parsed))).bytes);
    const runs = asParagraph(reopened.sections[0].blocks[1]).runs;
    expect(runs.some((r) => r.text === "bold" && r.formatting.bold === true)).toBe(true);
    expect(runs.some((r) => r.text === "italic" && r.formatting.italic === true)).toBe(true);
  });

  it("preserves underline and strikethrough character formatting (AC-26)", async () => {
    const parsed = await parseDocx(loadFixture(docxFixturesDir, "libreoffice-authored.docx"));
    const reopened = await parseOdt((await writeOdt(modelFromDocx(parsed))).bytes);
    const runs = asParagraph(reopened.sections[0].blocks[1]).runs;
    expect(runs.some((r) => r.text === "underlined" && r.formatting.underline === true)).toBe(true);
    expect(
      runs.some((r) => r.text === "struck through" && r.formatting.strikethrough === true),
    ).toBe(true);
  });
});

describe("writeDocx — from a document opened as .odt (AC-27, AC-28, AC-29)", () => {
  async function convert(fixtureName: string) {
    const parsed = await parseOdt(loadFixture(odfFixturesDir, fixtureName));
    const reopened = await parseDocx((await writeDocx(modelFrom(parsed))).bytes);
    return { parsed, reopened };
  }

  it("preserves all text (AC-27)", async () => {
    const { parsed, reopened } = await convert("libreoffice-writer.odt");
    expect(paragraphTexts(reopened)).toEqual(paragraphTexts(parsed));
  });

  it("preserves headings (AC-28)", async () => {
    const { parsed, reopened } = await convert("openoffice-era.odt");
    expect(headingLevels(reopened)).toEqual(headingLevels(parsed));
  });

  it("preserves character formatting (AC-29)", async () => {
    const { reopened } = await convert("libreoffice-writer.odt");
    const runs = asParagraph(reopened.sections[0].blocks[1]).runs;
    expect(runs.some((r) => r.text === "bold" && r.formatting.bold === true)).toBe(true);
    expect(runs.some((r) => r.text === "italic" && r.formatting.italic === true)).toBe(true);
    expect(runs.some((r) => r.text === "underlined" && r.formatting.underline === true)).toBe(true);
    expect(runs.some((r) => r.text === "struck" && r.formatting.strikethrough === true)).toBe(true);
    const red = asParagraph(reopened.sections[0].blocks[5]).runs[0];
    expect(red.text).toBe("Red Sans 14pt");
    expect(red.formatting.color).toBe("#c9211e");
    expect(red.formatting.fontSizePt).toBe(14);
    expect(red.formatting.fontFamily).toBe("Liberation Sans");
  });
});

describe("odtFileName (AC-2)", () => {
  it("is <title>.odt", () => {
    expect(odtFileName("Quarterly Report")).toBe("Quarterly Report.odt");
  });

  it("falls back to 'Document 1' for a blank title (spec decision)", () => {
    expect(odtFileName("   ")).toBe("Document 1.odt");
  });

  it("exposes the .odt MIME type for the download blob", () => {
    expect(ODT_MIMETYPE).toBe("application/vnd.oasis.opendocument.text");
  });
});

describe("writeOdt — paragraph geometry (SPEC-014 AC-25, AC-26)", () => {
  function geometryModel(): DocumentModel {
    const model = createDocument("geometry");
    model.sections[0].blocks = [
      {
        type: "paragraph",
        runs: [{ text: "rtl paragraph", formatting: { fontFamily: "Calibri", fontSizePt: 11 } }],
        formatting: {
          direction: "rtl",
          indentStartCm: 2.54,
          indentEndCm: 1.27,
          indentFirstLineCm: 0.635,
        },
      },
      {
        type: "paragraph",
        runs: [{ text: "exact", formatting: { fontFamily: "Calibri", fontSizePt: 11 } }],
        formatting: { lineSpacingPt: 14, lineSpacingRule: "exact", spaceBeforePt: 0 },
      },
      {
        type: "paragraph",
        runs: [{ text: "at least", formatting: { fontFamily: "Calibri", fontSizePt: 11 } }],
        formatting: { lineSpacingPt: 16, spaceAfterPt: 6 },
      },
    ];
    return model;
  }

  it("writes RTL indents on the physical edges LibreOffice expects (AC-25)", async () => {
    const contentXml = await partText((await writeOdt(geometryModel())).bytes, "content.xml");
    // Logical start (2.54) lands on the physical RIGHT edge; end (1.27) on
    // the left. text-indent stays start-relative.
    expect(contentXml).toContain('fo:margin-left="1.27cm"');
    expect(contentXml).toContain('fo:margin-right="2.54cm"');
    expect(contentXml).toContain('fo:text-indent="0.635cm"');
    expect(contentXml).toContain('style:writing-mode="rl-tb"');
  });

  it("writes an exact rule as fo:line-height and the default as line-height-at-least", async () => {
    const contentXml = await partText((await writeOdt(geometryModel())).bytes, "content.xml");
    expect(contentXml).toContain('fo:line-height="14pt"');
    expect(contentXml).toContain('style:line-height-at-least="16pt"');
  });

  it("save -> re-open reproduces every geometry field structurally (AC-26)", async () => {
    // The first save assigns generated automatic-style names ("P1") that the
    // reader records as styleIds, so the AC's equality is between the opened
    // and re-opened models: open(save(model)) vs open(save(open(save(model)))).
    const opened = await parseOdt((await writeOdt(geometryModel())).bytes);
    const openedModel = { ...createDocument("geometry"), sections: opened.sections, defaultStyles: opened.defaultStyles };
    const reopened = await parseOdt((await writeOdt(openedModel)).bytes);
    expect(firstStructuralDifference(opened.sections, reopened.sections)).toBeNull();
    const rtl = reopened.sections[0].blocks[0] as Paragraph;
    expect(rtl.formatting?.direction).toBe("rtl");
    expect(rtl.formatting?.indentStartCm).toBe(2.54);
    expect(rtl.formatting?.indentEndCm).toBe(1.27);
    const exact = reopened.sections[0].blocks[1] as Paragraph;
    expect(exact.formatting?.lineSpacingPt).toBe(14);
    expect(exact.formatting?.lineSpacingRule).toBe("exact");
    // An explicit 0 pt space-before survives — it must keep overriding a
    // style's value after reload (AC-21).
    expect(exact.formatting?.spaceBeforePt).toBe(0);
    const atLeast = reopened.sections[0].blocks[2] as Paragraph;
    expect(atLeast.formatting?.lineSpacingPt).toBe(16);
    expect(atLeast.formatting?.lineSpacingRule).toBeUndefined();
  });
});

function textParagraph(text: string): Paragraph {
  return { type: "paragraph", runs: [{ text, formatting: {} }] };
}

/** open(save(model)) — the canonical opened form — followed by a further
 * save+open; SPEC-028's odt mappings must make the second cycle a fixed
 * point even where the first canonicalises (break blocks, carrier
 * paragraphs, explicit section geometry). */
async function expectSecondCycleFixedPoint(model: DocumentModel) {
  const opened = await parseOdt((await writeOdt(model)).bytes);
  const openedModel = { ...createDocument(model.id), sections: opened.sections, defaultStyles: opened.defaultStyles };
  const reopened = await parseOdt((await writeOdt(openedModel)).bytes);
  expect(firstStructuralDifference(opened.sections, reopened.sections)).toBeNull();
  expect(reopened.defaultStyles).toEqual(opened.defaultStyles);
  return opened;
}

describe("writeOdt — pagination flags (SPEC-028 AC-22..29)", () => {
  function flaggedModel(): DocumentModel {
    const model = createDocument("flags");
    model.sections[0].blocks = [
      {
        ...textParagraph("keep me"),
        formatting: { keepNext: true, keepLines: true, widowControl: false },
      },
    ];
    return model;
  }

  it("emits the flags as fo: attributes on the automatic paragraph style", async () => {
    const contentXml = await partText((await writeOdt(flaggedModel())).bytes, "content.xml");
    expect(contentXml).toContain('fo:keep-with-next="always"');
    expect(contentXml).toContain('fo:keep-together="always"');
    expect(contentXml).toContain('fo:widows="0"');
    expect(contentXml).toContain('fo:orphans="0"');
  });

  it("round-trips keepNext, keepLines and widowControl; second cycle is a fixed point", async () => {
    const opened = await expectSecondCycleFixedPoint(flaggedModel());
    const formatting = asParagraph(opened.sections[0].blocks[0]).formatting;
    expect(formatting?.keepNext).toBe(true);
    expect(formatting?.keepLines).toBe(true);
    expect(formatting?.widowControl).toBe(false);
    expect(formatting?.pageBreakBefore).toBeUndefined();
  });
});

describe("writeOdt — manual breaks (SPEC-028 AC-32)", () => {
  it("writes a BreakBlock as fo:break-before='page' on the next paragraph's style, no extra block", async () => {
    const model = createDocument("break");
    model.sections[0].blocks = [
      textParagraph("before"),
      { type: "break", kind: "page" },
      textParagraph("after"),
    ];
    const contentXml = await partText((await writeOdt(model)).bytes, "content.xml");
    expect(contentXml).toContain('fo:break-before="page"');

    const opened = await parseOdt((await writeOdt(model)).bytes);
    const blocks = opened.sections[0].blocks;
    expect(blocks.map((block) => block.type)).toEqual(["paragraph", "break", "paragraph"]);
    expect(blocks[1]).toEqual({ type: "break", kind: "page" });
    // The break came back as the block, not as pageBreakBefore residue.
    expect(asParagraph(blocks[2]).formatting?.pageBreakBefore).toBeUndefined();
    await expectSecondCycleFixedPoint(model);
  });

  it("round-trips a column BreakBlock as fo:break-before='column' (AC-10/11)", async () => {
    const model = createDocument("column-break");
    model.sections[0].blocks = [
      textParagraph("before"),
      { type: "break", kind: "column" },
      textParagraph("after"),
    ];
    const contentXml = await partText((await writeOdt(model)).bytes, "content.xml");
    expect(contentXml).toContain('fo:break-before="column"');

    const opened = await parseOdt((await writeOdt(model)).bytes);
    const blocks = opened.sections[0].blocks;
    expect(blocks.map((block) => block.type)).toEqual(["paragraph", "break", "paragraph"]);
    expect(blocks[1]).toEqual({ type: "break", kind: "column" });
    await expectSecondCycleFixedPoint(model);
  });

  it("canonicalises pageBreakBefore FORMATTING to a break block through .odt (documented convergence)", async () => {
    // ODF has one channel for both concepts (fo:break-before), so a direct
    // pageBreakBefore flag re-opens as a manual BreakBlock — identical
    // pagination, and the second cycle is a fixed point.
    const model = createDocument("break-formatting");
    model.sections[0].blocks = [
      textParagraph("first"),
      { ...textParagraph("second"), formatting: { pageBreakBefore: true } },
    ];
    const contentXml = await partText((await writeOdt(model)).bytes, "content.xml");
    expect(contentXml).toContain('fo:break-before="page"');

    const opened = await parseOdt((await writeOdt(model)).bytes);
    const blocks = opened.sections[0].blocks;
    expect(blocks.map((block) => block.type)).toEqual(["paragraph", "break", "paragraph"]);
    expect(asParagraph(blocks[2]).formatting?.pageBreakBefore).toBeUndefined();
    await expectSecondCycleFixedPoint(model);
  });

  it("keeps a trailing break alive via an empty carrier paragraph", async () => {
    const model = createDocument("trailing-break");
    model.sections[0].blocks = [textParagraph("body"), { type: "break", kind: "page" }];
    const opened = await parseOdt((await writeOdt(model)).bytes);
    const blocks = opened.sections[0].blocks;
    expect(blocks.map((block) => block.type)).toEqual(["paragraph", "break", "paragraph"]);
    expect(asParagraph(blocks[2]).runs).toEqual([]);
    await expectSecondCycleFixedPoint(model);
  });

  it("carries a break before a table on an empty carrier paragraph", async () => {
    const model = createDocument("break-table");
    model.sections[0].blocks = [
      textParagraph("before"),
      { type: "break", kind: "page" },
      {
        type: "table",
        rows: [{ cells: [{ blocks: [textParagraph("cell")] }] }],
        columnWidthsCm: [5],
      },
    ];
    const opened = await parseOdt((await writeOdt(model)).bytes);
    const blocks = opened.sections[0].blocks;
    // The carrier re-opens as an empty paragraph between break and table —
    // the documented convergence for a break directly before a table.
    expect(blocks.map((block) => block.type)).toEqual(["paragraph", "break", "paragraph", "table"]);
    expect(blocks[1]).toEqual({ type: "break", kind: "page" });
  });
});

describe("writeOdt — sections via master pages (SPEC-028 AC-33)", () => {
  function twoSectionModel(): DocumentModel {
    const model = createDocument("sections");
    model.sections = [
      { blocks: [textParagraph("first section")] },
      { blocks: [textParagraph("second section")], pageSize: { widthCm: 14.8, heightCm: 21 } },
    ];
    return model;
  }

  it("emits a page layout + master page per later section, referenced from its first paragraph", async () => {
    const bytes = (await writeOdt(twoSectionModel())).bytes;
    const stylesXml = await partText(bytes, "styles.xml");
    expect(stylesXml).toContain('<style:master-page style:name="SectionMaster2" style:page-layout-name="pm2"/>');
    expect(stylesXml).toContain('<style:page-layout style:name="pm2">');
    expect(stylesXml).toContain('fo:page-width="14.8cm"');
    // The first section keeps the conventional pair.
    expect(stylesXml).toContain('<style:master-page style:name="Standard" style:page-layout-name="pm1"/>');
    const contentXml = await partText(bytes, "content.xml");
    expect(contentXml).toContain('style:master-page-name="SectionMaster2"');
  });

  it("re-opens as two sections, the second with its own explicit A5 geometry", async () => {
    const opened = await parseOdt((await writeOdt(twoSectionModel())).bytes);
    expect(opened.sections).toHaveLength(2);
    expect(asParagraph(opened.sections[0].blocks[0]).runs[0].text).toBe("first section");
    expect(asParagraph(opened.sections[1].blocks[0]).runs[0].text).toBe("second section");
    expect(opened.sections[1].pageSize).toEqual({ widthCm: 14.8, heightCm: 21 });
    // Margins were inherited at write time, so the master pinned the then-
    // effective values — explicit on re-open.
    expect(opened.sections[1].margins).toEqual(opened.defaultStyles.margins);
    expect(opened.sections[0].pageSize).toBeUndefined();
    await expectSecondCycleFixedPoint(twoSectionModel());
  });

  it("re-opens a continuous section as nextPage — ODF has no continuous concept (documented loss)", async () => {
    const model = twoSectionModel();
    model.sections[1].start = "continuous";
    const opened = await parseOdt((await writeOdt(model)).bytes);
    expect(opened.sections).toHaveLength(2);
    expect(opened.sections[1].start).toBeUndefined();
  });
});
