// SPEC-025 AC-29, AC-30: hyperlinks and bookmarks through the .docx reader
// and writer — w:hyperlink with TargetMode="External" relationships or
// w:anchor, and w:bookmarkStart/End pairs — plus the security stripping of
// dangerous URL schemes on open (AC-17, AC-18).
import { describe, expect, it } from "vitest";
import { parseDocx } from "./reader";
import { buildDocumentXml, planHyperlinkWrite, writeDocx } from "./writer";
import { zip } from "./zipWriter";
import { unzip } from "./zip";
import { firstStructuralDifference } from "../equality";
import {
  asParagraph,
  createDocument,
  type DocumentModel,
  type Paragraph,
} from "../documentModel";

const WORD_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const RELATIONSHIPS_NAMESPACE = "http://schemas.openxmlformats.org/package/2006/relationships";
const R_NAMESPACE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const HYPERLINK_REL_TYPE = `${R_NAMESPACE}/hyperlink`;

const encoder = new TextEncoder();
const decoder = new TextDecoder();

function bodyDocx(body: string): string {
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    `<w:document xmlns:w="${WORD_NAMESPACE}" xmlns:r="${R_NAMESPACE}"><w:body>${body}</w:body></w:document>`
  );
}

function relsXml(rels: string): string {
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    `<Relationships xmlns="${RELATIONSHIPS_NAMESPACE}">${rels}</Relationships>`
  );
}

async function makeDocx(documentXml: string, documentRelsXml?: string): Promise<Uint8Array> {
  const entries = [{ name: "word/document.xml", data: encoder.encode(documentXml) }];
  if (documentRelsXml !== undefined) {
    entries.push({ name: "word/_rels/document.xml.rels", data: encoder.encode(documentRelsXml) });
  }
  return zip(entries);
}

function runsOf(model: { sections: DocumentModel["sections"] }, block = 0): Paragraph["runs"] {
  return asParagraph(model.sections[0].blocks[block]).runs;
}

describe("docx reader — w:hyperlink (AC-29)", () => {
  it("maps an external hyperlink relationship onto linkHref", async () => {
    const bytes = await makeDocx(
      bodyDocx(
        "<w:p><w:r><w:t>See </w:t></w:r>" +
          '<w:hyperlink r:id="rId5"><w:r><w:t>example</w:t></w:r></w:hyperlink>' +
          "</w:p>",
      ),
      relsXml(
        `<Relationship Id="rId5" Type="${HYPERLINK_REL_TYPE}" Target="https://example.com/" TargetMode="External"/>`,
      ),
    );
    const parsed = await parseDocx(bytes);
    expect(runsOf(parsed)).toEqual([
      { text: "See ", formatting: {} },
      { text: "example", formatting: { linkHref: "https://example.com/" } },
    ]);
  });

  it("maps w:anchor onto linkAnchor", async () => {
    const bytes = await makeDocx(
      bodyDocx('<w:p><w:hyperlink w:anchor="target"><w:r><w:t>go</w:t></w:r></w:hyperlink></w:p>'),
    );
    const parsed = await parseDocx(bytes);
    expect(runsOf(parsed)).toEqual([{ text: "go", formatting: { linkAnchor: "target" } }]);
  });

  it("strips a dangerous scheme on open, keeping the text (AC-17)", async () => {
    const bytes = await makeDocx(
      bodyDocx('<w:p><w:hyperlink r:id="rId5"><w:r><w:t>evil</w:t></w:r></w:hyperlink></w:p>'),
      relsXml(
        `<Relationship Id="rId5" Type="${HYPERLINK_REL_TYPE}" Target="javascript:alert(1)" TargetMode="External"/>`,
      ),
    );
    const parsed = await parseDocx(bytes);
    expect(runsOf(parsed)).toEqual([{ text: "evil", formatting: {} }]);
  });

  it("keeps run formatting inside a hyperlink", async () => {
    const bytes = await makeDocx(
      bodyDocx(
        '<w:p><w:hyperlink r:id="rId5">' +
          "<w:r><w:rPr><w:b/></w:rPr><w:t>bold link</w:t></w:r>" +
          "</w:hyperlink></w:p>",
      ),
      relsXml(
        `<Relationship Id="rId5" Type="${HYPERLINK_REL_TYPE}" Target="https://example.com/" TargetMode="External"/>`,
      ),
    );
    const parsed = await parseDocx(bytes);
    expect(runsOf(parsed)).toEqual([
      { text: "bold link", formatting: { bold: true, linkHref: "https://example.com/" } },
    ]);
  });
});

describe("docx reader — w:bookmarkStart/End (AC-30)", () => {
  it("reads a point bookmark at its character offset", async () => {
    const bytes = await makeDocx(
      bodyDocx(
        "<w:p><w:r><w:t>ab</w:t></w:r>" +
          '<w:bookmarkStart w:id="1" w:name="mark"/><w:bookmarkEnd w:id="1"/>' +
          "<w:r><w:t>cd</w:t></w:r></w:p>",
      ),
    );
    const parsed = await parseDocx(bytes);
    expect(asParagraph(parsed.sections[0].blocks[0]).bookmarks).toEqual([
      { name: "mark", offset: 2 },
    ]);
  });

  it("collapses a ranged bookmark to its start and drops _GoBack", async () => {
    const bytes = await makeDocx(
      bodyDocx(
        '<w:p><w:bookmarkStart w:id="1" w:name="range"/>' +
          '<w:bookmarkStart w:id="2" w:name="_GoBack"/><w:bookmarkEnd w:id="2"/>' +
          "<w:r><w:t>spanned</w:t></w:r>" +
          '<w:bookmarkEnd w:id="1"/></w:p>',
      ),
    );
    const parsed = await parseDocx(bytes);
    const paragraph = asParagraph(parsed.sections[0].blocks[0]);
    expect(paragraph.bookmarks).toEqual([{ name: "range", offset: 0 }]);
    expect(paragraph.runs).toEqual([{ text: "spanned", formatting: {} }]);
    expect(paragraph.opaqueInlines).toBeUndefined();
  });
});

describe("docx writer — emission (AC-29, AC-30)", () => {
  it("groups consecutive same-target runs under one w:hyperlink with an external rel", () => {
    const model = createDocument("docx-links");
    model.sections = [
      {
        blocks: [
          {
            type: "paragraph",
            runs: [
              { text: "See ", formatting: {} },
              { text: "the ", formatting: { linkHref: "https://example.com/" } },
              { text: "site", formatting: { linkHref: "https://example.com/", bold: true } },
            ],
          },
        ],
      },
    ];
    const plan = planHyperlinkWrite(model);
    const xml = buildDocumentXml(model, undefined, undefined, plan);
    const relId = plan.relIdFor("https://example.com/");
    expect(relId).not.toBe("");
    expect(xml).toContain(`<w:hyperlink xmlns:r="${R_NAMESPACE}" r:id="${relId}">`);
    // one wrapper for both runs, closed after the second
    expect(xml.match(/<w:hyperlink /g)).toHaveLength(1);
    expect(plan.relationshipsXml).toContain('TargetMode="External"');
    expect(plan.relationshipsXml).toContain('Target="https://example.com/"');
  });

  it("emits internal links as w:anchor and bookmarks as start/end pairs", () => {
    const model = createDocument("docx-anchors");
    model.sections = [
      {
        blocks: [
          {
            type: "paragraph",
            runs: [{ text: "abcd", formatting: {} }],
            bookmarks: [{ name: "mark", offset: 2 }],
          },
          {
            type: "paragraph",
            runs: [{ text: "go", formatting: { linkAnchor: "mark" } }],
          },
        ],
      },
    ];
    const xml = buildDocumentXml(model);
    expect(xml).toContain('<w:bookmarkStart w:id="1" w:name="mark"/><w:bookmarkEnd w:id="1"/>');
    // the bookmark splits the run at its offset
    expect(xml).toContain("<w:t>ab</w:t>");
    expect(xml).toContain("<w:t>cd</w:t>");
    expect(xml).toContain('<w:hyperlink w:anchor="mark">');
  });
});

describe("docx round trip (AC-29, AC-30)", () => {
  it("round-trips external links, internal links and bookmarks", async () => {
    const model = createDocument("docx-roundtrip-links");
    model.sections = [
      {
        blocks: [
          {
            type: "paragraph",
            runs: [
              { text: "Visit ", formatting: {} },
              { text: "example", formatting: { linkHref: "https://example.com/" } },
              { text: " or jump ", formatting: {} },
              { text: "here", formatting: { linkAnchor: "mark" } },
            ],
          },
          {
            type: "paragraph",
            runs: [{ text: "target zone", formatting: {} }],
            bookmarks: [{ name: "mark", offset: 7 }],
          },
        ],
      },
    ];
    const { bytes, dropped } = await writeDocx(model);
    expect(dropped).toEqual([]);
    const reopened = await parseDocx(bytes);
    // links and bookmarks survive the first trip (the reader may enrich runs
    // with default-style formatting, so match the link fields specifically)
    const runs = runsOf(reopened);
    expect(runs.map((run) => run.text).join("")).toBe("Visit example or jump here");
    expect(runs.find((run) => run.text === "example")!.formatting.linkHref).toBe(
      "https://example.com/",
    );
    expect(runs.find((run) => run.text === "here")!.formatting.linkAnchor).toBe("mark");
    expect(asParagraph(reopened.sections[0].blocks[1]).bookmarks).toEqual([
      { name: "mark", offset: 7 },
    ]);
    // and a second write→parse is structurally lossless
    const model2: DocumentModel = {
      ...createDocument("docx-roundtrip-links-2"),
      sections: reopened.sections,
      defaultStyles: reopened.defaultStyles,
      styles: reopened.styles,
      opaque: reopened.opaque,
      sourceXml: reopened.sourceXml,
      format: "docx",
    };
    const again = await parseDocx((await writeDocx(model2)).bytes);
    expect(firstStructuralDifference(reopened.sections, again.sections)).toBeNull();
  });

  it("appends hyperlink relationships additively on a same-format save", async () => {
    const source = await makeDocx(
      bodyDocx(
        "<w:p>" +
          '<w:hyperlink r:id="rId7"><w:r><w:t>kept</w:t></w:r></w:hyperlink>' +
          "</w:p>",
      ),
      relsXml(
        `<Relationship Id="rId7" Type="${HYPERLINK_REL_TYPE}" Target="https://kept.example/" TargetMode="External"/>`,
      ),
    );
    const parsed = await parseDocx(source);
    const model: DocumentModel = {
      ...createDocument("docx-additive-rels"),
      sections: parsed.sections,
      defaultStyles: parsed.defaultStyles,
      opaque: parsed.opaque,
      sourceXml: parsed.sourceXml,
      format: "docx",
    };
    const { bytes } = await writeDocx(model);
    const parts = await unzip(bytes);
    const rels = decoder.decode(parts.get("word/_rels/document.xml.rels"));
    // the preserved rel is intact and the re-emitted link got a fresh id
    expect(rels).toContain('Id="rId7"');
    expect(rels.match(/Target="https:\/\/kept\.example\/"/g)!.length).toBeGreaterThanOrEqual(1);
    const documentXml = decoder.decode(parts.get("word/document.xml"));
    const relIdMatch = /<w:hyperlink [^>]*r:id="(rId\d+)"/.exec(documentXml);
    expect(relIdMatch).not.toBeNull();
    expect(relIdMatch![1]).not.toBe("rId7");
    expect(rels).toContain(`Id="${relIdMatch![1]}"`);
    // and the file still opens with the link intact
    const reopened = await parseDocx(bytes);
    expect(runsOf(reopened)).toEqual([
      { text: "kept", formatting: { linkHref: "https://kept.example/" } },
    ]);
  });
});
