// SPEC-025 AC-31, AC-32: hyperlinks and bookmarks through the .odt reader
// and writer — text:a with xlink:href (external or "#anchor"), and
// text:bookmark point anchors — plus the security stripping of dangerous
// URL schemes on open (AC-17, AC-18).
import { describe, expect, it } from "vitest";
import { parseOdt } from "./reader";
import { buildContentXml, writeOdt } from "./writer";
import { zip } from "../docx/zipWriter";
import { unzip } from "../docx/zip";
import { firstStructuralDifference } from "../equality";
import {
  asParagraph,
  createDocument,
  type DocumentModel,
  type Paragraph,
} from "../documentModel";

const encoder = new TextEncoder();
const decoder = new TextDecoder();

const CONTENT_OPEN =
  '<?xml version="1.0" encoding="UTF-8"?>' +
  '<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"' +
  ' xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"' +
  ' xmlns:xlink="http://www.w3.org/1999/xlink" office:version="1.3">' +
  "<office:body><office:text>";
const CONTENT_CLOSE = "</office:text></office:body></office:document-content>";

async function makeOdt(bodyXml: string): Promise<Uint8Array> {
  return zip([
    {
      name: "mimetype",
      data: encoder.encode("application/vnd.oasis.opendocument.text"),
      store: true,
    },
    { name: "content.xml", data: encoder.encode(CONTENT_OPEN + bodyXml + CONTENT_CLOSE) },
  ]);
}

function runsOf(model: { sections: DocumentModel["sections"] }, block = 0): Paragraph["runs"] {
  return asParagraph(model.sections[0].blocks[block]).runs;
}

describe("odt reader — text:a (AC-31)", () => {
  it("maps an external xlink:href onto linkHref", async () => {
    const parsed = await parseOdt(
      await makeOdt(
        '<text:p>See <text:a xlink:type="simple" xlink:href="https://example.com/">example</text:a></text:p>',
      ),
    );
    expect(runsOf(parsed)).toEqual([
      { text: "See ", formatting: {} },
      { text: "example", formatting: { linkHref: "https://example.com/" } },
    ]);
  });

  it('maps a "#fragment" href onto linkAnchor', async () => {
    const parsed = await parseOdt(
      await makeOdt('<text:p><text:a xlink:href="#mark">go</text:a></text:p>'),
    );
    expect(runsOf(parsed)).toEqual([{ text: "go", formatting: { linkAnchor: "mark" } }]);
  });

  it("strips a dangerous scheme on open, keeping the text (AC-17)", async () => {
    const parsed = await parseOdt(
      await makeOdt('<text:p><text:a xlink:href="javascript:alert(1)">evil</text:a></text:p>'),
    );
    expect(runsOf(parsed)).toEqual([{ text: "evil", formatting: {} }]);
  });
});

describe("odt reader — text:bookmark (AC-32)", () => {
  it("reads point bookmarks at their character offsets", async () => {
    const parsed = await parseOdt(
      await makeOdt('<text:p>ab<text:bookmark text:name="mark"/>cd</text:p>'),
    );
    expect(asParagraph(parsed.sections[0].blocks[0]).bookmarks).toEqual([
      { name: "mark", offset: 2 },
    ]);
    expect(runsOf(parsed)).toEqual([{ text: "abcd", formatting: {} }]);
  });

  it("collapses a ranged bookmark to its start", async () => {
    const parsed = await parseOdt(
      await makeOdt(
        '<text:p>a<text:bookmark-start text:name="range"/>bc<text:bookmark-end text:name="range"/>d</text:p>',
      ),
    );
    const paragraph = asParagraph(parsed.sections[0].blocks[0]);
    expect(paragraph.bookmarks).toEqual([{ name: "range", offset: 1 }]);
    expect(paragraph.runs).toEqual([{ text: "abcd", formatting: {} }]);
    expect(paragraph.opaqueInlines).toBeUndefined();
  });
});

describe("odt writer — emission (AC-31, AC-32)", () => {
  it("groups same-target runs under one text:a with the xlink namespace on the paragraph", () => {
    const model = createDocument("odt-links");
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
    const xml = buildContentXml(model);
    expect(xml).toContain('xmlns:xlink="http://www.w3.org/1999/xlink"');
    expect(xml).toContain('<text:a xlink:type="simple" xlink:href="https://example.com/">');
    // one wrapper covers both runs; a paragraph declares xlink once
    expect(xml.match(/<text:a /g)).toHaveLength(1);
    expect(xml.match(/xmlns:xlink/g)).toHaveLength(1);
  });

  it("emits internal links as #anchor and bookmarks as text:bookmark", () => {
    const model = createDocument("odt-anchors");
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
    const xml = buildContentXml(model);
    expect(xml).toContain('<text:bookmark text:name="mark"/>');
    expect(xml).toContain('<text:a xlink:type="simple" xlink:href="#mark">');
  });

  it("does not change output for documents without links or bookmarks", () => {
    const model = createDocument("odt-plain");
    model.sections = [
      { blocks: [{ type: "paragraph", runs: [{ text: "plain", formatting: {} }] }] },
    ];
    const xml = buildContentXml(model);
    expect(xml).not.toContain("xmlns:xlink");
    expect(xml).not.toContain("<text:a");
  });
});

describe("odt round trip (AC-31, AC-32)", () => {
  it("round-trips external links, internal links and bookmarks", async () => {
    const model = createDocument("odt-roundtrip-links");
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
    const { bytes, dropped } = await writeOdt(model);
    expect(dropped).toEqual([]);
    const reopened = await parseOdt(bytes);
    const runs = runsOf(reopened);
    expect(runs.map((run) => run.text).join("")).toBe("Visit example or jump here");
    expect(runs.find((run) => run.text === "example")!.formatting.linkHref).toBe(
      "https://example.com/",
    );
    expect(runs.find((run) => run.text === "here")!.formatting.linkAnchor).toBe("mark");
    expect(asParagraph(reopened.sections[0].blocks[1]).bookmarks).toEqual([
      { name: "mark", offset: 7 },
    ]);
    // a second save→open is structurally lossless
    const model2: DocumentModel = {
      ...createDocument("odt-roundtrip-links-2"),
      sections: reopened.sections,
      defaultStyles: reopened.defaultStyles,
      styles: reopened.styles,
      opaque: reopened.opaque,
      sourceXml: reopened.sourceXml,
      format: "odt",
    };
    const again = await parseOdt((await writeOdt(model2)).bytes);
    expect(firstStructuralDifference(reopened.sections, again.sections)).toBeNull();
  });

  it("keeps a bookmark inside a link's text at the right offset", async () => {
    const model = createDocument("odt-bookmark-in-link");
    model.sections = [
      {
        blocks: [
          {
            type: "paragraph",
            runs: [{ text: "linked", formatting: { linkHref: "https://example.com/" } }],
            bookmarks: [{ name: "mid", offset: 3 }],
          },
        ],
      },
    ];
    const { bytes } = await writeOdt(model);
    const parts = await unzip(bytes);
    const content = decoder.decode(parts.get("content.xml"));
    // the bookmark splits the linked run but stays inside one text:a
    expect(content.match(/<text:a /g)).toHaveLength(1);
    const reopened = await parseOdt(bytes);
    expect(asParagraph(reopened.sections[0].blocks[0]).bookmarks).toEqual([
      { name: "mid", offset: 3 },
    ]);
    // the split run may stay fragmented, but text and target are intact
    const runs = runsOf(reopened);
    expect(runs.map((run) => run.text).join("")).toBe("linked");
    expect(runs.every((run) => run.formatting.linkHref === "https://example.com/")).toBe(true);
  });
});
