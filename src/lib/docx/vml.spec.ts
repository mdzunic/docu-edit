// Legacy VML pictures (`w:pict`) — the markup every Word 97-2003 document
// keeps after conversion to .docx, and what some LibreOffice/Google Docs
// exports still write. `w:pict` has never been a modelled run child, so a run
// holding one went to the opaque-inline path: round-tripped, but rendered as
// nothing, which is why an old document's pictures simply did not appear.
import { describe, it, expect } from "vitest";
import { parseDocx, type ParsedDocxDocument } from "./reader";
import type { Image } from "../documentModel";

const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const PNG_BASE64 =
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";

function base64ToBytes(base64: string): Uint8Array {
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) {
    bytes[i] = binary.charCodeAt(i);
  }
  return bytes;
}

const DOCUMENT_NAMESPACES =
  ` xmlns:w="${W}"` +
  ' xmlns:v="urn:schemas-microsoft-com:vml"' +
  ' xmlns:o="urn:schemas-microsoft-com:office:office"' +
  ' xmlns:w10="urn:schemas-microsoft-com:office:word"' +
  ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';

/** The `v:shapetype` Word emits once beside its pictures — a shared shape
 * definition carrying no picture of its own. */
const SHAPETYPE = '<v:shapetype id="_x0000_t75" coordsize="21600,21600" o:spt="75"/>';

const RELS_HEADER = `${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`;

function packageParts(runContents: string[]): Map<string, Uint8Array> {
  const encoder = new TextEncoder();
  const body =
    runContents.map((content) => `<w:p><w:r>${content}</w:r></w:p>`).join("") +
    "<w:p><w:r><w:t>end</w:t></w:r></w:p>" +
    '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>' +
    '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>' +
    "</w:sectPr>";
  const parts = new Map<string, Uint8Array>();
  parts.set(
    "[Content_Types].xml",
    encoder.encode(
      `${XML}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        '<Default Extension="png" ContentType="image/png"/>' +
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
        "</Types>",
    ),
  );
  parts.set(
    "_rels/.rels",
    encoder.encode(
      `${RELS_HEADER}<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`,
    ),
  );
  parts.set(
    "word/document.xml",
    encoder.encode(`${XML}<w:document${DOCUMENT_NAMESPACES}><w:body>${body}</w:body></w:document>`),
  );
  parts.set(
    "word/_rels/document.xml.rels",
    encoder.encode(
      `${RELS_HEADER}<Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/></Relationships>`,
    ),
  );
  parts.set("word/media/image1.png", base64ToBytes(PNG_BASE64));
  return parts;
}

async function parse(runContents: string[]): Promise<ParsedDocxDocument> {
  return parseDocx(new Uint8Array(), packageParts(runContents));
}

async function blocksOf(runContents: string[]) {
  return (await parse(runContents)).sections[0].blocks;
}

function preservedXml(block: { type: string }): string {
  const paragraph = block as { type: string; opaqueInlines?: { xml: string }[] };
  return paragraph.type === "paragraph" ? (paragraph.opaqueInlines ?? []).map((i) => i.xml).join("") : "";
}

const at = (zIndex: string) =>
  `<w:pict><v:shape style="position:absolute;width:36pt;height:36pt;z-index:${zIndex}">` +
  '<v:imagedata r:id="rId5"/></v:shape></w:pict>';

describe("parseDocx — legacy VML pictures", () => {
  it("reads an inline v:shape/v:imagedata as an Image at its style's size", async () => {
    const pict =
      `<w:pict>${SHAPETYPE}` +
      '<v:shape id="_x0000_i1025" type="#_x0000_t75" style="width:99.75pt;height:75pt">' +
      '<v:imagedata r:id="rId5" o:title="Old photo"/></v:shape></w:pict>';
    const blocks = await blocksOf([pict]);
    expect(blocks.map((block) => block.type)).toEqual(["image", "paragraph"]);
    const image = blocks[0] as Image;
    expect(image.mediaId).toBe("word/media/image1.png");
    expect(image.wrap).toBe("inline");
    // 99.75pt / 75pt at 72pt per inch.
    expect(image.widthCm).toBeCloseTo((99.75 / 72) * 2.54, 6);
    expect(image.heightCm).toBeCloseTo((75 / 72) * 2.54, 6);
    // `o:title` is VML's alt text.
    expect(image.altText).toBe("Old photo");
  });

  it("registers the picture's media part", async () => {
    const pict =
      '<w:pict><v:shape style="width:36pt;height:36pt"><v:imagedata r:id="rId5"/></v:shape></w:pict>';
    const parsed = await parse([pict]);
    expect(parsed.media.map((part) => part.mediaId)).toEqual(["word/media/image1.png"]);
  });

  it("falls back to the older `alt` spelling Word 2003 wrote", async () => {
    const pict =
      '<w:pict><v:shape style="width:36pt;height:36pt" alt="Legacy alt">' +
      '<v:imagedata r:id="rId5"/></v:shape></w:pict>';
    expect(((await blocksOf([pict]))[0] as Image).altText).toBe("Legacy alt");
  });

  it("reads w10:wrap square as square wrap", async () => {
    const pict =
      '<w:pict><v:shape style="position:absolute;width:56.7pt;height:56.7pt">' +
      '<v:imagedata r:id="rId5"/><w10:wrap type="square"/></v:shape></w:pict>';
    expect(((await blocksOf([pict]))[0] as Image).wrap).toBe("square");
  });

  it("splits behind from in-front by VML's own z-index sign", async () => {
    expect(((await blocksOf([at("-1")]))[0] as Image).wrap).toBe("behind");
    expect(((await blocksOf([at("2")]))[0] as Image).wrap).toBe("infront");
  });

  it("accepts the other absolute CSS units a converted document may carry", async () => {
    const pict =
      '<w:pict><v:shape style="width:2cm;height:20mm"><v:imagedata r:id="rId5"/></v:shape></w:pict>';
    const image = (await blocksOf([pict]))[0] as Image;
    expect(image.widthCm).toBeCloseTo(2, 6);
    expect(image.heightCm).toBeCloseTo(2, 6);
  });

  it("reads a picture beside ordinary text and bare DrawingML in one document", async () => {
    const pict =
      '<w:pict><v:shape style="width:36pt;height:36pt"><v:imagedata r:id="rId5"/></v:shape></w:pict>';
    const blocks = await blocksOf([pict, "<w:t>between</w:t>", pict]);
    expect(blocks.map((block) => block.type)).toEqual(["image", "paragraph", "image", "paragraph"]);
  });
});

describe("parseDocx — w:pict this reader does not model stays preserved", () => {
  // VML's fill/stroke/geometry model is a second shape language; half-reading
  // it would paint shapes this module cannot describe. Everything below keeps
  // the verbatim preservation it had before VML pictures were read at all.
  it("preserves a VML shape that carries no picture", async () => {
    const pict = '<w:pict><v:rect id="_x0000_s1027" style="width:100pt;height:50pt" fillcolor="#4472c4"/></w:pict>';
    const blocks = await blocksOf([pict]);
    expect(blocks.map((block) => block.type)).toEqual(["paragraph", "paragraph"]);
    expect(preservedXml(blocks[0])).toContain("v:rect");
  });

  it("preserves a horizontal rule (o:hr) rather than resolving it to a picture", async () => {
    const pict = '<w:pict><v:rect id="_x0000_s1028" o:hr="t" style="width:100%;height:1.5pt"/></w:pict>';
    const blocks = await blocksOf([pict]);
    expect(blocks.map((block) => block.type)).toEqual(["paragraph", "paragraph"]);
    // The serializer is free to rebind the `o:` prefix, so match the
    // attribute's local name rather than its source spelling.
    expect(preservedXml(blocks[0])).toMatch(/:hr="t"/);
  });

  it("preserves a LINKED image, whose bytes are not in the package", async () => {
    // `r:href` instead of `r:id`: the model only carries embedded media.
    const pict =
      '<w:pict><v:shape style="width:36pt;height:36pt">' +
      '<v:imagedata r:href="rId9"/></v:shape></w:pict>';
    expect((await blocksOf([pict])).map((block) => block.type)).toEqual(["paragraph", "paragraph"]);
  });

  it("preserves a contour wrap the model has no mode for", async () => {
    const pict =
      '<w:pict><v:shape style="position:absolute;width:36pt;height:36pt">' +
      '<v:imagedata r:id="rId5"/><w10:wrap type="tight"/></v:shape></w:pict>';
    expect((await blocksOf([pict])).map((block) => block.type)).toEqual(["paragraph", "paragraph"]);
  });

  it("preserves a picture sized in percentages or VML's own coordinate units", async () => {
    const percent =
      '<w:pict><v:shape style="width:50%;height:36pt"><v:imagedata r:id="rId5"/></v:shape></w:pict>';
    const bare = '<w:pict><v:shape style="width:200;height:100"><v:imagedata r:id="rId5"/></v:shape></w:pict>';
    expect((await blocksOf([percent])).map((block) => block.type)).toEqual(["paragraph", "paragraph"]);
    expect((await blocksOf([bare])).map((block) => block.type)).toEqual(["paragraph", "paragraph"]);
  });

  it("preserves a w:pict holding several shapes", async () => {
    const pict =
      '<w:pict><v:shape style="width:36pt;height:36pt"><v:imagedata r:id="rId5"/></v:shape>' +
      '<v:shape style="width:36pt;height:36pt"><v:imagedata r:id="rId5"/></v:shape></w:pict>';
    expect((await blocksOf([pict])).map((block) => block.type)).toEqual(["paragraph", "paragraph"]);
  });
});
