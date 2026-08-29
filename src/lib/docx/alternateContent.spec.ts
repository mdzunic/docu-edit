// Word writes a drawing bare only when its markup predates the 2010
// extensions. A shape, a text box, a group or a picture with relative
// sizing is wrapped in `mc:AlternateContent` — an `mc:Choice` holding the
// modern `w:drawing` plus an `mc:Fallback` holding legacy VML — so a reader
// that only looks at a run's direct children sees the wrapper instead of the
// drawing inside it and preserves the whole run opaquely, which renders as
// nothing. These cover the run-level MCE unwrap that fixes that, and the
// preservation guarantee that has to survive it: a wrapper nothing models
// (a `wpg:wgp` group, a drawing canvas) is still kept verbatim.
import { describe, it, expect } from "vitest";
import { parseDocx, type ParsedDocxDocument } from "./reader";
import { writeDocx } from "./writer";
import { unzip } from "../core/zip";
import type { DocumentModel, Image, Shape } from "../documentModel";

const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

// A 1x1 PNG — enough for the media/dimension decode the picture path runs.
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

// The namespace declarations Word puts on `w:document`, including every
// prefix an `mc:Choice/@Requires` below names — MCE resolves those prefixes
// through the document's own declarations, so a fixture that omits them is
// testing the wrong thing.
const DOCUMENT_NAMESPACES =
  ` xmlns:w="${W}"` +
  ' xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"' +
  ' xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"' +
  ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"' +
  ' xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"' +
  ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"' +
  ' xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"' +
  ' xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"' +
  ' xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"' +
  ' xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"' +
  ' xmlns:v="urn:schemas-microsoft-com:vml"' +
  ' mc:Ignorable="wp14"';

const ANCHOR_ATTRS =
  'distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="251658240"' +
  ' behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"';
const ANCHOR_POSITION =
  '<wp:simplePos x="0" y="0"/>' +
  '<wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>' +
  '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>';

/** A floating picture drawing, 1cm x 1cm, wrap-square. */
function pictureDrawing(id: number): string {
  return (
    `<w:drawing><wp:anchor ${ANCHOR_ATTRS}>${ANCHOR_POSITION}` +
    '<wp:extent cx="360000" cy="360000"/><wp:wrapSquare wrapText="bothSides"/>' +
    `<wp:docPr id="${id}" name="Picture ${id}"/><wp:cNvGraphicFramePr/>` +
    '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
    `<pic:pic><pic:nvPicPr><pic:cNvPr id="${id}" name="Picture ${id}"/><pic:cNvPicPr/></pic:nvPicPr>` +
    '<pic:blipFill><a:blip r:embed="rId5"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>' +
    '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="360000" cy="360000"/></a:xfrm>' +
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>' +
    "</a:graphicData></a:graphic></wp:anchor></w:drawing>"
  );
}

/** A floating `wps:wsp` rectangle drawing, 1cm x 1cm, filled 4472C4. */
function shapeDrawing(id: number): string {
  return (
    `<w:drawing><wp:anchor ${ANCHOR_ATTRS}>${ANCHOR_POSITION}` +
    '<wp:extent cx="360000" cy="360000"/><wp:wrapNone/>' +
    `<wp:docPr id="${id}" name="Rectangle ${id}"/><wp:cNvGraphicFramePr/>` +
    '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
    '<wps:wsp><wps:cNvSpPr/><wps:spPr>' +
    '<a:xfrm><a:off x="0" y="0"/><a:ext cx="360000" cy="360000"/></a:xfrm>' +
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>' +
    '<a:solidFill><a:srgbClr val="4472C4"/></a:solidFill>' +
    "</wps:spPr><wps:bodyPr/></wps:wsp>" +
    "</a:graphicData></a:graphic></wp:anchor></w:drawing>"
  );
}

/** A `wpc:wpc` drawing canvas — markup no reader models, and (unlike a
 * `wpg:wgp` group, which group.ts now flattens) none is planned. */
function canvasDrawing(id: number): string {
  return (
    `<w:drawing><wp:anchor ${ANCHOR_ATTRS}>${ANCHOR_POSITION}` +
    '<wp:extent cx="720000" cy="360000"/><wp:wrapNone/>' +
    `<wp:docPr id="${id}" name="Canvas ${id}"/><wp:cNvGraphicFramePr/>` +
    '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas">' +
    "<wpc:wpc><wpc:bg/><wpc:whole/>" +
    '<wps:wsp><wps:cNvSpPr/><wps:spPr>' +
    '<a:xfrm><a:off x="0" y="0"/><a:ext cx="360000" cy="360000"/></a:xfrm>' +
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr><wps:bodyPr/></wps:wsp>' +
    "</wpc:wpc></a:graphicData></a:graphic></wp:anchor></w:drawing>"
  );
}

/** Wraps a drawing the way Word does: an `mc:Choice` gated on `requires`,
 * plus the legacy VML `mc:Fallback` older Word reads instead. */
function alternateContent(requires: string, drawing: string, fallbackId: string): string {
  return (
    `<mc:AlternateContent><mc:Choice Requires="${requires}">${drawing}</mc:Choice>` +
    `<mc:Fallback><w:pict><v:rect id="${fallbackId}" style="width:28pt;height:28pt"/></w:pict></mc:Fallback>` +
    "</mc:AlternateContent>"
  );
}

const XML_HEADER = `${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`;

/** A minimal package whose body is the given run contents, one per
 * paragraph, followed by a plain "end" paragraph so a dropped block is
 * visible as a missing block rather than a shorter document. */
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
      `${XML_HEADER}<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`,
    ),
  );
  parts.set(
    "word/document.xml",
    encoder.encode(`${XML}<w:document${DOCUMENT_NAMESPACES}><w:body>${body}</w:body></w:document>`),
  );
  parts.set(
    "word/_rels/document.xml.rels",
    encoder.encode(
      `${XML_HEADER}<Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/></Relationships>`,
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

/** The `word/document.xml` this model writes back out — the check that
 * preservation survived the read, not just that the model held the XML. */
async function savedDocumentXml(parsed: ParsedDocxDocument): Promise<string> {
  const model: DocumentModel = {
    id: "test-doc",
    title: "AlternateContent",
    sections: parsed.sections,
    defaultStyles: parsed.defaultStyles,
    format: "docx",
    opaque: parsed.opaque,
    sourceXml: parsed.sourceXml,
  };
  const { bytes } = await writeDocx(model);
  return new TextDecoder().decode((await unzip(bytes)).get("word/document.xml")!);
}

describe("parseDocx — mc:AlternateContent-wrapped drawings", () => {
  it("reads a picture out of the mc:Choice Word wraps it in", async () => {
    const blocks = await blocksOf([alternateContent("wp14", pictureDrawing(1), "_x0000_s1026")]);
    expect(blocks.map((block) => block.type)).toEqual(["image", "paragraph"]);
    const image = blocks[0] as Image;
    expect(image.mediaId).toBe("word/media/image1.png");
    expect(image.wrap).toBe("square");
  });

  it("reads a wps:wsp shape out of the mc:Choice — the form Word always writes", async () => {
    const blocks = await blocksOf([alternateContent("wps", shapeDrawing(2), "_x0000_s1027")]);
    expect(blocks.map((block) => block.type)).toEqual(["shape", "paragraph"]);
    const shape = blocks[0] as Shape;
    expect(shape.geometry).toBe("rect");
    expect(shape.fill).toBe("#4472c4");
  });

  it("reads a wrapped drawing identically to the same drawing written bare", async () => {
    const wrapped = await blocksOf([alternateContent("wps", shapeDrawing(3), "_x0000_s1028")]);
    const bare = await blocksOf([shapeDrawing(3)]);
    expect(wrapped).toEqual(bare);
  });

  it("keeps reading a drawing that sits beside an unrelated foreign element", async () => {
    // The filter this unwrap replaced ignored every non-`w:` child, so a run
    // carrying a drawing plus some producer's own extension element still
    // read as a drawing. Widening the unwrap must not narrow that.
    const blocks = await blocksOf([`${pictureDrawing(14)}<v:shapetype id="_x0000_t75"/>`]);
    expect(blocks.map((block) => block.type)).toEqual(["image", "paragraph"]);
  });

  it("preserves a run holding two drawings rather than reading one and losing the other", async () => {
    const blocks = await blocksOf([
      alternateContent("wps", shapeDrawing(15), "_x0000_s1035") + pictureDrawing(16),
    ]);
    expect(blocks.map((block) => block.type)).toEqual(["paragraph", "paragraph"]);
    const preserved = blocks[0];
    const xml = preserved.type === "paragraph" ? (preserved.opaqueInlines ?? []).map((i) => i.xml).join("") : "";
    expect(xml).toContain("AlternateContent");
    expect(xml).toContain("Picture 16");
  });

  it("still reads bare drawings, wrapped and bare drawings side by side", async () => {
    const blocks = await blocksOf([
      pictureDrawing(4),
      alternateContent("wps", shapeDrawing(5), "_x0000_s1029"),
      shapeDrawing(6),
      alternateContent("wp14", pictureDrawing(7), "_x0000_s1030"),
    ]);
    expect(blocks.map((block) => block.type)).toEqual(["image", "shape", "shape", "image", "paragraph"]);
  });
});

describe("parseDocx — mc:AlternateContent nothing models", () => {
  // A drawing canvas has no reader, so its Choice must NOT be selected: MCE
  // then says to read the Fallback, which is legacy VML nothing models
  // either, and the run falls through to the opaque-inline path — the whole
  // wrapper preserved verbatim, exactly as before the unwrap existed. Adding
  // a namespace to UNDERSTOOD_DRAWING_NAMESPACES before a reader exists for
  // it would select a branch nothing can read.
  it("preserves a wpc:wpc canvas wrapper verbatim rather than dropping it", async () => {
    const blocks = await blocksOf([alternateContent("wpc", canvasDrawing(8), "_x0000_s1031")]);
    expect(blocks.map((block) => block.type)).toEqual(["paragraph", "paragraph"]);
    const preserved = blocks[0];
    expect(preserved.type).toBe("paragraph");
    const xml = preserved.type === "paragraph" ? (preserved.opaqueInlines ?? []).map((i) => i.xml).join("") : "";
    expect(xml).toContain("AlternateContent");
    expect(xml).toContain("wpc");
    // The VML fallback older Word reads must survive too.
    expect(xml).toContain("_x0000_s1031");
  });

  it("preserves a Choice whose Requires names a prefix the document never declares", async () => {
    // MCE resolves Requires through the document's own declarations; an
    // undeclared prefix cannot be evaluated, so the Fallback is read and
    // nothing is modelled.
    const blocks = await blocksOf([alternateContent("wpz", shapeDrawing(9), "_x0000_s1032")]);
    expect(blocks.map((block) => block.type)).toEqual(["paragraph", "paragraph"]);
    const preserved = blocks[0];
    const xml = preserved.type === "paragraph" ? (preserved.opaqueInlines ?? []).map((i) => i.xml).join("") : "";
    expect(xml).toContain("AlternateContent");
  });

  it("reads the mc:Fallback when it holds a drawing a reader does model", async () => {
    // Word's own shape for a wp14-only feature: the Choice carries the
    // extended markup, the Fallback the same drawing without it. A reader
    // that understands neither the Choice nor VML still gets the picture.
    const wrapped =
      '<mc:AlternateContent><mc:Choice Requires="wpz">' +
      shapeDrawing(10) +
      `</mc:Choice><mc:Fallback>${pictureDrawing(11)}</mc:Fallback></mc:AlternateContent>`;
    const blocks = await blocksOf([wrapped]);
    expect(blocks.map((block) => block.type)).toEqual(["image", "paragraph"]);
  });
});

describe("writeDocx — after an mc:AlternateContent read", () => {
  it("writes an unmodelled canvas wrapper back byte-for-byte, VML fallback included", async () => {
    const parsed = await parse([alternateContent("wpc", canvasDrawing(12), "_x0000_s1033")]);
    const xml = await savedDocumentXml(parsed);
    expect(xml).toContain("AlternateContent");
    expect(xml).toContain("wpc");
    expect(xml).toContain("_x0000_s1033");
  });

  it("round-trips a shape read out of a Choice to the same model", async () => {
    const parsed = await parse([alternateContent("wps", shapeDrawing(13), "_x0000_s1034")]);
    const xml = await savedDocumentXml(parsed);
    const { bytes } = await writeDocx({
      id: "test-doc",
      title: "AlternateContent",
      sections: parsed.sections,
      defaultStyles: parsed.defaultStyles,
      format: "docx",
      opaque: parsed.opaque,
      sourceXml: parsed.sourceXml,
    });
    const reread = await parseDocx(bytes);
    expect(reread.sections[0].blocks.map((block) => block.type)).toEqual(["shape", "paragraph"]);
    const before = parsed.sections[0].blocks[0] as Shape;
    const after = reread.sections[0].blocks[0] as Shape;
    expect(after.geometry).toBe(before.geometry);
    expect(after.fill).toBe(before.fill);
    expect(after.widthCm).toBeCloseTo(before.widthCm, 5);
    // The wrapper is not re-emitted: the shape is modelled now, so it is
    // written as this project's own bare `w:drawing`.
    expect(xml).not.toContain("AlternateContent");
  });
});
