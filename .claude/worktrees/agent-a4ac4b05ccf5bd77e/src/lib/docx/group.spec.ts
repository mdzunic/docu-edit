// Word's real group shape (`wpg:wgp`), flattened into the model's flat
// `groupOrdinal` grouping — see group.ts's header for why flattening rather
// than a new container block, and what that costs.
import { describe, it, expect } from "vitest";
import { parseDocx, type ParsedDocxDocument } from "./reader";
import { writeDocx } from "./writer";
import { unzip } from "../core/zip";
import type { DocumentModel, Image, Shape } from "../documentModel";

const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const PNG_BASE64 =
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";

/** EMU per centimetre — every coordinate below is written in whole cm so the
 * expected page positions read directly off the fixture. */
const CM = 360000;

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
  ' xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"' +
  ' xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"' +
  ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"' +
  ' xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"' +
  ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"' +
  ' xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"' +
  ' xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"' +
  ' xmlns:de="urn:docu-edit:shape"' +
  ' xmlns:v="urn:schemas-microsoft-com:vml"' +
  ' mc:Ignorable="wp14"';

interface Box {
  x: number;
  y: number;
  cx: number;
  cy: number;
}

function xfrm(box: Box, child?: Box): string {
  const ch = child
    ? `<a:chOff x="${child.x}" y="${child.y}"/><a:chExt cx="${child.cx}" cy="${child.cy}"/>`
    : "";
  return `<a:xfrm><a:off x="${box.x}" y="${box.y}"/><a:ext cx="${box.cx}" cy="${box.cy}"/>${ch}</a:xfrm>`;
}

/** A `wps:wsp` group member: a filled rectangle at its own child coordinates. */
function member(id: number, box: Box, descr = `member ${id}`): string {
  return (
    `<wps:wsp><wps:cNvPr id="${id}" name="Rectangle ${id}" descr="${descr}"/><wps:cNvSpPr/>` +
    `<wps:spPr>${xfrm(box)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>` +
    '<a:solidFill><a:srgbClr val="4472C4"/></a:solidFill></wps:spPr><wps:bodyPr/></wps:wsp>'
  );
}

/** A `pic:pic` group member pointing at the package's one image part. */
function picMember(id: number, box: Box, descr = `picture ${id}`): string {
  return (
    `<pic:pic><pic:nvPicPr><pic:cNvPr id="${id}" name="Picture ${id}" descr="${descr}"/>` +
    "<pic:cNvPicPr/></pic:nvPicPr>" +
    '<pic:blipFill><a:blip r:embed="rId5"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>' +
    `<pic:spPr>${xfrm(box)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>`
  );
}

/** A floating `wpg:wgp` group drawing: the group box on the page, its child
 * coordinate space, and its member markup. */
function groupDrawing(options: {
  id: number;
  page: Box;
  childSpace?: Box;
  members: string;
  wrapXml?: string;
}): string {
  const { id, page, childSpace, members, wrapXml = "<wp:wrapNone/>" } = options;
  const bounds: Box = { x: 0, y: 0, cx: page.cx, cy: page.cy };
  return (
    '<w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0"' +
    ` relativeHeight="${251658240 + id}" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">` +
    '<wp:simplePos x="0" y="0"/>' +
    `<wp:positionH relativeFrom="column"><wp:posOffset>${page.x}</wp:posOffset></wp:positionH>` +
    `<wp:positionV relativeFrom="paragraph"><wp:posOffset>${page.y}</wp:posOffset></wp:positionV>` +
    `<wp:extent cx="${page.cx}" cy="${page.cy}"/>${wrapXml}` +
    `<wp:docPr id="${id}" name="Group ${id}"/><wp:cNvGraphicFramePr/>` +
    '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup">' +
    "<wpg:wgp><wpg:cNvGrpSpPr/>" +
    `<wpg:grpSpPr>${xfrm(bounds, childSpace ?? bounds)}</wpg:grpSpPr>` +
    members +
    "</wpg:wgp></a:graphicData></a:graphic></wp:anchor></w:drawing>"
  );
}

/** A nested `wpg:grpSp`: the same content model as `wpg:wgp`, one level in. */
function nestedGroup(box: Box, childSpace: Box, members: string): string {
  return `<wpg:grpSp><wpg:cNvGrpSpPr/><wpg:grpSpPr>${xfrm(box, childSpace)}</wpg:grpSpPr>${members}</wpg:grpSp>`;
}

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

function modelOf(parsed: ParsedDocxDocument): DocumentModel {
  return {
    id: "test-doc",
    title: "Groups",
    sections: parsed.sections,
    defaultStyles: parsed.defaultStyles,
    format: "docx",
    media: parsed.media,
    opaque: parsed.opaque,
    sourceXml: parsed.sourceXml,
  };
}

describe("parseDocx — a wpg:wgp group flattens into its members", () => {
  // Group box 4x2cm at (3,1). Child space 1:1, so a member's own coordinates
  // are its offset inside the group.
  const simple = groupDrawing({
    id: 9,
    page: { x: 3 * CM, y: 1 * CM, cx: 4 * CM, cy: 2 * CM },
    members: member(10, { x: 0, y: 0, cx: CM, cy: CM }) + member(11, { x: 2 * CM, y: CM / 2, cx: CM, cy: CM }),
  });

  it("yields one block per member, positioned in page space", async () => {
    const blocks = await blocksOf([simple]);
    expect(blocks.map((block) => block.type)).toEqual(["shape", "shape", "paragraph"]);
    const [first, second] = blocks as [Shape, Shape];
    // Member offsets are the group's page offset plus their own position
    // inside it — 0,0 and 2,0.5 cm respectively.
    expect(first.offsetXCm).toBeCloseTo(3, 6);
    expect(first.offsetYCm).toBeCloseTo(1, 6);
    expect(second.offsetXCm).toBeCloseTo(5, 6);
    expect(second.offsetYCm).toBeCloseTo(1.5, 6);
    expect(first.widthCm).toBeCloseTo(1, 6);
    expect(second.heightCm).toBeCloseTo(1, 6);
  });

  it("gives every member the same group ordinal and the group's wrap and z-order", async () => {
    const blocks = await blocksOf([simple]);
    const [first, second] = blocks as [Shape, Shape];
    expect(first.groupOrdinal).toBe(1);
    expect(second.groupOrdinal).toBe(first.groupOrdinal);
    expect(first.wrap).toBe("infront");
    expect(second.wrap).toBe("infront");
    expect(first.zOrder).toBe(9);
    expect(second.zOrder).toBe(9);
  });

  it("reads each member's own appearance and alt text", async () => {
    const shape = (await blocksOf([simple]))[0] as Shape;
    expect(shape.geometry).toBe("rect");
    expect(shape.fill).toEqual({ kind: "solid", color: "#4472c4" });
    // A member's alt text is on its own cNvPr, not the container's docPr.
    expect(shape.altText).toBe("member 10");
  });

  it("flattens a picture member beside a shape one — Word's most common group", async () => {
    // This took the whole group down with it until Image gained real
    // placement: an image had no position anywhere in the model, so a
    // flattened one would have painted at the group's origin.
    const mixed = groupDrawing({
      id: 42,
      page: { x: 3 * CM, y: 1 * CM, cx: 4 * CM, cy: 2 * CM },
      members:
        member(43, { x: 0, y: 0, cx: CM, cy: CM }) + picMember(44, { x: 2 * CM, y: CM / 2, cx: CM, cy: CM }),
    });
    const parsed = await parse([mixed]);
    const blocks = parsed.sections[0].blocks;
    expect(blocks.map((block) => block.type)).toEqual(["shape", "image", "paragraph"]);
    const shape = blocks[0] as Shape;
    const image = blocks[1] as Image;
    expect(image.groupOrdinal).toBe(shape.groupOrdinal);
    expect(image.offsetXCm).toBeCloseTo(5, 6);
    expect(image.offsetYCm).toBeCloseTo(1.5, 6);
    expect(image.widthCm).toBeCloseTo(1, 6);
    expect(image.altText).toBe("picture 44");
    expect(parsed.media.map((part) => part.mediaId)).toContain("word/media/image1.png");
  });

  it("reads the group whether or not Word wrapped it in mc:AlternateContent", async () => {
    const wrapped =
      '<mc:AlternateContent><mc:Choice Requires="wpg">' +
      simple +
      '</mc:Choice><mc:Fallback><w:pict><v:group id="_x0000_s1040"/></w:pict></mc:Fallback></mc:AlternateContent>';
    expect(await blocksOf([wrapped])).toEqual(await blocksOf([simple]));
  });
});

describe("parseDocx — group coordinate mapping", () => {
  it("scales member positions and sizes through chOff/chExt", async () => {
    // A 4x2cm group whose children are authored in an 8x4cm space: every
    // child coordinate and extent halves.
    const scaled = groupDrawing({
      id: 12,
      page: { x: 3 * CM, y: 1 * CM, cx: 4 * CM, cy: 2 * CM },
      childSpace: { x: 0, y: 0, cx: 8 * CM, cy: 4 * CM },
      members: member(13, { x: 4 * CM, y: 2 * CM, cx: 2 * CM, cy: 2 * CM }),
    });
    const shape = (await blocksOf([scaled]))[0] as Shape;
    expect(shape.offsetXCm).toBeCloseTo(3 + 2, 6);
    expect(shape.offsetYCm).toBeCloseTo(1 + 1, 6);
    expect(shape.widthCm).toBeCloseTo(1, 6);
    expect(shape.heightCm).toBeCloseTo(1, 6);
  });

  it("honours a non-zero chOff", async () => {
    const shifted = groupDrawing({
      id: 14,
      page: { x: 2 * CM, y: 0, cx: 4 * CM, cy: 4 * CM },
      childSpace: { x: CM, y: CM, cx: 4 * CM, cy: 4 * CM },
      members: member(15, { x: 2 * CM, y: 3 * CM, cx: CM, cy: CM }),
    });
    const shape = (await blocksOf([shifted]))[0] as Shape;
    // Child (2,3) sits 1,2 cm into a child space starting at (1,1).
    expect(shape.offsetXCm).toBeCloseTo(2 + 1, 6);
    expect(shape.offsetYCm).toBeCloseTo(0 + 2, 6);
  });

  it("composes a nested group's mapping with its parent's", async () => {
    // Outer 4x2cm at (3,1), 1:1. Nested group occupies (2,0)-(4,2) of that
    // space but authors its own children in a 4x4cm space, halving them.
    const nested = groupDrawing({
      id: 16,
      page: { x: 3 * CM, y: 1 * CM, cx: 4 * CM, cy: 2 * CM },
      members:
        member(17, { x: 0, y: 0, cx: CM, cy: CM }) +
        nestedGroup(
          { x: 2 * CM, y: 0, cx: 2 * CM, cy: 2 * CM },
          { x: 0, y: 0, cx: 4 * CM, cy: 4 * CM },
          member(18, { x: 2 * CM, y: 0, cx: 2 * CM, cy: 2 * CM }),
        ),
    });
    const blocks = await blocksOf([nested]);
    expect(blocks.map((block) => block.type)).toEqual(["shape", "shape", "paragraph"]);
    const inner = blocks[1] as Shape;
    expect(inner.offsetXCm).toBeCloseTo(3 + 2 + 1, 6);
    expect(inner.offsetYCm).toBeCloseTo(1, 6);
    expect(inner.widthCm).toBeCloseTo(1, 6);
  });

  it("flattens a nested group into the SAME ordinal — the flat model has no nesting", async () => {
    const nested = groupDrawing({
      id: 19,
      page: { x: 0, y: 0, cx: 4 * CM, cy: 2 * CM },
      members:
        member(20, { x: 0, y: 0, cx: CM, cy: CM }) +
        nestedGroup(
          { x: 2 * CM, y: 0, cx: 2 * CM, cy: 2 * CM },
          { x: 0, y: 0, cx: 2 * CM, cy: 2 * CM },
          member(21, { x: 0, y: 0, cx: CM, cy: CM }),
        ),
    });
    const [outer, inner] = (await blocksOf([nested])) as [Shape, Shape];
    expect(inner.groupOrdinal).toBe(outer.groupOrdinal);
  });

  it("places members at their own coordinates when the group has no usable chExt", async () => {
    const degenerate = groupDrawing({
      id: 22,
      page: { x: CM, y: CM, cx: 4 * CM, cy: 2 * CM },
      childSpace: { x: 0, y: 0, cx: 0, cy: 0 },
      members: member(23, { x: 2 * CM, y: 0, cx: CM, cy: CM }),
    });
    const shape = (await blocksOf([degenerate]))[0] as Shape;
    expect(shape.offsetXCm).toBeCloseTo(1 + 2, 6);
    expect(shape.widthCm).toBeCloseTo(1, 6);
  });

  it("floats an inline group's members rather than putting each in the text flow", async () => {
    // The model's grouping is floating-only (shapes.ts's groupShapes converts
    // an inline member the same way).
    const inline =
      '<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">' +
      `<wp:extent cx="${4 * CM}" cy="${2 * CM}"/>` +
      '<wp:docPr id="24" name="Group 24"/><wp:cNvGraphicFramePr/>' +
      '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup">' +
      "<wpg:wgp><wpg:cNvGrpSpPr/>" +
      `<wpg:grpSpPr>${xfrm({ x: 0, y: 0, cx: 4 * CM, cy: 2 * CM }, { x: 0, y: 0, cx: 4 * CM, cy: 2 * CM })}</wpg:grpSpPr>` +
      member(25, { x: 0, y: 0, cx: CM, cy: CM }) +
      "</wpg:wgp></a:graphicData></a:graphic></wp:inline></w:drawing>";
    const shape = (await blocksOf([inline]))[0] as Shape;
    expect(shape.wrap).toBe("square");
    expect(shape.groupOrdinal).toBe(1);
  });
});

describe("parseDocx — group ordinals", () => {
  it("gives two groups in one document different ordinals", async () => {
    const first = groupDrawing({
      id: 26,
      page: { x: 0, y: 0, cx: 2 * CM, cy: 2 * CM },
      members: member(27, { x: 0, y: 0, cx: CM, cy: CM }),
    });
    const second = groupDrawing({
      id: 28,
      page: { x: 0, y: 0, cx: 2 * CM, cy: 2 * CM },
      members: member(29, { x: 0, y: 0, cx: CM, cy: CM }),
    });
    const blocks = await blocksOf([first, second]);
    const ordinals = blocks
      .filter((block) => block.type === "shape")
      .map((block) => (block as Shape).groupOrdinal);
    expect(ordinals).toEqual([1, 2]);
  });

  it("mints above every de:group the document already carries", async () => {
    // Both group flavours coexist once a user groups shapes in Word inside a
    // file this editor saved: minting from 1 would merge them.
    const existing =
      '<w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="251658240"' +
      ' behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">' +
      '<wp:simplePos x="0" y="0"/>' +
      '<wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>' +
      '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>' +
      `<wp:extent cx="${CM}" cy="${CM}"/><wp:wrapNone/>` +
      '<wp:docPr id="30" name="Rectangle 30" de:group="4"/><wp:cNvGraphicFramePr/>' +
      '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
      `<wps:wsp><wps:cNvSpPr/><wps:spPr>${xfrm({ x: 0, y: 0, cx: CM, cy: CM })}` +
      '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr><wps:bodyPr/></wps:wsp>' +
      "</a:graphicData></a:graphic></wp:anchor></w:drawing>";
    const imported = groupDrawing({
      id: 31,
      page: { x: 0, y: 0, cx: 2 * CM, cy: 2 * CM },
      members: member(32, { x: 0, y: 0, cx: CM, cy: CM }),
    });
    const blocks = await blocksOf([existing, imported]);
    const shapes = blocks.filter((block) => block.type === "shape") as Shape[];
    expect(shapes.map((shape) => shape.groupOrdinal)).toEqual([4, 5]);
  });
});

describe("parseDocx — a group nothing can fully read stays preserved", () => {
  // All-or-nothing: reading the members we understand and dropping the rest
  // would discard document content (SPEC-008).
  it("preserves the whole run when one member is markup no reader models", async () => {
    const withCanvasMember = groupDrawing({
      id: 33,
      page: { x: 0, y: 0, cx: 4 * CM, cy: 2 * CM },
      members:
        member(34, { x: 0, y: 0, cx: CM, cy: CM }) +
        '<wpg:graphicFrame><wpg:cNvPr id="35" name="Table 35"/></wpg:graphicFrame>',
    });
    const blocks = await blocksOf([withCanvasMember]);
    expect(blocks.map((block) => block.type)).toEqual(["paragraph", "paragraph"]);
    const preserved = blocks[0];
    const xml = preserved.type === "paragraph" ? (preserved.opaqueInlines ?? []).map((i) => i.xml).join("") : "";
    expect(xml).toContain("wgp");
    expect(xml).toContain("graphicFrame");
  });

  it("preserves a group whose members carry no a:xfrm to place them by", async () => {
    const noXfrm = groupDrawing({
      id: 36,
      page: { x: 0, y: 0, cx: 4 * CM, cy: 2 * CM },
      members:
        '<wps:wsp><wps:cNvPr id="37" name="Rectangle 37"/><wps:cNvSpPr/>' +
        '<wps:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr><wps:bodyPr/></wps:wsp>',
    });
    const blocks = await blocksOf([noXfrm]);
    expect(blocks.map((block) => block.type)).toEqual(["paragraph", "paragraph"]);
  });

  it("preserves an empty group rather than erasing it", async () => {
    const empty = groupDrawing({ id: 38, page: { x: 0, y: 0, cx: CM, cy: CM }, members: "" });
    const blocks = await blocksOf([empty]);
    expect(blocks.map((block) => block.type)).toEqual(["paragraph", "paragraph"]);
  });
});

describe("writeDocx — a flattened group survives save and reopen", () => {
  const simple = groupDrawing({
    id: 39,
    page: { x: 3 * CM, y: 1 * CM, cx: 4 * CM, cy: 2 * CM },
    members: member(40, { x: 0, y: 0, cx: CM, cy: CM }) + member(41, { x: 2 * CM, y: CM / 2, cx: CM, cy: CM }),
  });

  it("keeps both members grouped, positioned and sized through a round trip", async () => {
    const parsed = await parse([simple]);
    const { bytes } = await writeDocx(modelOf(parsed));
    const reread = await parseDocx(bytes);
    const blocks = reread.sections[0].blocks;
    expect(blocks.map((block) => block.type)).toEqual(["shape", "shape", "paragraph"]);
    const [first, second] = blocks as [Shape, Shape];
    expect(first.groupOrdinal).toBe(1);
    expect(second.groupOrdinal).toBe(1);
    expect(first.offsetXCm).toBeCloseTo(3, 4);
    expect(second.offsetXCm).toBeCloseTo(5, 4);
    expect(second.offsetYCm).toBeCloseTo(1.5, 4);
    expect(second.widthCm).toBeCloseTo(1, 4);
  });

  it("writes the members as sibling drawings, not as a wpg:wgp", async () => {
    // The documented trade-off (group.ts's header): Word sees N ungrouped
    // drawings; this editor still sees one group via `de:group`.
    const parsed = await parse([simple]);
    const { bytes } = await writeDocx(modelOf(parsed));
    const documentXml = new TextDecoder().decode((await unzip(bytes)).get("word/document.xml")!);
    expect(documentXml).not.toContain("wgp");
    expect(documentXml.match(/de:group="1"/g)).toHaveLength(2);
  });
});
