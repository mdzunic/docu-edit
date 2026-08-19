// SPEC-057 native `wpg:wgp` group reading (.docx). These parse constructed
// native group markup — the shape catalogue member (`wps:wsp`), a picture
// member (`pic:pic`), a nested group (`wpg:grpSp`), container/member rotation
// (`@rot`) and a non-identity child coordinate space (`a:chOff`/`a:chExt`) —
// straight through `parseGroupDrawing`, then round-trip a container back out
// through the writer (`groupDrawingXml`) and re-read it. AC-53..59.
import { describe, it, expect } from "vitest";
import { parseXml } from "../core/xml";
import { parseGroupDrawing } from "./groupReader";
import { groupDrawingXml, type DocPrIdAllocator } from "./group";
import { isDrawingGroup, isImage, isShape, type DrawingGroup } from "../documentModel";
import type { MediaPart } from "../core/media";
import type { MediaWritePlan } from "./media";

const WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const IMAGE_RELATIONSHIP = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
const MAIN_PART_PATH = "word/document.xml";

function buildPngBytes(width: number, height: number): Uint8Array {
  const bytes = new Uint8Array(24);
  bytes.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a], 0);
  bytes.set([0x49, 0x48, 0x44, 0x52], 12);
  bytes[16] = (width >>> 24) & 0xff;
  bytes[17] = (width >>> 16) & 0xff;
  bytes[18] = (width >>> 8) & 0xff;
  bytes[19] = width & 0xff;
  bytes[20] = (height >>> 24) & 0xff;
  bytes[21] = (height >>> 16) & 0xff;
  bytes[22] = (height >>> 8) & 0xff;
  bytes[23] = height & 0xff;
  return bytes;
}
const PNG_BYTES = buildPngBytes(400, 200);
const rels = new Map([["rId1", { id: "rId1", type: IMAGE_RELATIONSHIP, target: "media/image1.png" }]]);
const parts = new Map([["word/media/image1.png", PNG_BYTES]]);

function drawingElement(xml: string): Element {
  return parseXml(`<root xmlns:w="${WORD_NS}">${xml}</root>`).documentElement.children[0];
}

/** A `wps:wsp` catalogue-shape member at `off`/`ext` EMU. */
function wspMember(prst: string, x: number, y: number, cx: number, cy: number, rot?: number): string {
  const rotAttr = rot ? ` rot="${rot}"` : "";
  return (
    '<wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
    `<wps:cNvPr id="0" name="${prst}"/><wps:cNvSpPr/>` +
    "<wps:spPr>" +
    `<a:xfrm${rotAttr}><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>` +
    `<a:prstGeom prst="${prst}"><a:avLst/></a:prstGeom>` +
    "</wps:spPr>" +
    "</wps:wsp>"
  );
}

/** A `pic:pic` picture member at `off`/`ext` EMU referencing rId1. */
function picMember(x: number, y: number, cx: number, cy: number, descr?: string): string {
  const descrAttr = descr ? ` descr="${descr}"` : "";
  return (
    '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
    `<pic:nvPicPr><pic:cNvPr id="0" name="Picture"${descrAttr}/><pic:cNvPicPr/></pic:nvPicPr>` +
    '<pic:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>' +
    `<pic:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>` +
    "</pic:pic>"
  );
}

/** A floating `w:drawing` holding a `wpg:wgp` group with the given xfrm and
 * member markup. */
function groupDrawing(opts: {
  ext: [number, number];
  chOff?: [number, number];
  chExt?: [number, number];
  rot?: number;
  members: string;
  descr?: string;
  posX?: number;
  posY?: number;
}): string {
  const [cx, cy] = opts.ext;
  const [chOffX, chOffY] = opts.chOff ?? [0, 0];
  const [chExtCx, chExtCy] = opts.chExt ?? [cx, cy];
  const rotAttr = opts.rot ? ` rot="${opts.rot}"` : "";
  const descrAttr = opts.descr ? ` descr="${opts.descr}"` : "";
  return (
    '<w:drawing xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">' +
    '<wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="251658240" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">' +
    '<wp:simplePos x="0" y="0"/>' +
    `<wp:positionH relativeFrom="column"><wp:posOffset>${opts.posX ?? 0}</wp:posOffset></wp:positionH>` +
    `<wp:positionV relativeFrom="paragraph"><wp:posOffset>${opts.posY ?? 0}</wp:posOffset></wp:positionV>` +
    `<wp:extent cx="${cx}" cy="${cy}"/>` +
    '<wp:effectExtent l="0" t="0" r="0" b="0"/>' +
    '<wp:wrapSquare wrapText="bothSides"/>' +
    `<wp:docPr id="1" name="Group 1"${descrAttr}/>` +
    "<wp:cNvGraphicFramePr/>" +
    '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
    '<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup">' +
    '<wpg:wgp xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup">' +
    "<wpg:cNvGrpSpPr/>" +
    "<wpg:grpSpPr>" +
    `<a:xfrm${rotAttr}><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/><a:chOff x="${chOffX}" y="${chOffY}"/><a:chExt cx="${chExtCx}" cy="${chExtCy}"/></a:xfrm>` +
    "</wpg:grpSpPr>" +
    opts.members +
    "</wpg:wgp>" +
    "</a:graphicData></a:graphic></wp:anchor></w:drawing>"
  );
}

const CTX = {
  parseParagraph: () => ({ type: "paragraph" as const, runs: [] }),
  mainRels: rels,
  mainPartPath: MAIN_PART_PATH,
  parts,
};

describe("SPEC-057 native wpg:wgp reading (.docx)", () => {
  it("reads a two-shape group as one container (AC-53)", () => {
    const el = drawingElement(groupDrawing({ ext: [2160000, 720000], members: wspMember("rect", 0, 0, 720000, 720000) + wspMember("ellipse", 1440000, 0, 720000, 720000) }));
    const group = parseGroupDrawing(el, CTX);
    expect(group).not.toBeNull();
    expect(group!.members).toHaveLength(2);
    expect(group!.members.every(isShape)).toBe(true);
    expect((group!.members[1] as { offsetXCm?: number }).offsetXCm).toBeCloseTo(4, 5);
  });

  it("reads a picture member as an Image, collecting its media (AC-54)", () => {
    const collected: MediaPart[] = [];
    const el = drawingElement(groupDrawing({ ext: [2520000, 1080000], members: wspMember("rect", 0, 0, 1080000, 1080000) + picMember(1440000, 0, 1080000, 1080000, "photo") }));
    const group = parseGroupDrawing(el, { ...CTX, collectMedia: (p) => collected.push(p) });
    expect(group).not.toBeNull();
    const image = group!.members.find(isImage);
    expect(image).toBeDefined();
    expect(image!.mediaId).toBe("word/media/image1.png");
    expect(image!.altText).toBe("photo");
    expect(image!.offsetXCm).toBeCloseTo(4, 5);
    expect(collected).toHaveLength(1);
    expect(collected[0].mediaId).toBe("word/media/image1.png");
  });

  it("seals (returns null) a picture member when no media context is supplied", () => {
    const el = drawingElement(groupDrawing({ ext: [2520000, 1080000], members: wspMember("rect", 0, 0, 1080000, 1080000) + picMember(1440000, 0, 1080000, 1080000) }));
    expect(parseGroupDrawing(el, { parseParagraph: CTX.parseParagraph })).toBeNull();
  });

  it("reads a nested wpg:grpSp as a nested container (AC-56)", () => {
    const nested =
      '<wpg:grpSp xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
      '<wpg:cNvPr id="0" name="Group"/><wpg:cNvGrpSpPr/>' +
      "<wpg:grpSpPr><a:xfrm><a:off x=\"1440000\" y=\"0\"/><a:ext cx=\"1440000\" cy=\"720000\"/><a:chOff x=\"0\" y=\"0\"/><a:chExt cx=\"1440000\" cy=\"720000\"/></a:xfrm></wpg:grpSpPr>" +
      wspMember("rect", 0, 0, 720000, 720000) +
      wspMember("ellipse", 720000, 0, 720000, 720000) +
      "</wpg:grpSp>";
    const el = drawingElement(groupDrawing({ ext: [2880000, 720000], members: wspMember("triangle", 0, 0, 720000, 720000) + nested }));
    const group = parseGroupDrawing(el, CTX);
    expect(group).not.toBeNull();
    expect(group!.members).toHaveLength(2);
    const inner = group!.members.find(isDrawingGroup) as DrawingGroup;
    expect(inner).toBeDefined();
    expect(inner.members).toHaveLength(2);
    expect(inner.offsetXCm).toBeCloseTo(4, 5);
  });

  it("reads container rotation (@rot) and a member's own rotation (AC-57)", () => {
    const el = drawingElement(groupDrawing({ ext: [2160000, 720000], rot: 2700000, members: wspMember("rect", 0, 0, 720000, 720000, 900000) + wspMember("ellipse", 1440000, 0, 720000, 720000) }));
    const group = parseGroupDrawing(el, CTX);
    expect(group!.rotationDeg).toBe(45);
    expect((group!.members[0] as { rotationDeg?: number }).rotationDeg).toBe(15);
  });

  it("applies a non-identity child coordinate space (AC-58)", () => {
    // chExt = 2× ext → scale 0.5: a child at off 2880000 / ext 1440000 lands at
    // 1440000 EMU (4cm) offset, 720000 EMU (2cm) size.
    const el = drawingElement(groupDrawing({
      ext: [2160000, 720000],
      chExt: [4320000, 1440000],
      members: wspMember("rect", 0, 0, 1440000, 1440000) + wspMember("ellipse", 2880000, 0, 1440000, 1440000),
    }));
    const group = parseGroupDrawing(el, CTX);
    expect((group!.members[0] as { widthCm?: number }).widthCm).toBeCloseTo(2, 5);
    expect((group!.members[1] as { offsetXCm?: number }).offsetXCm).toBeCloseTo(4, 5);
  });
});

describe("SPEC-057 native wpg:wgp writing round-trips (.docx) (AC-59)", () => {
  const mediaPlan: MediaWritePlan = { relIdFor: (id) => (id === "word/media/image1.png" ? "rId1" : undefined) } as MediaWritePlan;
  const ids: DocPrIdAllocator = (() => {
    let n = 100;
    return { next: () => n++ };
  })();
  const writeCtx = { paragraphToXml: () => "" } as never;

  it("writes one wp:anchor holding wpg:wgp with the members as children", () => {
    const group: DrawingGroup = {
      type: "drawingGroup",
      widthCm: 6,
      heightCm: 3,
      wrap: "square",
      offsetXCm: 1,
      offsetYCm: 2,
      relativeToH: "column",
      relativeToV: "paragraph",
      members: [
        { type: "shape", geometry: "rect", widthCm: 2, heightCm: 2, wrap: "square", offsetXCm: 0, offsetYCm: 0 },
        { type: "image", mediaId: "word/media/image1.png", naturalWidthPx: 400, naturalHeightPx: 200, widthCm: 3, heightCm: 3, wrap: "square", offsetXCm: 3, offsetYCm: 0 },
      ],
    };
    const xml = groupDrawingXml(group, writeCtx, mediaPlan, ids);
    expect(xml).toContain("<wp:anchor");
    expect(xml).toContain("wordprocessingGroup");
    expect(xml).toContain("<wpg:wgp");
    expect(xml).toContain("<pic:pic");

    // Re-read the written drawing and confirm the members survive.
    const reparsed = parseGroupDrawing(drawingElement(xml), { ...CTX, collectMedia: () => {} });
    expect(reparsed).not.toBeNull();
    expect(reparsed!.members).toHaveLength(2);
    expect(reparsed!.members.some(isImage)).toBe(true);
    expect(reparsed!.members.some(isShape)).toBe(true);
  });
});
