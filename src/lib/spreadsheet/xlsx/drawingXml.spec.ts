// SPEC-131's xlsx drawing layer (xlsx/drawingXml.ts): every reading rule of
// "Reading xlsx drawings" (each anchor and element kind, cNvPr fields,
// rotation, crop, linked and EMF pictures, groups, frames, ink), the MCE
// labels and faces, golden equality of the canonical serialiser against
// xlsx/drawingGolden.ts, anchor patching that leaves the content element's
// bytes alone (and an untouched part byte-identical), and the writer's
// package bookkeeping: a first drawing's part, rels, <drawing>, relationship
// and Override; the last drawing's removal; media that survive or go by
// reachability; a clean workbook verbatim; drawings from scratch.
import { describe, expect, it } from "vitest";
import { ensureMediaEncoded } from "../../core/media";
import { bytesToBase64 } from "../../core/packageParts";
import { unzip } from "../../core/zip";
import { zip } from "../../core/zipWriter";
import { applyDrawingResult, deleteDrawing, insertImage, insertShape, setDrawingAltText, setDrawingRect } from "../drawings";
import { packageIntegrity } from "../fixtures/packageIntegrity";
import { buildGridMetrics } from "../geometry";
import type { Sheet, SheetDrawing, SpreadsheetModel } from "../model";
import { readSpreadsheetBytes } from "../open";
import { createBlankWorkbook } from "../sheets";
import { applyStructuralEdit } from "../structure";
import * as G from "./drawingGolden";
import { canonicalAnchorXml, spliceWorksheetDrawing, UNDERSTOOD_SHEET_DRAWING_NAMESPACES, writeDrawingPart } from "./drawingXml";
import { substitutedXlsxParts, writeXlsx } from "./writer";

const DECL = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`;
const MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const MC = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const PKG = "http://schemas.openxmlformats.org/package/2006/relationships";
const PNG = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 1, 2, 3, 4]);
const EMF = new Uint8Array([1, 0, 0, 0, 0x58, 0, 0, 0]);
const decoder = new TextDecoder();
const encoder = new TextEncoder();

// --- building small workbooks -------------------------------------------------------------

/** `<xdr:from>`/`<xdr:to>` at [col, colOff EMU, row, rowOff EMU]. */
const corner = (tag: "from" | "to", [col, colOff, row, rowOff]: readonly [number, number, number, number]): string =>
  `<xdr:${tag}><xdr:col>${col}</xdr:col><xdr:colOff>${colOff}</xdr:colOff><xdr:row>${row}</xdr:row><xdr:rowOff>${rowOff}</xdr:rowOff></xdr:${tag}>`;
const twoCell = (from: readonly [number, number, number, number], to: readonly [number, number, number, number], content: string, editAs?: string): string =>
  `<xdr:twoCellAnchor${editAs ? ` editAs="${editAs}"` : ""}>${corner("from", from)}${corner("to", to)}${content}<xdr:clientData/></xdr:twoCellAnchor>`;
const sp = (id: number, name: string, prst: string, extra = "", attrs = ""): string =>
  `<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="${id}" name="${name}"${attrs}/><xdr:cNvSpPr/></xdr:nvSpPr>` +
  `<xdr:spPr><a:xfrm><a:off x="609600" y="190500"/><a:ext cx="1219200" cy="381000"/></a:xfrm><a:prstGeom prst="${prst}"><a:avLst/></a:prstGeom>${extra}</xdr:spPr></xdr:sp>`;
const pic = (id: number, name: string, blip: string, attrs = "", srcRect = ""): string =>
  `<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="${id}" name="${name}"${attrs}/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip ${blip}/>${srcRect}<a:stretch><a:fillRect/></a:stretch></xdr:blipFill>` +
  `<xdr:spPr><a:xfrm><a:off x="609600" y="190500"/><a:ext cx="609600" cy="381000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic>`;
const wsDr = (children: string): string => `${DECL}<xdr:wsDr xmlns:xdr="${XDR}" xmlns:a="${A}" xmlns:r="${R}">${children}</xdr:wsDr>`;
const rel = (id: string, kind: string, target: string): string => `<Relationship Id="${id}" Type="${R}/${kind}" Target="${target}"/>`;
const rels = (body: string): string => `${DECL}<Relationships xmlns="${PKG}">${body}</Relationships>`;

interface BookOptions {
  /** Sheet1's drawing part children; absent = Sheet1 has no drawing. */
  drawing?: string;
  drawingRels?: string;
  /** Sheet2's drawing (sharing media). */
  drawing2?: string;
  drawing2Rels?: string;
  media?: Record<string, Uint8Array>;
  extraParts?: Record<string, string>;
  defaults?: string;
  overrides?: string;
}

/** A two-sheet Excel-shaped workbook; Sheet1 holds cells A1:A3. */
async function bookBytes(options: BookOptions = {}): Promise<Uint8Array> {
  const sheet = (hasDrawing: boolean, selected = false): string =>
    `${DECL}<worksheet xmlns="${MAIN}" xmlns:r="${R}"><dimension ref="A1:A3"/><sheetViews><sheetView${selected ? ` tabSelected="1"` : ""} workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/>` +
    `<sheetData><row r="1"><c r="A1"><v>1</v></c></row><row r="2"><c r="A2"><v>2</v></c></row><row r="3"><c r="A3"><v>3</v></c></row></sheetData>` +
    `<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>${hasDrawing ? `<drawing r:id="rId1"/>` : ""}</worksheet>`;
  const parts: Record<string, string | Uint8Array> = {
    "[Content_Types].xml":
      `${DECL}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
      `<Default Extension="xml" ContentType="application/xml"/>${options.defaults ?? ""}` +
      `<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>` +
      `<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>` +
      `<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>` +
      (options.drawing === undefined ? "" : `<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>`) +
      (options.drawing2 === undefined ? "" : `<Override PartName="/xl/drawings/drawing2.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>`) +
      `${options.overrides ?? ""}</Types>`,
    "_rels/.rels": rels(`<Relationship Id="rId1" Type="${R}/officeDocument" Target="xl/workbook.xml"/>`),
    "xl/workbook.xml": `${DECL}<workbook xmlns="${MAIN}" xmlns:r="${R}"><bookViews><workbookView activeTab="0"/></bookViews><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/><sheet name="Sheet2" sheetId="2" r:id="rId2"/></sheets></workbook>`,
    "xl/_rels/workbook.xml.rels": rels(rel("rId1", "worksheet", "worksheets/sheet1.xml") + rel("rId2", "worksheet", "worksheets/sheet2.xml")),
    "xl/worksheets/sheet1.xml": sheet(options.drawing !== undefined, true),
    "xl/worksheets/sheet2.xml": sheet(options.drawing2 !== undefined),
  };
  if (options.drawing !== undefined) {
    parts["xl/worksheets/_rels/sheet1.xml.rels"] = rels(rel("rId1", "drawing", "../drawings/drawing1.xml"));
    parts["xl/drawings/drawing1.xml"] = wsDr(options.drawing);
    parts["xl/drawings/_rels/drawing1.xml.rels"] = rels(options.drawingRels ?? "");
  }
  if (options.drawing2 !== undefined) {
    parts["xl/worksheets/_rels/sheet2.xml.rels"] = rels(rel("rId1", "drawing", "../drawings/drawing2.xml"));
    parts["xl/drawings/drawing2.xml"] = wsDr(options.drawing2);
    parts["xl/drawings/_rels/drawing2.xml.rels"] = rels(options.drawing2Rels ?? "");
  }
  Object.assign(parts, options.media ?? {}, options.extraParts ?? {});
  return zip(Object.entries(parts).map(([name, data]) => ({ name, data: typeof data === "string" ? encoder.encode(data) : data })));
}

async function open(options: BookOptions = {}): Promise<{ bytes: Uint8Array; model: SpreadsheetModel }> {
  const bytes = await bookBytes(options);
  return { bytes, model: await readSpreadsheetBytes(bytes, "xlsx", "book") };
}

/** Sheet1's drawings, sourceXml left out. */
const drawingsOf = (model: SpreadsheetModel, index = 0): Omit<SheetDrawing, "sourceXml">[] =>
  (model.sheets[index]!.drawings ?? []).map(({ sourceXml: _source, ...rest }) => rest);

async function save(model: SpreadsheetModel): Promise<Map<string, Uint8Array>> {
  const bytes = await writeXlsx(model, { modified: new Date(Date.UTC(2026, 0, 1)) });
  expect(await packageIntegrity(bytes)).toEqual([]);
  return unzip(bytes);
}

const text = (parts: Map<string, Uint8Array>, name: string): string => {
  expect(parts.has(name), name).toBe(true);
  return decoder.decode(parts.get(name));
};

const PNG_DEFAULT = `<Default Extension="png" ContentType="image/png"/>`;
const PICTURE = twoCell([1, 0, 2, 0], [2, 304800, 5, 38100], pic(2, "Picture 1", `r:embed="rId1"`, ` descr="Company logo"`), "oneCell");

// --- reading -------------------------------------------------------------------------------

/** A group member shape over a child-space box. */
const member = (id: number, name: string, prst: string, [x, y, cx, cy]: readonly number[], attrs = ""): string =>
  `<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="${id}" name="${name}"${attrs}/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>` +
  `<a:prstGeom prst="${prst}"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill></xdr:spPr></xdr:sp>`;
/** A graphic frame holding `inner` under `uri`. */
const frame = (id: number, uri: string, inner: string): string =>
  `<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${id}" name="Frame ${id}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>` +
  `<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="${uri}">${inner}</a:graphicData></a:graphic></xdr:graphicFrame>`;
const chart = (rId: string): string => `<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="${rId}"/>`;

describe("reading xlsx drawings", () => {
  it("reads each anchor kind: twoCell (editAs always set), oneCell, absolute; move and fixed twoCell anchors gain their size truth", async () => {
    const { model } = await open({
      drawing:
        twoCell([1, 0, 1, 0], [3, 0, 4, 95250], sp(2, "Rectangle 1", "rect")) +
        twoCell([1, 0, 1, 0], [3, 0, 4, 0], sp(3, "Rectangle 2", "rect"), "oneCell") +
        twoCell([1, 0, 1, 0], [3, 0, 4, 0], sp(4, "Rectangle 3", "rect"), "absolute") +
        `<xdr:oneCellAnchor>${corner("from", [7, 0, 2, 0])}<xdr:ext cx="609600" cy="381000"/>${sp(5, "Oval 4", "ellipse")}<xdr:clientData/></xdr:oneCellAnchor>` +
        `<xdr:absoluteAnchor><xdr:pos x="6096000" y="190500"/><xdr:ext cx="914400" cy="381000"/>${sp(6, "Rectangle 5", "rect")}<xdr:clientData/></xdr:absoluteAnchor>`,
    });
    const anchors = drawingsOf(model).map((drawing) => drawing.anchor);
    expect(anchors[0]).toEqual({ kind: "twoCell", editAs: "twoCell", from: { col: 1, colOffEmu: 0, row: 1, rowOffEmu: 0 }, to: { col: 3, colOffEmu: 0, row: 4, rowOffEmu: 95250 } });
    expect(anchors[1]).toMatchObject({ editAs: "oneCell", extEmu: { cx: 2 * 609600, cy: 3 * 190500 } });
    expect(anchors[1]!.posEmu).toBeUndefined();
    expect(anchors[2]).toMatchObject({ editAs: "absolute", posEmu: { x: 609600, y: 190500 }, extEmu: { cx: 2 * 609600, cy: 3 * 190500 } });
    expect(anchors[3]).toEqual({ kind: "oneCell", from: { col: 7, colOffEmu: 0, row: 2, rowOffEmu: 0 }, extEmu: { cx: 609600, cy: 381000 } });
    expect(anchors[4]).toEqual({ kind: "absolute", posEmu: { x: 6096000, y: 190500 }, extEmu: { cx: 914400, cy: 381000 } });
    expect(drawingsOf(model).map((drawing) => drawing.id)).toEqual(["dr1", "dr2", "dr3", "dr4", "dr5"]);
    expect(model.sheets[0]!.drawingPart).toBe("xl/drawings/drawing1.xml");
    expect(model.sheets[1]!.drawingPart).toBeUndefined();
    expect(model.sheets[1]!.drawings).toBeUndefined();
  });

  it("slices each top-level child with the whitespace before it, so head + every sourceXml + tail is the part", async () => {
    const body = `\n  ${twoCell([0, 0, 0, 0], [1, 0, 1, 0], sp(2, "A", "rect"))}\n  <!-- note -->${twoCell([2, 0, 2, 0], [3, 0, 3, 0], sp(3, "B", "rect"))}\n`;
    const { model } = await open({ drawing: body });
    const sources = model.sheets[0]!.drawings!.map((drawing) => drawing.sourceXml!);
    expect(sources[0]!.startsWith("\n  <xdr:twoCellAnchor>")).toBe(true);
    expect(sources[1]!.startsWith("\n  <!-- note --><xdr:twoCellAnchor>")).toBe(true);
    expect(sources.join("") + "\n").toBe(body);
  });

  it("cNvPr name/descr/hidden and a:xfrm@rot populate the base fields", async () => {
    const rotated = sp(2, "Arrow 1", "rightArrow", "", ` descr="Look here" hidden="1"`).replace("<a:xfrm>", `<a:xfrm rot="5400000">`);
    const { model } = await open({ drawing: twoCell([0, 0, 0, 0], [1, 0, 1, 0], rotated) });
    expect(drawingsOf(model)[0]).toMatchObject({ name: "Arrow 1", altText: "Look here", hidden: true, rotationDeg: 90 });
  });

  it("pictures resolve r:embed into one shared MediaPart per path (across drawings and sheets), deferred; srcRect crops; r:link is a Linked picture; an EMF stays an image; an unresolvable embed is a foreign Picture", async () => {
    const { model } = await open({
      drawing:
        PICTURE +
        twoCell([4, 0, 2, 0], [5, 0, 4, 0], pic(3, "Picture 2", `r:embed="rId1"`, "", `<a:srcRect l="10000" t="20000"/>`), "oneCell") +
        twoCell([6, 0, 2, 0], [7, 0, 4, 0], pic(4, "Picture 3", `r:embed="rId1" r:link="rId3"`), "oneCell") +
        twoCell([8, 0, 2, 0], [9, 0, 4, 0], pic(5, "Picture 4", `r:embed="rId2"`), "oneCell") +
        twoCell([8, 0, 6, 0], [9, 0, 8, 0], pic(6, "Picture 5", `r:embed="rId9"`), "oneCell"),
      drawingRels: rel("rId1", "image", "../media/image1.png") + rel("rId2", "image", "../media/image2.emf") + `<Relationship Id="rId3" Type="${R}/image" Target="https://example.com/logo.png" TargetMode="External"/>`,
      drawing2: twoCell([0, 0, 0, 0], [1, 0, 1, 0], pic(2, "Picture 1", `r:embed="rId1"`), "oneCell"),
      drawing2Rels: rel("rId1", "image", "../media/image1.png"),
      media: { "xl/media/image1.png": PNG, "xl/media/image2.emf": EMF },
      defaults: `${PNG_DEFAULT}<Default Extension="emf" ContentType="image/x-emf"/>`,
    });
    const contents = drawingsOf(model).map((drawing) => drawing.content);
    expect(contents[0]).toEqual({ kind: "image", mediaId: "xl/media/image1.png" });
    expect(contents[1]).toEqual({ kind: "image", mediaId: "xl/media/image1.png", crop: { l: 0.1, t: 0.2, r: 0, b: 0 } });
    expect(contents[2]).toEqual({ kind: "foreign", label: "Linked picture", face: { kind: "image", mediaId: "xl/media/image1.png" } });
    expect(contents[3]).toEqual({ kind: "image", mediaId: "xl/media/image2.emf" });
    expect(contents[4]).toEqual({ kind: "foreign", label: "Picture" });
    expect(drawingsOf(model, 1)[0]!.content).toEqual({ kind: "image", mediaId: "xl/media/image1.png" });
    expect(model.media!.map((part) => [part.mediaId, part.contentType])).toEqual([
      ["xl/media/image1.png", "image/png"],
      ["xl/media/image2.emf", "image/x-emf"],
    ]);
    const png = model.media![0]!;
    expect(png.data).toBe("");
    ensureMediaEncoded(png);
    expect(png.data).toBe(bytesToBase64(PNG));
  });

  it("shapes: presets and adjustments, custGeom/unknown presets as placeholders, spPr fill and outline, style refs, gradients, preset colours, text", async () => {
    const styled =
      `<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="3" name="Rounded Rectangle 2"/><xdr:cNvSpPr/></xdr:nvSpPr>` +
      `<xdr:spPr><a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 25000"/></a:avLst></a:prstGeom></xdr:spPr>` +
      `<xdr:style><a:lnRef idx="2"><a:schemeClr val="accent1"><a:shade val="15000"/></a:schemeClr></a:lnRef><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef>` +
      `<a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></xdr:style>` +
      `<xdr:txBody><a:bodyPr anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en-US" sz="1400" b="1"/><a:t>Line one</a:t></a:r><a:br/><a:r><a:t>more</a:t></a:r></a:p>` +
      `<a:p><a:pPr algn="r"/><a:r><a:t>Two</a:t></a:r></a:p></xdr:txBody></xdr:sp>`;
    const { model } = await open({
      drawing:
        twoCell([0, 0, 0, 0], [2, 0, 3, 0], styled) +
        twoCell([0, 0, 4, 0], [2, 0, 5, 0], sp(4, "Freeform 3", "rect").replace(/<a:prstGeom[^]*?<\/a:prstGeom>/, `<a:custGeom><a:pathLst/></a:custGeom>`)) +
        twoCell([0, 0, 6, 0], [2, 0, 7, 0], sp(5, "Odd 4", "notAPreset", `<a:gradFill><a:gsLst><a:gs pos="100000"><a:srgbClr val="00FF00"/></a:gs><a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs></a:gsLst></a:gradFill>`)) +
        twoCell([0, 0, 8, 0], [2, 0, 9, 0], sp(6, "Box 5", "rect", `<a:solidFill><a:prstClr val="white"/></a:solidFill><a:ln w="25400"><a:solidFill><a:srgbClr val="C00000"/></a:solidFill><a:prstDash val="sysDot"/></a:ln>`)) +
        twoCell([0, 0, 10, 0], [2, 0, 11, 0], sp(7, "Empty 6", "rect", `<a:noFill/><a:ln><a:noFill/></a:ln>`).replace("</xdr:spPr>", `</xdr:spPr><xdr:txBody><a:bodyPr/><a:p><a:endParaRPr lang="en-US"/></a:p></xdr:txBody>`)),
    });
    const [rounded, freeform, odd, box, empty] = drawingsOf(model).map((drawing) => drawing.content);
    expect(rounded).toEqual({
      kind: "shape",
      visual: { geometry: "roundRect", adjustments: [0.25], fill: { kind: "solid", color: { theme: "accent1" } }, outline: { color: { theme: "accent1", shade: 0.15 }, widthPt: 1, dash: "solid" } },
      text: { paragraphs: [{ text: "Line one\nmore", align: "center" }, { text: "Two", align: "right" }], anchor: "middle", sizePt: 14, color: { theme: "lt1" }, bold: true },
    });
    expect(freeform).toMatchObject({ kind: "shape", visual: { geometry: "rect" } });
    expect((freeform as { visual: { unsupportedGeometryXml?: string } }).visual.unsupportedGeometryXml).toContain("custGeom");
    expect(odd).toMatchObject({ visual: { geometry: "rect", fill: { kind: "solid", color: "#ff0000" } } });
    expect((odd as { visual: { unsupportedGeometryXml?: string } }).visual.unsupportedGeometryXml).toContain("notAPreset");
    expect(box).toEqual({ kind: "shape", visual: { geometry: "rect", fill: { kind: "solid", color: "#ffffff" }, outline: { color: "#c00000", widthPt: 2, dash: "dot" } } });
    expect(empty).toEqual({ kind: "shape", visual: { geometry: "rect", fill: { kind: "none" } } });
  });

  it("a connector reads its preset (straightConnector1 → arrow); a group its members normalised to the group box, a nested group or frame foreign, a hidden member left out", async () => {
    const connector =
      `<xdr:cxnSp macro=""><xdr:nvCxnSpPr><xdr:cNvPr id="4" name="Straight Arrow Connector 3"/><xdr:cNvCxnSpPr/></xdr:nvCxnSpPr>` +
      `<xdr:spPr><a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom><a:ln w="19050"><a:solidFill><a:srgbClr val="C00000"/></a:solidFill><a:tailEnd type="triangle"/></a:ln></xdr:spPr></xdr:cxnSp>`;
    const group =
      `<xdr:grpSp><xdr:nvGrpSpPr><xdr:cNvPr id="9" name="Group 8"/><xdr:cNvGrpSpPr/></xdr:nvGrpSpPr>` +
      `<xdr:grpSpPr><a:xfrm><a:off x="1000" y="2000"/><a:ext cx="4000" cy="2000"/><a:chOff x="0" y="0"/><a:chExt cx="400" cy="200"/></a:xfrm></xdr:grpSpPr>` +
      member(7, "Rectangle 6", "rect", [0, 0, 200, 200]) +
      member(8, "Oval 7", "ellipse", [200, 100, 200, 100]) +
      member(10, "Hidden 9", "rect", [0, 0, 10, 10], ` hidden="1"`) +
      `<xdr:grpSp><xdr:nvGrpSpPr><xdr:cNvPr id="11" name="Inner"/><xdr:cNvGrpSpPr/></xdr:nvGrpSpPr><xdr:grpSpPr><a:xfrm><a:off x="0" y="100"/><a:ext cx="100" cy="100"/><a:chOff x="0" y="0"/><a:chExt cx="100" cy="100"/></a:xfrm></xdr:grpSpPr></xdr:grpSp>` +
      `</xdr:grpSp>`;
    const { model } = await open({ drawing: twoCell([3, 0, 7, 0], [5, 0, 9, 0], connector) + twoCell([1, 0, 11, 0], [4, 0, 15, 0], group) });
    const [arrow, grouped] = drawingsOf(model).map((drawing) => drawing.content);
    expect(arrow).toEqual({ kind: "shape", visual: { geometry: "arrow", outline: { color: "#c00000", widthPt: 1.5, dash: "solid" } } });
    expect(grouped).toEqual({
      kind: "group",
      members: [
        { x: 0, y: 0, w: 0.5, h: 1, content: { kind: "shape", visual: { geometry: "rect", fill: { kind: "solid", color: "#4472c4" } } } },
        { x: 0.5, y: 0.5, w: 0.5, h: 0.5, content: { kind: "shape", visual: { geometry: "ellipse", fill: { kind: "solid", color: "#4472c4" } } } },
        { x: 0, y: 0.5, w: 0.25, h: 0.5, content: { kind: "foreign", label: "Object" } },
      ],
    });
  });

  it("graphic frames: a chart (sourcePart), an unresolvable chart, SmartArt, any other object; a content part is Ink; an unknown child a hidden Object", async () => {
    const { model } = await open({
      drawing:
        twoCell([0, 0, 0, 0], [4, 0, 8, 0], frame(2, "http://schemas.openxmlformats.org/drawingml/2006/chart", chart("rId1"))) +
        twoCell([0, 0, 0, 0], [4, 0, 8, 0], frame(3, "http://schemas.openxmlformats.org/drawingml/2006/chart", chart("rId7"))) +
        twoCell([0, 0, 0, 0], [4, 0, 8, 0], frame(4, "http://schemas.openxmlformats.org/drawingml/2006/diagram", "")) +
        twoCell([0, 0, 0, 0], [4, 0, 8, 0], frame(5, "urn:example:thing", "")) +
        twoCell([0, 0, 0, 0], [4, 0, 8, 0], `<xdr:contentPart r:id="rId2"/>`) +
        `<x:extra xmlns:x="urn:example"/>`,
      drawingRels: rel("rId1", "chart", "../charts/chart1.xml") + rel("rId2", "customXml", "../ink/ink1.xml"),
      extraParts: { "xl/charts/chart1.xml": `${DECL}<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>`, "xl/ink/ink1.xml": `${DECL}<ink/>` },
    });
    const drawings = drawingsOf(model);
    // SPEC-132: a chart is its part read (readXlsxChart) — an empty chartSpace
    // is out of the catalogue, so its visual is the placeholder.
    expect(drawings[0]!.content).toMatchObject({ kind: "chart", sourcePart: "xl/charts/chart1.xml" });
    expect(drawings[0]!.content).not.toHaveProperty("uncached");
    expect(drawings.slice(1).map((drawing) => drawing.content)).toEqual([
      { kind: "foreign", label: "Chart" },
      { kind: "foreign", label: "SmartArt" },
      { kind: "foreign", label: "Object" },
      { kind: "foreign", label: "Ink" },
      { kind: "foreign", label: "Object" },
    ]);
    expect(drawings[5]).toMatchObject({ hidden: true, anchor: { kind: "absolute" } });
  });

  it("a drawing part that is not well-formed reads as no drawings and stays in opaque", async () => {
    const { model } = await open({ drawing: `<xdr:twoCellAnchor>` });
    expect(model.sheets[0]!.drawings).toBeUndefined();
    expect(model.sheets[0]!.drawingPart).toBe("xl/drawings/drawing1.xml");
  });
});

// --- MCE ---------------------------------------------------------------------------------

const A14 = "http://schemas.microsoft.com/office/drawing/2010/main";
const FORM_CONTROL =
  `<mc:AlternateContent xmlns:mc="${MC}"><mc:Choice xmlns:a14="${A14}" Requires="a14">` +
  twoCell(
    [5, 0, 11, 0],
    [7, 0, 12, 0],
    `<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="1025" name="Check Box 1" hidden="1"><a:extLst><a:ext uri="{63B3BB69-23CF-44E3-9099-C40C66FF867C}"><a14:compatExt spid="_x0000_s1025"/></a:ext></a:extLst></xdr:cNvPr><xdr:cNvSpPr/></xdr:nvSpPr>` +
      `<xdr:spPr bwMode="auto"><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:sp>`,
    "absolute",
  ) +
  `</mc:Choice><mc:Fallback/></mc:AlternateContent>`;
const SLICER_TEXT = "This shape represents a table slicer.";
const slicerBranch = (inner: string): string => twoCell([10, 0, 2, 0], [12, 381000, 12, 0], inner, "oneCell");
const SLICER =
  `<mc:AlternateContent xmlns:mc="${MC}"><mc:Choice xmlns:sle15="http://schemas.microsoft.com/office/drawing/2012/slicer" Requires="sle15">` +
  slicerBranch(
    `<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="11" name="Region"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>` +
      `<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/drawing/2010/slicer"><sle:slicer xmlns:sle="http://schemas.microsoft.com/office/drawing/2010/slicer" name="Region"/></a:graphicData></a:graphic></xdr:graphicFrame>`,
  ) +
  `</mc:Choice><mc:Fallback xmlns="">` +
  slicerBranch(
    `<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="0" name=""/><xdr:cNvSpPr><a:spLocks noTextEdit="1"/></xdr:cNvSpPr></xdr:nvSpPr>` +
      `<xdr:spPr><a:xfrm><a:off x="6096000" y="381000"/><a:ext cx="1600200" cy="1905000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>` +
      `<a:solidFill><a:prstClr val="white"/></a:solidFill><a:ln w="1"><a:solidFill><a:prstClr val="green"/></a:solidFill></a:ln></xdr:spPr>` +
      `<xdr:txBody><a:bodyPr vertOverflow="clip" horzOverflow="clip"/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" sz="1100"/><a:t>${SLICER_TEXT}</a:t></a:r></a:p></xdr:txBody></xdr:sp>`,
  ) +
  `</mc:Fallback></mc:AlternateContent>`;
const INK =
  `<mc:AlternateContent xmlns:mc="${MC}"><mc:Choice xmlns:aink="http://schemas.microsoft.com/office/drawing/2016/ink" Requires="aink">` +
  twoCell([1, 0, 17, 0], [3, 0, 20, 0], `<xdr:contentPart r:id="rId3"/>`, "oneCell") +
  `</mc:Choice><mc:Fallback>${twoCell([1, 0, 17, 0], [3, 0, 20, 0], pic(12, "Ink 11", `r:embed="rId1"`), "oneCell")}</mc:Fallback></mc:AlternateContent>`;
const wrapped = (requires: string, ns: string, choice: string, fallback: string): string =>
  `<mc:AlternateContent xmlns:mc="${MC}"><mc:Choice xmlns:${requires}="${ns}" Requires="${requires}">${choice}</mc:Choice><mc:Fallback>${fallback}</mc:Fallback></mc:AlternateContent>`;

describe("MCE wrappers", () => {
  it("reads the Fallback (no namespace is understood) and keeps the whole wrapper as sourceXml", async () => {
    expect(UNDERSTOOD_SHEET_DRAWING_NAMESPACES.size).toBe(0);
    const { model } = await open({ drawing: FORM_CONTROL });
    expect(model.sheets[0]!.drawings![0]!.sourceXml).toBe(FORM_CONTROL);
  });

  it("a form control: 'Form control', named from the Choice, no face, not hidden (hidden only from the branch read), geometry from the Choice", async () => {
    const { model } = await open({ drawing: FORM_CONTROL });
    const [drawing] = drawingsOf(model);
    expect(drawing).toMatchObject({ name: "Check Box 1", content: { kind: "foreign", label: "Form control" }, anchor: { kind: "twoCell", editAs: "absolute", posEmu: { x: 5 * 609600, y: 11 * 190500 } } });
    expect(drawing!.hidden).toBeUndefined();
    expect((drawing!.content as { face?: unknown }).face).toBeUndefined();
  });

  it("a slicer: 'Slicer', named 'Region' from the Choice (the Fallback's name is empty), its text-box Fallback the face", async () => {
    const { model } = await open({ drawing: SLICER });
    const [drawing] = drawingsOf(model);
    expect(drawing).toMatchObject({ name: "Region", content: { kind: "foreign", label: "Slicer", face: { kind: "shape", visual: { geometry: "rect", fill: { kind: "solid", color: "#ffffff" } } } } });
    expect((drawing!.content as { face: { text: { paragraphs: { text: string }[] } } }).face.text.paragraphs[0]!.text).toBe(SLICER_TEXT);
  });

  it("ink: 'Ink' with its picture Fallback as the face", async () => {
    const { model } = await open({ drawing: INK, drawingRels: rel("rId1", "image", "../media/image1.png"), media: { "xl/media/image1.png": PNG }, defaults: PNG_DEFAULT });
    expect(drawingsOf(model)[0]).toMatchObject({ name: "Ink 11", content: { kind: "foreign", label: "Ink", face: { kind: "image", mediaId: "xl/media/image1.png" } } });
  });

  it("chartex is a 'Chart' never given a face; an a14 text body with a14:m an 'Equation' with its Fallback face; a hidden Fallback no face and hidden; a wrapper inside an anchor too", async () => {
    const box = (name: string, attrs = ""): string => sp(2, name, "rect", "", attrs).replace("</xdr:spPr>", `</xdr:spPr><xdr:txBody><a:bodyPr/><a:p><a:r><a:t>x = 1</a:t></a:r></a:p></xdr:txBody>`);
    const anchor = (inner: string): string => twoCell([0, 0, 0, 0], [2, 0, 2, 0], inner);
    const equationChoice = anchor(sp(2, "TextBox 1", "rect").replace("</xdr:spPr>", `</xdr:spPr><xdr:txBody><a:bodyPr/><a:p><a14:m><m:oMathPara xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"/></a14:m></a:p></xdr:txBody>`));
    const { model } = await open({
      drawing:
        wrapped("cx1", "http://schemas.microsoft.com/office/drawing/2015/9/8/chartex", anchor(`<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Chart 1"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr></xdr:graphicFrame>`), anchor(box("Chart 1"))) +
        wrapped("a14", A14, equationChoice, anchor(box("TextBox 1"))) +
        wrapped("a16", "http://schemas.microsoft.com/office/drawing/2014/main", anchor(box("Thing 1")), anchor(box("Thing 1", ` hidden="1"`))) +
        twoCell([4, 0, 4, 0], [6, 0, 6, 0], `<mc:AlternateContent xmlns:mc="${MC}"><mc:Choice xmlns:aink="http://schemas.microsoft.com/office/drawing/2016/ink" Requires="aink"><xdr:contentPart r:id="rId3"/></mc:Choice><mc:Fallback>${box("Ink 2")}</mc:Fallback></mc:AlternateContent>`),
    });
    const [chartex, equation, hidden, inner] = drawingsOf(model);
    expect(chartex!.content).toEqual({ kind: "foreign", label: "Chart" });
    expect(equation!.content).toMatchObject({ kind: "foreign", label: "Equation", face: { kind: "shape", text: { paragraphs: [{ text: "x = 1" }] } } });
    expect(hidden).toMatchObject({ hidden: true, content: { kind: "foreign", label: "Object" } });
    expect((hidden!.content as { face?: unknown }).face).toBeUndefined();
    expect(inner).toMatchObject({ name: "Ink 2", anchor: { from: { col: 4, row: 4 } }, content: { kind: "foreign", label: "Ink", face: { kind: "shape" } } });
  });
});

// --- the canonical serialiser ------------------------------------------------------------------

/** A default 64 × 20 px grid sheet. */
function defaultSheet(): Sheet {
  return createBlankWorkbook().sheets[0]!;
}

describe("canonical new drawings (xlsx/drawingGolden.ts)", () => {
  const metrics = buildGridMetrics(defaultSheet());

  it("serialises each golden byte for byte", () => {
    expect(canonicalAnchorXml(G.GOLDEN_IMAGE_DRAWING, metrics, { cNvPrId: 2, embedRId: "rId1" })).toBe(G.GOLDEN_IMAGE_ANCHOR);
    expect(canonicalAnchorXml(G.GOLDEN_SHAPE_DRAWING, metrics, { cNvPrId: 3 })).toBe(G.GOLDEN_SHAPE_ANCHOR);
    expect(canonicalAnchorXml(G.GOLDEN_LINE_DRAWING, metrics, { cNvPrId: 5 })).toBe(G.GOLDEN_LINE_ANCHOR);
    expect(canonicalAnchorXml(G.GOLDEN_TEXT_SHAPE_DRAWING, metrics, { cNvPrId: 7 })).toBe(G.GOLDEN_TEXT_SHAPE_ANCHOR);
    expect(canonicalAnchorXml(G.GOLDEN_ABSOLUTE_IMAGE_DRAWING, metrics, { cNvPrId: 9, embedRId: "rId2" })).toBe(G.GOLDEN_ABSOLUTE_IMAGE_ANCHOR);
  });

  it("writes a new part and its rels as the goldens, minting cNvPr id k + 1 from each name", () => {
    const written = writeDrawingPart({ partName: "xl/drawings/drawing1.xml", drawings: [G.GOLDEN_IMAGE_DRAWING, G.GOLDEN_SHAPE_DRAWING], metrics, mediaPath: (id) => id });
    expect(written.xml).toBe(G.GOLDEN_DRAWING_PART);
    expect(written.rels).toBe(G.GOLDEN_DRAWING_RELS);
    expect(written.targets).toEqual(["xl/media/image1.png"]);
  });

  it("puts <drawing r:id> at its CT_Worksheet place and takes it out again", () => {
    const sheet = `<worksheet xmlns="${MAIN}" xmlns:r="${R}"><sheetData><row r="1"/></sheetData><pageMargins left="0.7"/><legacyDrawing r:id="rId9"/></worksheet>`;
    const added = spliceWorksheetDrawing(sheet, "rId1");
    expect(added).toBe(sheet.replace("<legacyDrawing", `${G.GOLDEN_WORKSHEET_DRAWING}<legacyDrawing`));
    expect(spliceWorksheetDrawing(added, null)).toBe(sheet);
    const headerFooter = `<worksheet xmlns="${MAIN}" xmlns:r="${R}"><sheetData/><pageMargins left="0.7"/><drawingHF r:id="rId9"/></worksheet>`;
    expect(spliceWorksheetDrawing(headerFooter, "rId1")).toBe(headerFooter.replace("<drawingHF", `${G.GOLDEN_WORKSHEET_DRAWING}<drawingHF`));
  });
});

// --- patching and the package --------------------------------------------------------------

const SHAPE_ANCHOR = twoCell([4, 0, 2, 0], [6, 0, 5, 0], sp(3, "Rounded Rectangle 2", "roundRect"));
const MC_BRANCH = twoCell([4, 0, 8, 0], [6, 0, 10, 0], sp(4, "Thing 3", "rect"));
const MC_DRAWING = `<mc:AlternateContent xmlns:mc="${MC}"><mc:Choice xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" Requires="a16">${MC_BRANCH}</mc:Choice><mc:Fallback>${MC_BRANCH}</mc:Fallback></mc:AlternateContent>`;
const pictureBook = (): Promise<{ bytes: Uint8Array; model: SpreadsheetModel }> =>
  open({ drawing: PICTURE + SHAPE_ANCHOR + MC_DRAWING, drawingRels: rel("rId1", "image", "../media/image1.png"), media: { "xl/media/image1.png": PNG }, defaults: PNG_DEFAULT });

/** The text of a drawing's content element (its xdr:sp / xdr:pic). */
const contentOf = (xml: string, local: "sp" | "pic"): string[] => [...xml.matchAll(new RegExp(`<xdr:${local}\\b[^]*?</xdr:${local}>`, "g"))].map((match) => match[0]);
const withoutBox = (xml: string): string => xml.replace(/<a:off [^>]*\/><a:ext [^>]*\/>/g, "");
/** Drawings as an equality compares them: no minted id, no source text. */
const strip = (drawings: readonly SheetDrawing[]): unknown[] => drawings.map(({ sourceXml: _s, id: _id, ...rest }) => rest);

describe("writing a touched layer", () => {
  it("a clean workbook is written verbatim; a dirty layer nothing changed rebuilds byte-identical", async () => {
    const { bytes, model } = await pictureBook();
    const source = await unzip(bytes);
    const clean = await unzip(await writeXlsx(model));
    for (const [name, data] of source) {
      expect(decoder.decode(clean.get(name)), name).toBe(decoder.decode(data));
    }
    model.sheets[0] = { ...model.sheets[0]!, drawingsDirty: true };
    const out = await save(model);
    for (const name of ["xl/drawings/drawing1.xml", "xl/drawings/_rels/drawing1.xml.rels", "xl/media/image1.png", "xl/worksheets/sheet1.xml"]) {
      expect(decoder.decode(out.get(name)), name).toBe(decoder.decode(source.get(name)));
    }
  });

  it("rows inserted above move the anchors (every MCE branch alike) and their a:off; the content elements are otherwise untouched", async () => {
    const { bytes, model } = await pictureBook();
    const before = text(await unzip(bytes), "xl/drawings/drawing1.xml");
    expect(applyStructuralEdit(model, { kind: "insertRows", sheetId: model.sheets[0]!.id, at: 0, count: 2 }).ok).toBe(true);
    const xml = text(await save(model), "xl/drawings/drawing1.xml");
    expect(xml).toContain(corner("from", [4, 0, 4, 0]) + corner("to", [6, 0, 7, 0]));
    expect(xml.split(corner("from", [4, 0, 10, 0]) + corner("to", [6, 0, 12, 0]))).toHaveLength(3);
    expect(xml).toContain(`<a:off x="${4 * 609600}" y="${4 * 190500}"/>`);
    expect(contentOf(xml, "sp").map(withoutBox)).toEqual(contentOf(before, "sp").map(withoutBox));
    expect(contentOf(xml, "pic").map(withoutBox)).toEqual(contentOf(before, "pic").map(withoutBox));
    // A second write of what was written is stable.
    const reread = await readSpreadsheetBytes(await writeXlsx(model), "xlsx", "again");
    reread.sheets[0] = { ...reread.sheets[0]!, drawingsDirty: true };
    expect(text(await save(reread), "xl/drawings/drawing1.xml")).toBe(xml);
  });

  it("a move keeps a move anchor's size (its to recomputed), alt text patches cNvPr@descr only, editAs changes in place", async () => {
    const { model } = await pictureBook();
    const sheetId = model.sheets[0]!.id;
    applyDrawingResult(model, setDrawingRect(model, sheetId, "dr1", { left: 64 * 3, top: 20, width: 96, height: 64 })!);
    applyDrawingResult(model, setDrawingAltText(model, sheetId, "dr2", "Assumptions box")!);
    const sheet = model.sheets[0]!;
    model.sheets[0] = { ...sheet, drawings: sheet.drawings!.map((drawing) => (drawing.id === "dr3" ? { ...drawing, anchor: { ...drawing.anchor, editAs: "oneCell" as const } } : drawing)) };
    const xml = text(await save(model), "xl/drawings/drawing1.xml");
    expect(xml).toContain(`<xdr:twoCellAnchor editAs="oneCell">${corner("from", [3, 0, 1, 0])}${corner("to", [4, 304800, 4, 38100])}`);
    expect(xml).toContain(`<xdr:cNvPr id="3" name="Rounded Rectangle 2" descr="Assumptions box"/>`);
    expect(xml.match(/<xdr:twoCellAnchor editAs="oneCell">/g)).toHaveLength(3);
  });

  it("typed alt text keeps its line breaks and tabs through save → reopen, patched or canonical; an apostrophe in a single-quoted descr stays well-formed", async () => {
    const typed = "John's chart\n\tsecond line";
    const { model } = await open({ drawing: twoCell([4, 0, 2, 0], [6, 0, 5, 0], sp(3, "Rounded Rectangle 2", "roundRect", "", ` descr='Old'`)) });
    const sheetId = model.sheets[0]!.id;
    applyDrawingResult(model, setDrawingAltText(model, sheetId, "dr1", typed)!);
    const created = insertShape(model, sheetId, "rect", { left: 256, top: 40, width: 128, height: 100 })!;
    applyDrawingResult(model, created);
    applyDrawingResult(model, setDrawingAltText(model, sheetId, created.drawingId!, typed)!);
    const xml = text(await save(model), "xl/drawings/drawing1.xml");
    expect(xml).toContain(`<xdr:cNvPr id="3" name="Rounded Rectangle 2" descr='John&apos;s chart&#xA;&#x9;second line'/>`);
    expect(xml).toContain(`descr="John's chart&#xA;&#x9;second line"`);
    const reread = await readSpreadsheetBytes(await writeXlsx(model), "xlsx", "again");
    expect(reread.sheets[0]!.drawings!.map((drawing) => drawing.altText)).toEqual([typed, typed]);
  });

  it("a first drawing on a sheet adds the part, its rels, <drawing>, the sheet relationship, the Override, the media part and its Default", async () => {
    const { model } = await open();
    const sheetId = model.sheets[1]!.id;
    applyDrawingResult(model, insertImage(model, sheetId, { bytes: PNG, contentType: "image/png", naturalPx: { width: 96, height: 64 }, at: { row: 1, col: 1 }, visiblePx: { width: 1000, height: 800 } })!);
    applyDrawingResult(model, insertShape(model, sheetId, "rect", { left: 256, top: 40, width: 128, height: 100 })!);
    const out = await save(model);
    expect(text(out, "xl/worksheets/sheet2.xml")).toContain(`<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>${G.GOLDEN_WORKSHEET_DRAWING}</worksheet>`);
    expect(text(out, "xl/worksheets/_rels/sheet2.xml.rels")).toContain(G.GOLDEN_SHEET_DRAWING_RELATIONSHIP);
    const types = text(out, "[Content_Types].xml");
    expect(types).toContain(G.GOLDEN_DRAWING_OVERRIDE);
    expect(types).toContain(PNG_DEFAULT);
    expect(text(out, "xl/drawings/_rels/drawing1.xml.rels")).toBe(G.GOLDEN_DRAWING_RELS);
    expect(out.get("xl/media/image1.png")).toEqual(PNG);
    const drawing = text(out, "xl/drawings/drawing1.xml");
    expect(drawing).toContain(`<xdr:cNvPr id="2" name="Picture 1"/>`);
    expect(drawing).toContain(`<xdr:cNvPr id="3" name="Rectangle 2"/>`);
    const reread = await readSpreadsheetBytes(await writeXlsx(model), "xlsx", "again");
    expect(strip(reread.sheets[1]!.drawings!)).toEqual(strip(model.sheets[1]!.drawings!));
  });

  it("the last drawing deleted removes <drawing>, the part, its rels, the relationship and the Override; an unreferenced image goes, one another sheet embeds survives", async () => {
    const options: BookOptions = {
      drawing: PICTURE,
      drawingRels: rel("rId1", "image", "../media/image1.png"),
      drawing2: twoCell([0, 0, 0, 0], [1, 0, 1, 0], pic(2, "Picture 1", `r:embed="rId1"`), "oneCell"),
      drawing2Rels: rel("rId1", "image", "../media/image1.png"),
      media: { "xl/media/image1.png": PNG },
      defaults: PNG_DEFAULT,
    };
    const shared = await open(options);
    applyDrawingResult(shared.model, deleteDrawing(shared.model, shared.model.sheets[0]!.id, "dr1")!);
    const out = await save(shared.model);
    expect(text(out, "xl/worksheets/sheet1.xml")).not.toContain("<drawing");
    expect(out.has("xl/drawings/drawing1.xml")).toBe(false);
    expect(out.has("xl/drawings/_rels/drawing1.xml.rels")).toBe(false);
    expect(text(out, "xl/worksheets/_rels/sheet1.xml.rels")).not.toContain("drawing1.xml");
    expect(text(out, "[Content_Types].xml")).not.toContain("/xl/drawings/drawing1.xml");
    expect(out.get("xl/media/image1.png")).toEqual(PNG);

    const alone = await open({ ...options, drawing2: undefined, drawing2Rels: undefined });
    applyDrawingResult(alone.model, deleteDrawing(alone.model, alone.model.sheets[0]!.id, "dr1")!);
    expect((await save(alone.model)).has("xl/media/image1.png")).toBe(false);
  });

  it("deleting one of two drawings drops only the relationship nobody uses any more", async () => {
    const { model } = await open({
      drawing: PICTURE + twoCell([4, 0, 2, 0], [5, 0, 4, 0], pic(3, "Picture 2", `r:embed="rId2"`), "oneCell"),
      drawingRels: rel("rId1", "image", "../media/image1.png") + rel("rId2", "image", "../media/image2.png"),
      media: { "xl/media/image1.png": PNG, "xl/media/image2.png": PNG },
      defaults: PNG_DEFAULT,
    });
    applyDrawingResult(model, deleteDrawing(model, model.sheets[0]!.id, "dr2")!);
    const out = await save(model);
    expect(text(out, "xl/drawings/_rels/drawing1.xml.rels")).toBe(rels(rel("rId1", "image", "../media/image1.png")));
    expect(text(out, "xl/drawings/drawing1.xml")).toBe(wsDr(PICTURE));
    expect(out.has("xl/media/image2.png")).toBe(false);
  });

  it("substitutedXlsxParts lays the retained drawing part out as it is", async () => {
    const { bytes, model } = await pictureBook();
    applyDrawingResult(model, deleteDrawing(model, model.sheets[0]!.id, "dr1")!);
    expect(decoder.decode(substitutedXlsxParts(model).get("xl/drawings/drawing1.xml"))).toBe(text(await unzip(bytes), "xl/drawings/drawing1.xml"));
  });

  it("writes created drawings from scratch with their parts, and reopens them equal", async () => {
    const model = createBlankWorkbook();
    const sheetId = model.sheets[0]!.id;
    applyDrawingResult(model, insertImage(model, sheetId, { bytes: PNG, contentType: "image/png", naturalPx: { width: 96, height: 64 }, at: { row: 2, col: 1 }, visiblePx: { width: 1000, height: 800 } })!);
    applyDrawingResult(model, insertShape(model, sheetId, "ellipse", { left: 300, top: 50, width: 80, height: 60 })!);
    const sheet = model.sheets[0]!;
    model.sheets[0] = { ...sheet, drawings: [...sheet.drawings!, { ...G.GOLDEN_TEXT_SHAPE_DRAWING, id: "dr3" }] };
    const out = await save(model);
    expect(text(out, "xl/worksheets/sheet1.xml")).toContain(G.GOLDEN_WORKSHEET_DRAWING);
    expect(text(out, "xl/worksheets/_rels/sheet1.xml.rels")).toContain(G.GOLDEN_SHEET_DRAWING_RELATIONSHIP);
    expect(text(out, "[Content_Types].xml")).toContain(G.GOLDEN_DRAWING_OVERRIDE);
    expect(text(out, "[Content_Types].xml")).toContain(PNG_DEFAULT);
    expect(text(out, "xl/drawings/drawing1.xml")).toContain(G.GOLDEN_TEXT_SHAPE_ANCHOR);
    const reread = await readSpreadsheetBytes(await writeXlsx(model), "xlsx", "again");
    expect(strip(reread.sheets[0]!.drawings!)).toEqual(strip(model.sheets[0]!.drawings!));
    expect(reread.media!.map((part) => part.mediaId)).toEqual(["xl/media/image1.png"]);
  });
});
