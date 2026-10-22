// SPEC-131 "Reading ODS drawings", "Writing" (ods), "Canonical new drawings"
// and "Media parts" over small inline workbooks: every reading rule, golden
// equality of the canonical serialiser (ods/drawingGolden.ts), anchor
// patching that leaves the element's content byte-identical, an untouched
// drawing re-emitted verbatim by a dirty table, a renamed sheet's end cell,
// the materialised anchor cell, media parts and manifest entries in both
// write paths, and the reader-upgrade lift.
import { describe, expect, it } from "vitest";
import { equalBytes } from "../../core/equality";
import { unzip } from "../../core/zip";
import { zip } from "../../core/zipWriter";
import { getCell } from "../address";
import { deleteDrawing, EMU_PER_PX, insertImage, insertShape, newShapeVisual, setDrawingRect, applyDrawingResult, drawingRectPx } from "../drawings";
import { buildGridMetrics } from "../geometry";
import type { Sheet, SheetDrawing, SpreadsheetModel } from "../model";
import { readSpreadsheetBytes } from "../open";
import { createBlankWorkbook, renameSheet } from "../sheets";
import {
  GOLDEN_DRAWING_SHEET,
  GOLDEN_ODS_ARROW_STYLE,
  GOLDEN_ODS_FIXED_SHAPE,
  GOLDEN_ODS_IMAGE,
  GOLDEN_ODS_IMAGE_STYLE,
  GOLDEN_ODS_IMAGE_WITH_ALT_TEXT,
  GOLDEN_ODS_LINE,
  GOLDEN_ODS_LINE_STYLE,
  GOLDEN_ODS_SHAPE,
  GOLDEN_ODS_SHAPE_QUOTED_WITH_ALT_TEXT,
  GOLDEN_ODS_SHAPE_STYLE,
  GOLDEN_ODS_TABLE_SHAPES,
} from "./drawingGolden";
import { readOdsChart } from "./chartXml";
import { canonicalOdsDrawing, odsSheetDrawingsXml, odsShapeGraphicProperties, upgradeOdsSheetDrawings } from "./drawingXml";
import { OdsStyleRegistry } from "./styles";
import { writeOds } from "./writer";

const encoder = new TextEncoder();
const decoder = new TextDecoder();
const IN = 914400; // EMU per inch

const NS =
  `xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" ` +
  `xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" ` +
  `xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" ` +
  `xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" ` +
  `xmlns:dr3d="urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0" xmlns:calcext="urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0"`;

/** A PNG's signature and IHDR (all decodeImageDimensions reads). */
function png(width: number, height: number): Uint8Array {
  const bytes = new Uint8Array(33);
  bytes.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 13, 0x49, 0x48, 0x44, 0x52]);
  new DataView(bytes.buffer).setUint32(16, width);
  new DataView(bytes.buffer).setUint32(20, height);
  bytes.set([8, 6, 0, 0, 0], 24);
  return bytes;
}

const STYLES =
  `<style:style style:name="co1" style:family="table-column"><style:table-column-properties style:column-width="1in"/></style:style>` +
  `<style:style style:name="ro1" style:family="table-row"><style:table-row-properties style:row-height="0.25in"/></style:style>` +
  `<style:style style:name="gr1" style:family="graphic"><style:graphic-properties draw:fill="none" draw:stroke="solid" svg:stroke-color="#ff0000" svg:stroke-width="2pt"/></style:style>` +
  `<style:style style:name="gr2" style:family="graphic"><style:graphic-properties draw:fill="solid" draw:fill-color="#00ff00" draw:stroke="none" draw:textarea-vertical-align="top"/></style:style>` +
  `<style:style style:name="gr3" style:family="graphic"><style:graphic-properties draw:stroke="none" draw:fill="none" fo:clip="rect(0in, 0.1in, 0in, 0.1in)"/></style:style>` +
  `<style:style style:name="P1" style:family="paragraph"><style:paragraph-properties fo:text-align="center"/></style:style>` +
  `<style:style style:name="T1" style:family="text"><style:text-properties fo:font-size="14pt" fo:font-weight="bold" fo:color="#112233"/></style:style>`;

const BOX =
  `<draw:custom-shape table:end-cell-address="Sheet1.E4" table:end-x="0.25in" table:end-y="0in" draw:z-index="0" draw:name="Box" draw:style-name="gr2" ` +
  `svg:width="2in" svg:height="1in" svg:x="2.5in" svg:y="0in"><text:p text:style-name="P1"><text:span text:style-name="T1">Hello</text:span></text:p>` +
  `<draw:enhanced-geometry svg:viewBox="0 0 21600 21600" draw:type="round-rectangle" draw:modifiers="25000"/></draw:custom-shape>`;
const PAGE = `<draw:rect draw:z-index="1" draw:name="Page" draw:style-name="gr1" svg:x="1in" svg:y="0.5in" svg:width="2in" svg:height="1in"/>`;
const PIC = `<draw:frame draw:z-index="3" draw:name="Pic" draw:style-name="gr3" svg:x="0.5in" svg:y="0in" svg:width="1in" svg:height="0.5in"><draw:image xlink:href="Pictures/a.png"/><svg:desc>Logo</svg:desc></draw:frame>`;
const CHART = `<draw:frame draw:z-index="2" svg:x="3in" svg:y="0.25in" svg:width="1in" svg:height="1in"><draw:object xlink:href="./Object 1"/></draw:frame>`;
const ANNOTATION = `<office:annotation><text:p>note</text:p></office:annotation>`;
const GROUP = `<draw:g draw:name="G"><draw:rect svg:x="0in" svg:y="0.5in" svg:width="1in" svg:height="1in"/><draw:ellipse svg:x="2in" svg:y="0.5in" svg:width="1in" svg:height="1in"/></draw:g>`;
const ROTATED = `<draw:custom-shape draw:z-index="4" svg:width="1in" svg:height="0.5in" draw:transform="rotate (1.5707963267949) translate (2in 1.5in)"><draw:enhanced-geometry draw:type="non-primitive"/></draw:custom-shape>`;
const CONTROL = `<draw:control draw:z-index="5" svg:x="2in" svg:y="0.5in" svg:width="1in" svg:height="0.25in"/>`;
const LINE = `<draw:line draw:z-index="6" svg:x1="3in" svg:y1="0.5in" svg:x2="4in" svg:y2="1in"/>`;
const SCENE = `<dr3d:scene draw:z-index="7" svg:x="4in" svg:y="0.5in" svg:width="1in" svg:height="1in"/>`;
const OLE = `<draw:frame draw:z-index="8" svg:x="5in" svg:y="0.5in" svg:width="1in" svg:height="1in"><draw:object-ole xlink:href="./Object 2"/></draw:frame>`;

const cell = (inner: string, attrs = ""): string => `<table:table-cell${attrs}>${inner}</table:table-cell>`;
const empty = (n: number): string => `<table:table-cell table:number-columns-repeated="${n}"/>`;
const row = (cells: string, attrs = ""): string => `<table:table-row table:style-name="ro1"${attrs}>${cells}</table:table-row>`;

const TABLE =
  `<table:table table:name="Sheet1"><office:forms/><table:shapes>${PAGE}</table:shapes>` +
  `<table:table-column table:style-name="co1" table:number-columns-repeated="16384"/>` +
  row(empty(2) + cell(BOX) + empty(16381)) +
  row(cell(PIC, ` table:number-columns-repeated="3"`) + cell(`${ANNOTATION}${CHART}<text:p>x</text:p>`, ` office:value-type="string"`) + empty(16380)) +
  row(cell(GROUP) + cell(ROTATED) + cell(CONTROL) + cell(LINE) + cell(SCENE) + cell(OLE) + empty(16378)) +
  row(empty(16384), ` table:number-rows-repeated="1048573"`) +
  `</table:table>`;

const OFFICE_CHART_NS =
  `xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0" ` +
  `xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"`;
/** An embedded object's content.xml with `body` as its office:body child. */
const objectContent = (body: string): string =>
  `<?xml version="1.0" encoding="UTF-8"?><office:document-content ${OFFICE_CHART_NS} office:version="1.2"><office:body>${body}</office:body></office:document-content>`;
/** SPEC-132: CHART's object — a chart outside the catalogue (a surface chart), so it stays unmodelled. */
const OBJECT_1 = objectContent(`<office:chart><chart:chart chart:class="chart:surface"><chart:plot-area/></chart:chart></office:chart>`);

const manifestOf = (entries: string): string =>
  `<?xml version="1.0" encoding="UTF-8"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.2">` +
  `<manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/>` +
  `<manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>` +
  `<manifest:file-entry manifest:full-path="Pictures/a.png" manifest:media-type="image/png"/>` +
  `<manifest:file-entry manifest:full-path="Object 1/" manifest:media-type="application/vnd.oasis.opendocument.chart"/>` +
  `<manifest:file-entry manifest:full-path="Object 1/content.xml" manifest:media-type="text/xml"/>${entries}</manifest:manifest>`;
const MANIFEST = manifestOf("");

/** The workbook's package; `extra` adds parts (and their manifest entries). */
async function workbookBytes(table = TABLE, extra: { parts?: { name: string; text: string }[]; manifest?: string } = {}): Promise<Uint8Array> {
  const content = `<?xml version="1.0" encoding="UTF-8"?><office:document-content ${NS} office:version="1.2"><office:automatic-styles>${STYLES}</office:automatic-styles><office:body><office:spreadsheet>${table}</office:spreadsheet></office:body></office:document-content>`;
  return zip([
    { name: "mimetype", data: encoder.encode("application/vnd.oasis.opendocument.spreadsheet"), store: true },
    { name: "META-INF/manifest.xml", data: encoder.encode(manifestOf(extra.manifest ?? "")) },
    { name: "content.xml", data: encoder.encode(content) },
    { name: "Pictures/a.png", data: png(96, 48) },
    { name: "Object 1/content.xml", data: encoder.encode(OBJECT_1) },
    ...(extra.parts ?? []).map((part) => ({ name: part.name, data: encoder.encode(part.text) })),
  ]);
}

async function open(table = TABLE, extra?: Parameters<typeof workbookBytes>[1]): Promise<SpreadsheetModel> {
  return readSpreadsheetBytes(await workbookBytes(table, extra), "ods", "t");
}

/** A group drawing's members as [x, y, w, h] (to 9 places) and kind. */
const groupMembers = (drawing: SheetDrawing): unknown =>
  drawing.content.kind === "group" && drawing.content.members.map((member) => [member.x, member.y, member.w, member.h].map((n) => n.toFixed(9)).concat(member.content.kind));
const named = (sheet: Sheet, name: string): SheetDrawing => sheet.drawings!.find((drawing) => drawing.name === name)!;
/** SPEC-132: the chart content a frame showing `sourcePart` reads as (ods/chartXml.ts readOdsChart's read, `uncached` never stored). */
const chartRead = (text: string, sourcePart: string): SheetDrawing["content"] => {
  const read = readOdsChart(text);
  return read.refs ? { kind: "chart", visual: read.visual, refs: read.refs, sourcePart } : { kind: "chart", visual: read.visual, sourcePart };
};
const at = (sheet: Sheet, index: number): SheetDrawing => sheet.drawings![index]!;

/** The golden geometry's sheet (drawingGolden.ts). */
const goldenSheet = (name: string = GOLDEN_DRAWING_SHEET.name): Sheet => ({
  id: "s",
  name,
  visibility: "visible",
  kind: "grid",
  rows: [],
  columns: [],
  defaultColWidthPx: GOLDEN_DRAWING_SHEET.defaultColWidthPx,
  defaultRowHeightPt: GOLDEN_DRAWING_SHEET.defaultRowHeightPt,
  merges: [],
  conditionalFormats: [],
  validations: [],
});
const px = (n: number): number => n * EMU_PER_PX;

/** The workbook with its sheet marked dirty and B2's row edited. */
function dirty(model: SpreadsheetModel): void {
  const first = model.sheets[0]!;
  model.sheets = [{ ...first, dirty: true, rows: first.rows.map((each) => (each.index === 1 ? { ...each, cells: each.cells.map((c) => ({ ...c, value: "y" })) } : each)) }];
}

async function written(model: SpreadsheetModel): Promise<{ parts: Map<string, Uint8Array>; content: string; bytes: Uint8Array }> {
  const bytes = await writeOds(model);
  const parts = await unzip(bytes);
  return { parts, content: decoder.decode(parts.get("content.xml")), bytes };
}


/** A frame showing an embedded object, as SPEC-132's object-kind test lays them out. */
function objectFrame(object: string, z: number): string {
  return `<draw:frame draw:z-index="${z}" svg:x="0in" svg:y="0in" svg:width="1in" svg:height="1in"><draw:object xlink:href="./${object}"/></draw:frame>`;
}
describe("reading ods drawings", () => {
  it("orders the layer by draw:z-index, a drawing without one in front, and mints ids in that order", async () => {
    const sheet = (await open()).sheets[0]!;
    expect(sheet.drawings!.map((drawing) => [drawing.id, drawing.name])).toEqual([
      ["dr1", "Box"],
      ["dr2", "Page"],
      ["dr3", ""],
      ["dr4", "Pic"],
      ["dr5", ""],
      ["dr6", ""],
      ["dr7", ""],
      ["dr8", ""],
      ["dr9", ""],
      ["dr10", "G"],
    ]);
  });

  it("reads a table:shapes child as absolute: position from svg:x/svg:y, size from svg:width/svg:height", async () => {
    expect(named((await open()).sheets[0]!, "Page").anchor).toEqual({ kind: "absolute", posEmu: { x: IN, y: IN / 2 }, extEmu: { cx: 2 * IN, cy: IN } });
  });

  it("reads an in-cell element with an end cell as twoCell editAs twoCell, its sheet-absolute svg:x the offset into its cell", async () => {
    expect(named((await open()).sheets[0]!, "Box").anchor).toEqual({
      kind: "twoCell",
      editAs: "twoCell",
      from: { col: 2, colOffEmu: IN / 2, row: 0, rowOffEmu: 0 },
      to: { col: 4, colOffEmu: IN / 4, row: 3, rowOffEmu: 0 },
    });
  });

  it("reads one without an end cell as oneCell sized by svg:width/svg:height, a repeated cell at its first column, a negative offset clamped at 0", async () => {
    expect(named((await open()).sheets[0]!, "Pic").anchor).toEqual({ kind: "oneCell", from: { col: 0, colOffEmu: IN / 2, row: 1, rowOffEmu: 0 }, extEmu: { cx: IN, cy: IN / 2 } });
  });

  it("places a draw:transform's box by its centre and reads its rotate as clockwise degrees", async () => {
    const rotated = at((await open()).sheets[0]!, 4);
    expect(rotated.rotationDeg).toBe(270);
    expect(rotated.anchor.from).toEqual({ col: 1, colOffEmu: (3 * IN) / 4, row: 2, rowOffEmu: IN / 4 });
    expect(rotated.content).toMatchObject({ kind: "shape", visual: { geometry: "rect", unsupportedGeometryXml: expect.stringContaining("non-primitive") } });
  });

  it("reads every content kind", async () => {
    const sheet = (await open()).sheets[0]!;
    expect(named(sheet, "Pic").content).toEqual({ kind: "image", mediaId: "Pictures/a.png", crop: { l: 0.1, t: 0, r: 0.1, b: 0 } });
    // SPEC-132: a chart object's content is what readOdsChart reads from its part.
    expect(at(sheet, 2).content).toEqual(chartRead(OBJECT_1, "Object 1/content.xml"));
    expect(at(sheet, 5).content).toEqual({ kind: "foreign", label: "Form control" });
    expect(at(sheet, 6).content).toMatchObject({ kind: "shape", visual: { geometry: "line" } });
    expect(at(sheet, 7).content).toEqual({ kind: "foreign", label: "3-D object" });
    expect(at(sheet, 8).content).toEqual({ kind: "foreign", label: "Object" });
    const group = named(sheet, "G");
    expect(group.anchor).toEqual({ kind: "oneCell", from: { col: 0, colOffEmu: 0, row: 2, rowOffEmu: 0 }, extEmu: { cx: 3 * IN, cy: IN } });
    expect(group.content.kind === "group" && group.content.members.map(({ x, y, w, h, content }) => [x, y, w, h, content.kind === "shape" && content.visual.geometry])).toEqual([
      [0, 0, 1 / 3, 1, "rect"],
      [2 / 3, 0, 1 / 3, 1, "ellipse"],
    ]);
  });

  it("SPEC-132: reads only a chart object as a chart — by the manifest's media type or an office:chart body; another embedded object is foreign Object, a missing part foreign Chart, a group's chart member foreign Chart", async () => {
    const frame = objectFrame;
    const table =
      `<table:table table:name="Sheet1"><table:table-column table:style-name="co1" table:number-columns-repeated="16384"/>` +
      row(cell(frame("Object 2", 0)) + cell(frame("Object 3", 1)) + cell(frame("Object 4", 2)) + cell(frame("Object 5", 3)) + cell(`<draw:g draw:name="G">${frame("Object 1", 4)}${frame("Object 2", 5)}</draw:g>`) + empty(16379)) +
      row(empty(16384), ` table:number-rows-repeated="1048575"`) +
      `</table:table>`;
    const formula = objectContent(`<office:formula/>`);
    const unlisted = objectContent(`<office:chart><chart:chart chart:class="chart:bar"/></office:chart>`);
    const sheet = (
      await open(table, {
        parts: [
          { name: "Object 2/content.xml", text: formula },
          { name: "Object 3/content.xml", text: unlisted },
          { name: "Object 5/content.xml", text: formula },
        ],
        manifest:
          `<manifest:file-entry manifest:full-path="Object 2/" manifest:media-type="application/vnd.oasis.opendocument.formula"/>` +
          `<manifest:file-entry manifest:full-path="Object 5/" manifest:media-type="application/vnd.oasis.opendocument.chart"/>`,
      })
    ).sheets[0]!;
    expect(sheet.drawings!.map((drawing) => drawing.content)).toEqual([
      { kind: "foreign", label: "Object" },
      chartRead(unlisted, "Object 3/content.xml"),
      { kind: "foreign", label: "Chart" },
      // The manifest says chart: whatever the body, the object is read as one.
      chartRead(formula, "Object 5/content.xml"),
      { kind: "group", members: [expect.objectContaining({ content: { kind: "foreign", label: "Chart" } }), expect.objectContaining({ content: { kind: "foreign", label: "Object" } })] },
    ]);
  });

  it("styles a shape from its graphic style and reads its text, alt text and name", async () => {
    const sheet = (await open()).sheets[0]!;
    expect(named(sheet, "Box").content).toEqual({
      kind: "shape",
      visual: { geometry: "roundRect", adjustments: [0.25], fill: { kind: "solid", color: "#00ff00" } },
      text: { paragraphs: [{ text: "Hello", align: "center" }], anchor: "top", sizePt: 14, color: "#112233", bold: true },
    });
    expect(named(sheet, "Page").content).toEqual({ kind: "shape", visual: { geometry: "rect", outline: { color: "#ff0000", widthPt: 2, dash: "solid" } } });
    // No style: LibreOffice's default fill and line.
    expect(at(sheet, 6).content).toMatchObject({ visual: { outline: { color: "#3465a4" } } });
    expect(named(sheet, "Pic").altText).toBe("Logo");
    expect(named(sheet, "Pic").sourceXml).toBe(PIC);
  });

  it("keeps drawings out of preservedXml — an annotation is the cell's note (SPEC-135) — and stores no cell holding only a drawing", async () => {
    const model = await open();
    const sheet = model.sheets[0]!;
    expect(getCell(sheet, 1, 3)).toEqual({ col: 3, value: "x" });
    expect(sheet.notes?.map((note) => [note.ref, note.text, note.source?.odsAnnotation])).toEqual([[{ row: 1, col: 3 }, "note", ANNOTATION]]);
    expect(getCell(sheet, 0, 2)).toBeUndefined();
    expect(getCell(sheet, 1, 0)).toBeUndefined();
    expect(model.media).toEqual([{ mediaId: "Pictures/a.png", contentType: "image/png", data: expect.any(String) }]);
  });
});

describe("canonical new ods drawings (drawingGolden.ts)", () => {
  const image: SheetDrawing = {
    id: "dr1",
    name: "Picture 1",
    anchor: { kind: "twoCell", editAs: "oneCell", from: { col: 1, colOffEmu: 0, row: 1, rowOffEmu: 0 }, to: { col: 3, colOffEmu: 0, row: 5, rowOffEmu: 0 }, extEmu: { cx: px(192), cy: px(96) } },
    content: { kind: "image", mediaId: "Pictures/image1.png" },
  };
  const shape: SheetDrawing = {
    id: "dr2",
    name: "Rectangle 2",
    anchor: { kind: "twoCell", editAs: "twoCell", from: { col: 2, colOffEmu: px(48), row: 2, rowOffEmu: 0 }, to: { col: 4, colOffEmu: 0, row: 5, rowOffEmu: px(12) } },
    content: { kind: "shape", visual: newShapeVisual("rect") },
  };
  const fixed: SheetDrawing = {
    id: "dr3",
    name: "Rounded Rectangle 3",
    anchor: { kind: "absolute", posEmu: { x: 0, y: 0 }, extEmu: { cx: px(96), cy: px(48) } },
    content: { kind: "shape", visual: { ...newShapeVisual("roundRect"), adjustments: [0.25] } },
  };
  const line: SheetDrawing = {
    id: "dr4",
    name: "Line 4",
    anchor: { kind: "twoCell", editAs: "twoCell", from: { col: 0, colOffEmu: 0, row: 0, rowOffEmu: 0 }, to: { col: 2, colOffEmu: 0, row: 1, rowOffEmu: 0 } },
    content: { kind: "shape", visual: newShapeVisual("line") },
  };
  const context = (registry: OdsStyleRegistry, name?: string) => ({ sheetName: name ?? GOLDEN_DRAWING_SHEET.name, metrics: buildGridMetrics(goldenSheet(name)), registry });

  it("writes each kind exactly as its golden, with its de-gr graphic style", () => {
    const registry = new OdsStyleRegistry();
    expect(canonicalOdsDrawing(image, 0, context(registry))).toBe(GOLDEN_ODS_IMAGE);
    expect(canonicalOdsDrawing(shape, 1, context(registry))).toBe(GOLDEN_ODS_SHAPE);
    expect(canonicalOdsDrawing(fixed, 2, context(registry))).toBe(GOLDEN_ODS_FIXED_SHAPE);
    expect(canonicalOdsDrawing(line, 3, context(registry))).toBe(GOLDEN_ODS_LINE);
    expect(registry.graphicStyle(odsShapeGraphicProperties(newShapeVisual("arrow")))).toBe("de-gr4");
    expect(registry.xml()).toBe(GOLDEN_ODS_IMAGE_STYLE + GOLDEN_ODS_SHAPE_STYLE + GOLDEN_ODS_LINE_STYLE + GOLDEN_ODS_ARROW_STYLE);
    expect(canonicalOdsDrawing({ ...image, altText: "Company logo" }, 0, context(registry))).toBe(GOLDEN_ODS_IMAGE_WITH_ALT_TEXT);
    expect(canonicalOdsDrawing({ ...shape, altText: "Note" }, 1, context(registry, "My Sheet"))).toBe(GOLDEN_ODS_SHAPE_QUOTED_WITH_ALT_TEXT);
  });

  it("places fixed drawings in table:shapes and the rest in their from cell, z-indexes in array order", () => {
    const xml = odsSheetDrawingsXml([image, shape, fixed, line], context(new OdsStyleRegistry()));
    expect(xml.shapes).toBe(GOLDEN_ODS_TABLE_SHAPES);
    expect([...xml.cells].map(([r, cols]) => [r, [...cols.keys()]])).toEqual([
      [1, [1]],
      [2, [2]],
      [0, [0]],
    ]);
    expect(xml.cells.get(1)!.get(1)).toBe(GOLDEN_ODS_IMAGE);
  });
});

describe("writing a dirty table's drawings", () => {
  it("re-emits every untouched drawing byte-identical, and an untouched workbook every part", async () => {
    const model = await open();
    const source = await unzip(await workbookBytes());
    const clean = await written(model);
    for (const [name, data] of source) {
      expect(equalBytes(clean.parts.get(name)!, data), name).toBe(true);
    }
    dirty(model);
    const { content } = await written(model);
    for (const xml of [BOX, PAGE, CHART, GROUP, ROTATED, CONTROL, LINE, SCENE, OLE]) {
      expect(content).toContain(xml);
    }
    // Pic's svg:y sits above its cell (read clamped): the rebuilt table writes where it is.
    expect(content).toContain(PIC.replace(`svg:y="0in"`, `svg:y="0.25in"`));
    expect(content).toContain(`<office:forms/><table:shapes>${PAGE}</table:shapes><table:table-column`);
    expect(content).toContain(`<table:table-cell office:value-type="string" calcext:value-type="string">${ANNOTATION}${CHART}<text:p>y</text:p></table:table-cell>`);
  });

  it("patches a moved drawing's anchor only, its content byte-identical, and reads it back where it was put", async () => {
    const model = await open();
    const sheet = model.sheets[0]!;
    const box = named(sheet, "Box");
    const metrics = buildGridMetrics(sheet);
    const rect = drawingRectPx(box.anchor, metrics);
    applyDrawingResult(model, setDrawingRect(model, sheet.id, box.id, { ...rect, top: rect.top + 48 })!);
    const { content, bytes } = await written(model);
    const element = /<draw:custom-shape table:end-cell-address[^>]*>[\s\S]*?<\/draw:custom-shape>/.exec(content)![0];
    expect(element.slice(element.indexOf(">") + 1)).toBe(BOX.slice(BOX.indexOf(">") + 1));
    expect(element).toContain(`svg:y="0.5in"`);
    expect(element).toContain(`svg:x="2.5in"`);
    expect(element).toContain(`table:end-cell-address="Sheet1.E6"`);
    const reread = (await readSpreadsheetBytes(bytes, "ods", "t2")).sheets[0]!;
    expect(named(reread, "Box").anchor).toEqual(model.sheets[0]!.drawings!.find((drawing) => drawing.name === "Box")!.anchor);
  });

  it("moves a group by its members — svg:x/svg:y, a draw:transform's translate, a line's ends, a nested group's — never a draw:transform on draw:g", async () => {
    const rotated = `<draw:custom-shape svg:width="1in" svg:height="0.5in" draw:transform="rotate (1.5707963267949) translate (2in 1.5in)"><draw:enhanced-geometry draw:type="non-primitive"/></draw:custom-shape>`;
    const group =
      `<draw:g draw:name="G2"><draw:rect svg:x="0in" svg:y="0.5in" svg:width="1in" svg:height="1in"/><draw:line svg:x1="1in" svg:y1="0.5in" svg:x2="2in" svg:y2="1in"/>` +
      `${rotated}<draw:g><draw:ellipse svg:x="2in" svg:y="0.5in" svg:width="1in" svg:height="1in"/></draw:g></draw:g>`;
    const model = await open(TABLE.replace(GROUP, group));
    const sheet = model.sheets[0]!;
    const moved = named(sheet, "G2");
    const rect = drawingRectPx(moved.anchor, buildGridMetrics(sheet));
    applyDrawingResult(model, setDrawingRect(model, sheet.id, moved.id, { ...rect, left: rect.left + 96, top: rect.top + 48 })!);
    const { content, bytes } = await written(model);
    expect(content).toContain(
      `<draw:g draw:name="G2"><draw:rect svg:x="1in" svg:y="1in" svg:width="1in" svg:height="1in"/><draw:line svg:x1="2in" svg:y1="1in" svg:x2="3in" svg:y2="1.5in"/>` +
        `${rotated.replace("translate (2in 1.5in)", "translate (3in 2in)")}<draw:g><draw:ellipse svg:x="3in" svg:y="1in" svg:width="1in" svg:height="1in"/></draw:g></draw:g>`,
    );
    const reread = named((await readSpreadsheetBytes(bytes, "ods", "t2")).sheets[0]!, "G2");
    const now = named(model.sheets[0]!, "G2");
    expect(reread.anchor).toEqual(now.anchor);
    // The members keep their places in the group (to float noise).
    expect(groupMembers(reread)).toEqual(groupMembers(moved));
  });

  it("writes a renamed sheet's touched drawing's end cell with the new name", async () => {
    const model = await open();
    renameSheet(model, 0, "My Data");
    const sheet = model.sheets[0]!;
    const box = named(sheet, "Box");
    const rect = drawingRectPx(box.anchor, buildGridMetrics(sheet));
    applyDrawingResult(model, setDrawingRect(model, sheet.id, box.id, { ...rect, width: rect.width + 96 })!);
    const { content } = await written(model);
    expect(content).toMatch(/table:end-cell-address="\$?'My Data'\.F4"/);
    expect(content).not.toContain("Sheet1.E4");
  });

  it("materialises a new drawing's anchor cell — splitting the repeat runs around it — with its de-gr style, and the picture and manifest entry of an inserted image", async () => {
    const model = await open();
    const sheet = model.sheets[0]!;
    applyDrawingResult(model, insertShape(model, sheet.id, "rect", { left: 10 * 96, top: 10 * 24, width: 96, height: 48 })!);
    applyDrawingResult(model, insertImage(model, model.sheets[0]!.id, { bytes: png(10, 10), contentType: "image/png", naturalPx: { width: 10, height: 10 }, at: { row: 20, col: 0 }, visiblePx: { width: 1000, height: 1000 } })!);
    applyDrawingResult(model, deleteDrawing(model, model.sheets[0]!.id, named(model.sheets[0]!, "Pic").id)!);
    const { content, parts, bytes } = await written(model);
    expect(content).toContain(
      `<table:table-row table:style-name="ro1" table:number-rows-repeated="7"><table:table-cell table:number-columns-repeated="16384"/></table:table-row>` +
        `<table:table-row table:style-name="ro1"><table:table-cell table:number-columns-repeated="10"/><table:table-cell><draw:custom-shape table:end-cell-address="Sheet1.L13"`,
    );
    expect(content).toMatch(/<\/draw:custom-shape><\/table:table-cell><table:table-cell table:number-columns-repeated="16373"\/><\/table:table-row><table:table-row table:style-name="ro1" table:number-rows-repeated="9">/);
    expect(content).toContain(`<style:style style:name="de-gr1" style:family="graphic">`);
    expect(content).toContain(`<draw:image xlink:href="Pictures/image1.png" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/>`);
    expect(equalBytes(parts.get("Pictures/image1.png")!, png(10, 10))).toBe(true);
    const manifest = decoder.decode(parts.get("META-INF/manifest.xml"));
    expect(manifest).toContain(`<manifest:file-entry manifest:full-path="Pictures/image1.png" manifest:media-type="image/png"/>`);
    // The deleted picture is shown by nothing any more: it leaves the package.
    expect(parts.has("Pictures/a.png")).toBe(false);
    expect(manifest).not.toContain("Pictures/a.png");
    const reread = (await readSpreadsheetBytes(bytes, "ods", "t2")).sheets[0]!;
    expect(reread.drawings!.map((drawing) => drawing.name).slice(-2)).toEqual(["Rectangle 1", "Picture 2"]);
  });

  it("from scratch: created drawings, their pictures, manifest entries, graphic styles and namespaces", async () => {
    const model = { ...createBlankWorkbook(), format: "ods" as const };
    const id = model.sheets[0]!.id;
    applyDrawingResult(model, insertImage(model, id, { bytes: png(4, 4), contentType: "image/png", naturalPx: { width: 4, height: 4 }, at: { row: 1, col: 1 }, visiblePx: { width: 500, height: 500 } })!);
    applyDrawingResult(model, insertShape(model, id, "ellipse", { left: 0, top: 0, width: 50, height: 50 })!);
    const { content, parts, bytes } = await written(model);
    expect(parts.has("Pictures/image1.png")).toBe(true);
    expect(decoder.decode(parts.get("META-INF/manifest.xml"))).toContain(`manifest:full-path="Pictures/image1.png" manifest:media-type="image/png"`);
    expect(content).toMatch(/xmlns:draw="[^"]+"/);
    expect(content).toMatch(/xmlns:xlink="[^"]+"/);
    expect(content).toContain(`draw:type="ellipse"`);
    expect(content).toContain(`style:family="graphic"`);
    const reread = (await readSpreadsheetBytes(bytes, "ods", "t2")).sheets[0]!;
    expect(reread.drawings!.map((drawing) => drawing.content.kind)).toEqual(["image", "shape"]);
  });
});

describe("upgradeOdsSheetDrawings", () => {
  it("lifts drawings out of preservedXml at each cell's current position, reads table:shapes, drops a cell left holding nothing", async () => {
    const model = await open();
    const sheet = model.sheets[0]!;
    // A version-10 sheet: the frames still in preservedXml, the cells since moved down a row.
    const stored: Sheet = {
      ...sheet,
      drawings: undefined,
      rows: [
        { index: 2, cells: [{ col: 0, value: null, preservedXml: PIC }] },
        { index: 3, cells: [{ col: 3, value: "x", preservedXml: `${ANNOTATION}${CHART}` }] },
      ],
    };
    const prefix = `<?xml version="1.0" encoding="UTF-8"?><office:document-content ${NS}><office:automatic-styles>${STYLES}</office:automatic-styles><office:body><office:spreadsheet>`;
    const parts = new Map([
      ["Pictures/a.png", png(96, 48)],
      ["META-INF/manifest.xml", encoder.encode(MANIFEST)],
    ]);
    const upgraded = upgradeOdsSheetDrawings({ sheet: stored, prefix, parts });
    expect(upgraded.rows).toEqual([{ index: 3, cells: [{ col: 3, value: "x", preservedXml: ANNOTATION }] }]);
    expect(upgraded.drawings.map((drawing) => [drawing.id, drawing.name, drawing.anchor.from?.row ?? null, drawing.anchor.kind])).toEqual([
      ["dr1", "Page", null, "absolute"],
      ["dr2", "", 3, "oneCell"],
      ["dr3", "Pic", 2, "oneCell"],
    ]);
    expect(upgraded.media.map((part) => part.mediaId)).toEqual(["Pictures/a.png"]);
    expect(stored.rows[0]!.cells[0]!.preservedXml).toBe(PIC);
  });
});
