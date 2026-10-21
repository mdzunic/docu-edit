// SPEC-131 "Tests": the drawing layer's round trips. Untouched drawings come
// back byte-identical — every drawing part, its .rels, every media part and
// every ods table, `equalBytes` with the source (the SPEC-121/122 clean-part
// rule) — even when another sheet is edited, and when the drawing sheet's
// layer is marked changed without any drawing moving (the rebuilt part is
// the source's, splice by splice). Drawings created here survive save →
// reopen `equalStructural` in both formats, geometry within one pixel, and a
// second write of the reopened model is byte-identical to the first.
import { describe, expect, it } from "vitest";
import { equalBytes } from "../core/equality";
import { decodedPartBytes } from "../core/packageParts";
import { unzip } from "../core/zip";
import { firstStructuralDifference } from "../equality";
import { applyDrawingResult, drawingRectPx, insertImage, insertShape, setDrawingRect } from "./drawings";
import {
  buildSpec124StructureOdsBytes,
  buildSpec124StructureXlsxBytes,
  buildSpec131DrawingsOdsBytes,
  buildSpec131DrawingsXlsxBytes,
  SPEC131_PICTURE_PNG,
} from "./fixtures/build";
import { buildGridMetrics } from "./geometry";
import type { Sheet, SheetDrawing, SpreadsheetModel } from "./model";
import { readSpreadsheetBytes } from "./open";
import { writeOds } from "./ods/writer";
import { writeXlsx } from "./xlsx/writer";

const decoder = new TextDecoder();

const FIXTURES = [
  { name: "spec131-drawings.xlsx", format: "xlsx", build: buildSpec131DrawingsXlsxBytes },
  { name: "spec131-drawings.ods", format: "ods", build: buildSpec131DrawingsOdsBytes },
  { name: "spec124-structure.xlsx", format: "xlsx", build: buildSpec124StructureXlsxBytes },
  { name: "spec124-structure.ods", format: "ods", build: buildSpec124StructureOdsBytes },
] as const;

async function write(model: SpreadsheetModel): Promise<Uint8Array> {
  return model.format === "ods" ? writeOds(model) : writeXlsx(model);
}

/** The parts a drawing layer lives in: xlsx drawing parts, their .rels and
 * the media; ods pictures (content.xml's tables are compared on their own). */
function drawingParts(parts: Map<string, Uint8Array>): string[] {
  return [...parts.keys()].filter((name) => /^xl\/(?:drawings|media)\//.test(name) || name.startsWith("Pictures/"));
}

/** Each table's text in an ods content.xml. */
function odsTables(content: string): string[] {
  return [...content.matchAll(/<table:table [\s\S]*?<\/table:table>/g)].map((match) => match[0]);
}

/** The model with one cell on the last sheet set — another sheet edited. */
function editOtherSheet(model: SpreadsheetModel): void {
  const last = model.sheets.length - 1;
  const sheet = model.sheets[last]!;
  model.sheets = model.sheets.map((each, index) =>
    index === last ? { ...sheet, dirty: true, rows: [...sheet.rows.filter((row) => row.index !== 60), { index: 60, cells: [{ col: 0, value: "edited" }] }] } : each,
  );
}

/** The first sheet holding drawings. */
const drawingSheet = (model: SpreadsheetModel): Sheet => model.sheets.find((sheet) => (sheet.drawings ?? []).length > 0)!;

describe("untouched drawings round-trip byte-identical", () => {
  for (const fixture of FIXTURES) {
    it(`${fixture.name}: every drawing part, .rels, media part and ods table equalBytes the source when another sheet is edited`, async () => {
      const bytes = await fixture.build();
      const source = await unzip(bytes);
      const model = await readSpreadsheetBytes(bytes, fixture.format, fixture.name);
      expect(drawingSheet(model)).toBeDefined();
      editOtherSheet(model);
      const written = await unzip(await write(model));
      for (const name of drawingParts(source)) {
        expect(written.has(name), name).toBe(true);
        expect(equalBytes(written.get(name)!, source.get(name)!), name).toBe(true);
      }
      if (fixture.format === "ods") {
        const before = odsTables(decoder.decode(source.get("content.xml")));
        const after = odsTables(decoder.decode(written.get("content.xml")));
        const name = drawingSheet(model).name;
        const index = model.sheets.findIndex((sheet) => sheet.name === name);
        expect(after[index]).toBe(before[index]);
      }
    });

    it(`${fixture.name}: a layer marked changed with no drawing moved writes its drawings' bytes as they were`, async () => {
      const bytes = await fixture.build();
      const source = await unzip(bytes);
      const model = await readSpreadsheetBytes(bytes, fixture.format, fixture.name);
      const sheet = drawingSheet(model);
      model.sheets = model.sheets.map((each) => (each === sheet ? { ...each, ...(fixture.format === "xlsx" ? { drawingsDirty: true as const } : { dirty: true as const }) } : each));
      const written = await unzip(await write(model));
      if (fixture.format === "xlsx") {
        for (const name of drawingParts(source)) {
          expect(equalBytes(written.get(name)!, source.get(name)!), name).toBe(true);
        }
      } else {
        // The dirty table is re-serialised; each imported drawing's element is its source, verbatim.
        const content = decoder.decode(written.get("content.xml"));
        for (const drawing of sheet.drawings!) {
          expect(content, drawing.name).toContain(drawing.sourceXml!);
        }
        expect(equalBytes(written.get(drawingParts(source)[0]!) ?? new Uint8Array(), source.get(drawingParts(source)[0]!) ?? new Uint8Array())).toBe(true);
      }
    });
  }
});

/** A drawing without its minted id, for comparing across a reopen. */
function comparable(drawing: SheetDrawing): unknown {
  const { id: _id, sourceXml: _source, anchor: _anchor, ...rest } = drawing;
  return rest;
}

describe("created drawings survive save → reopen", () => {
  for (const fixture of FIXTURES.slice(0, 2)) {
    it(`${fixture.name}: an inserted image and shape come back equalStructural, geometry within a pixel; a second write is byte-stable`, async () => {
      const model = await readSpreadsheetBytes(await fixture.build(), fixture.format, fixture.name);
      const sheet = drawingSheet(model);
      const metrics = buildGridMetrics(sheet);
      applyDrawingResult(
        model,
        insertImage(model, sheet.id, { bytes: SPEC131_PICTURE_PNG, contentType: "image/png", naturalPx: { width: 96, height: 64 }, at: { row: 24, col: 3 }, visiblePx: { width: 1000, height: 700 } }, metrics)!,
      );
      applyDrawingResult(model, insertShape(model, sheet.id, "ellipse", { left: 400, top: 520, width: 90, height: 45 }, metrics)!);
      // An imported drawing moved, too.
      const first = drawingSheet(model).drawings![0]!;
      const moved = drawingRectPx(first.anchor, metrics);
      applyDrawingResult(model, setDrawingRect(model, sheet.id, first.id, { ...moved, left: moved.left + 30, top: moved.top + 25 }, metrics)!);

      const before = drawingSheet(model);
      const written = await write(model);
      const reopened = await readSpreadsheetBytes(written, fixture.format, fixture.name);
      const after = drawingSheet(reopened);
      expect(after.drawings!.map((drawing) => drawing.name)).toEqual(before.drawings!.map((drawing) => drawing.name));
      const created = before.drawings!.slice(-2);
      const back = after.drawings!.slice(-2);
      created.forEach((drawing, index) => {
        expect(firstStructuralDifference(comparable(back[index]!), comparable(drawing)), drawing.name).toBeNull();
      });
      const beforeMetrics = buildGridMetrics(before);
      const afterMetrics = buildGridMetrics(after);
      before.drawings!.forEach((drawing, index) => {
        const want = drawingRectPx(drawing.anchor, beforeMetrics);
        const got = drawingRectPx(after.drawings![index]!.anchor, afterMetrics);
        for (const key of ["left", "top", "width", "height"] as const) {
          expect(Math.abs(got[key] - want[key]), `${drawing.name} ${key}`).toBeLessThanOrEqual(1);
        }
      });
      // The created image's bytes are its part.
      const image = back[0]!;
      expect(image.content.kind).toBe("image");
      const parts = await unzip(written);
      expect(equalBytes(parts.get(image.content.kind === "image" ? image.content.mediaId : "")!, SPEC131_PICTURE_PNG)).toBe(true);

      // A second write of the reopened model is byte-identical.
      const again = await write(reopened);
      const second = await unzip(again);
      for (const [name, data] of parts) {
        expect(equalBytes(second.get(name)!, data), name).toBe(true);
      }
      expect([...second.keys()]).toEqual([...parts.keys()]);
      expect(decodedPartBytes(reopened.opaque).size).toBe(parts.size);
    });
  }
});
