import { describe, expect, it } from "vitest";
import { unzip } from "../../core/zip";
import * as F from "../fixtures/build";
import { readSpreadsheetBytes } from "../open";
import type { Sheet, SpreadsheetModel, CellStyle } from "../model";
import { cellStyle, effectiveStyleId, columnSpanAt } from "../styles";
import { writeOds } from "./writer";

function plain(style: CellStyle): CellStyle {
  const { sourceStyleName: _a, sourceXf: _b, namedStyle: _c, ...rest } = style;
  return JSON.parse(JSON.stringify(rest));
}
function facts(model: SpreadsheetModel) {
  return model.sheets.map((sheet: Sheet) => {
    const cells: Record<string, unknown> = {};
    for (const row of sheet.rows) for (const cell of row.cells) {
      const { styleId: _s, col: _c, ...rest } = cell;
      cells[`${row.index},${cell.col}`] = { ...rest, style: plain(cellStyle(model, effectiveStyleId(sheet, row.index, cell.col))) };
    }
    return { name: sheet.name, cells, rows: sheet.rows.map((r) => ({ i: r.index, h: r.heightPt, c: r.customHeight, hid: r.hidden })).filter((r) => r.h !== undefined || r.hid), columns: sheet.columns.map((c) => ({ min: c.min, max: c.max, w: c.widthPx, hid: c.hidden, style: c.styleId === undefined ? undefined : plain(cellStyle(model, c.styleId)) })), dc: sheet.defaultColWidthPx, dr: sheet.defaultRowHeightPt, merges: sheet.merges, protection: sheet.protection };
  });
}
describe("scratch", () => {
  for (const build of [F.buildMiniBookOdsBytes, F.buildSpec117StylesOdsBytes, F.buildSpec118FormatsOdsBytes, F.buildSpec119FormulasOdsBytes, F.buildSpec121LinksOdsBytes, F.buildSpec120CalcOdsBytes]) {
    it(build.name, async () => {
      const model = await readSpreadsheetBytes(await build(), "ods", "t");
      const dirty = { ...model, sheets: model.sheets.map((s) => ({ ...s, dirty: true as const })) };
      const bytes = await writeOds(dirty);
      const parts = await unzip(bytes);
      const content = new TextDecoder().decode(parts.get("content.xml"));
      const { writeFileSync } = await import("node:fs");
      writeFileSync(`/tmp/scratch/${build.name}.xml`, content);
      const back = await readSpreadsheetBytes(bytes, "ods", "t");
      expect(facts(back)).toEqual(facts(model));
    });
  }
});
