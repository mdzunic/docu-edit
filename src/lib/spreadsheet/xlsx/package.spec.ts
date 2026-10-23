// SPEC-137: worksheetTextWithoutCells — a retained worksheet's text with its
// cells cut out, decoded from each end of the part's base64, equals the
// whole text with the same cut (UTF-8 across the slice boundaries, a
// prefixed namespace, a part too small to slice, and the fallbacks).
import { describe, expect, it } from "vitest";
import { bytesToBase64 } from "../../core/packageParts";
import { worksheetTextWithoutCells } from "./package";

const encoder = new TextEncoder();

/** A worksheet whose `<sheetData>` holds `rows` rows of multi-byte text. */
function worksheet(rows: number, prefix = "", tail = `<${prefix}hyperlinks><${prefix}hyperlink ref="A1" r:id="rId1"/></${prefix}hyperlinks><${prefix}legacyDrawing r:id="rId2"/>`): string {
  const p = prefix;
  const cells: string[] = [];
  for (let r = 1; r <= rows; r += 1) {
    cells.push(`<${p}row r="${r}"><${p}c r="A${r}" t="inlineStr"><${p}is><${p}t>Zürich ✓ 東京 ${r}</${p}t></${p}is></${p}c></${p}row>`);
  }
  const ns = p === "" ? `xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"` : `xmlns:${p.slice(0, -1)}="http://schemas.openxmlformats.org/spreadsheetml/2006/main"`;
  return `<?xml version="1.0" encoding="UTF-8"?><${p}worksheet ${ns} xmlns:r="r"><${p}sheetViews><${p}sheetView workbookViewId="0"/></${p}sheetViews><${p}sheetData>${cells.join("")}</${p}sheetData>${tail}</${p}worksheet>`;
}

/** The text with `<sheetData>`'s content removed, cut the slow way. */
function cut(xml: string, prefix = ""): string {
  const open = xml.indexOf(">", xml.indexOf(`<${prefix}sheetData`)) + 1;
  return xml.slice(0, open) + xml.slice(xml.lastIndexOf(`</${prefix}sheetData`));
}

const base64 = (xml: string): string => bytesToBase64(encoder.encode(xml));

describe("worksheetTextWithoutCells (SPEC-137)", () => {
  it("decodes a small part whole", () => {
    const xml = worksheet(3);
    expect(worksheetTextWithoutCells(base64(xml))).toBe(xml);
  });

  it("cuts a large part's cells out from its two ends, as a whole decode would", () => {
    for (const rows of [6_000, 6_001, 6_002, 40_000]) {
      const xml = worksheet(rows);
      expect(worksheetTextWithoutCells(base64(xml))).toBe(cut(xml));
    }
  });

  it("finds a prefixed sheetData and a long tail", () => {
    const xml = worksheet(8_000, "x:", `<x:mergeCells>${'<x:mergeCell ref="B2:C3"/>'.repeat(20_000)}</x:mergeCells>`);
    expect(worksheetTextWithoutCells(base64(xml))).toBe(cut(xml, "x:"));
  });

  it("decodes whole a large part whose cut does not show at its ends", () => {
    const empty = `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/>${"<!-- padding ✓ -->".repeat(30_000)}</worksheet>`;
    expect(worksheetTextWithoutCells(base64(empty))).toBe(empty);
  });
});
