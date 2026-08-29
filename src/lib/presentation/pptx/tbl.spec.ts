import { describe, expect, it } from "vitest";
import { serializeTableFrame } from "./serialize";
import { parseTableGraphicFrameForTest } from "./reader";
import type { TableFrame } from "../model";
import { GOLDEN_TABLE, TBL_GOLDEN } from "./tblGolden";

/** Wraps a bare `a:tbl` in the minimal graphicFrame the reader expects. */
function graphicFrame(tbl: string, xfrm = `<p:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></p:xfrm>`): string {
  return (
    `<p:graphicFrame xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" ` +
    `xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
    `<p:nvGraphicFramePr><p:cNvPr id="5" name="Table 5"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>` +
    `${xfrm}<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">${tbl}</a:graphicData></a:graphic></p:graphicFrame>`
  );
}

describe("pptx table golden (AC-16)", () => {
  it("serialises GOLDEN_TABLE to TBL_GOLDEN exactly (write direction)", () => {
    const frame = { ...GOLDEN_TABLE, id: "any-id" } as TableFrame;
    expect(serializeTableFrame(frame, 5)).toBe(TBL_GOLDEN);
  });

  it("reads TBL_GOLDEN back to the clean GOLDEN_TABLE (read direction, AC-16b)", () => {
    const parsed = parseTableGraphicFrameForTest(TBL_GOLDEN);
    expect(parsed).toBeDefined();
    const { id: _id, ...rest } = parsed!;
    expect(rest).toEqual(GOLDEN_TABLE);
  });

  it("round-trips serialiser output back to identical bytes (stability)", () => {
    const parsed = parseTableGraphicFrameForTest(TBL_GOLDEN)!;
    expect(serializeTableFrame(parsed, 5)).toBe(TBL_GOLDEN);
  });
});

describe("pptx table reader accepts flavoured variants (AC-16c)", () => {
  it("accepts hMerge/vMerge as \"true\", reordered attributes and self-closing tcPr", () => {
    const tbl =
      `<a:tbl><a:tblPr firstRow="1"/>` +
      `<a:tblGrid><a:gridCol w="914400"/><a:gridCol w="914400"/></a:tblGrid>` +
      // Row 0: a 2-column merge (owner gridSpan=2, continuation hMerge="true").
      `<a:tr h="360000">` +
      `<a:tc gridSpan="2"><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Header</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>` +
      `<a:tc hMerge="true"><a:tcPr/></a:tc>` +
      `</a:tr>` +
      // Row 1: two plain cells, attributes reordered on the second.
      `<a:tr h="360000">` +
      `<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>A</a:t></a:r></a:p></a:txBody></a:tc>` +
      `<a:tc><a:txBody><a:lstStyle/><a:bodyPr/><a:p><a:r><a:t>B</a:t></a:r></a:p></a:txBody></a:tc>` +
      `</a:tr></a:tbl>`;
    const parsed = parseTableGraphicFrameForTest(graphicFrame(tbl));
    expect(parsed).toBeDefined();
    expect(parsed!.columnWidthsCm).toHaveLength(2);
    expect(parsed!.rows[0].cells[0].colSpan).toBe(2);
    expect(parsed!.rows[0].cells[1].covered).toBe(true);
    expect(parsed!.rows[0].cells[0].paragraphs[0].runs[0].text).toBe("Header");
    expect(parsed!.rows[1].cells[1].paragraphs[0].runs[0].text).toBe("B");
    expect(parsed!.firstRow).toBe(true);
  });

  it("reads a vertical merge (rowSpan + vMerge continuation)", () => {
    const tbl =
      `<a:tbl>` +
      `<a:tblGrid><a:gridCol w="914400"/><a:gridCol w="914400"/></a:tblGrid>` +
      `<a:tr h="360000">` +
      `<a:tc rowSpan="2"><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Tall</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>` +
      `<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>x</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>` +
      `</a:tr>` +
      `<a:tr h="360000">` +
      `<a:tc vMerge="1"><a:tcPr/></a:tc>` +
      `<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>y</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>` +
      `</a:tr></a:tbl>`;
    const parsed = parseTableGraphicFrameForTest(graphicFrame(tbl));
    expect(parsed).toBeDefined();
    expect(parsed!.rows[0].cells[0].rowSpan).toBe(2);
    expect(parsed!.rows[1].cells[0].covered).toBe(true);
  });
});

describe("pptx table abort conditions map to undefined (AC-16d)", () => {
  const cell = `<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>x</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>`;

  it("aborts when the grid has no columns", () => {
    const tbl = `<a:tbl><a:tblGrid/><a:tr>${cell}</a:tr></a:tbl>`;
    expect(parseTableGraphicFrameForTest(graphicFrame(tbl))).toBeUndefined();
  });

  it("aborts when there are no rows", () => {
    const tbl = `<a:tbl><a:tblGrid><a:gridCol w="914400"/></a:tblGrid></a:tbl>`;
    expect(parseTableGraphicFrameForTest(graphicFrame(tbl))).toBeUndefined();
  });

  it("aborts when a row's cell count does not match the column count", () => {
    const tbl =
      `<a:tbl><a:tblGrid><a:gridCol w="914400"/><a:gridCol w="914400"/></a:tblGrid>` +
      `<a:tr>${cell}${cell}${cell}</a:tr></a:tbl>`; // 3 cells, 2 columns
    expect(parseTableGraphicFrameForTest(graphicFrame(tbl))).toBeUndefined();
  });

  it("aborts on an inconsistent merge map (gridSpan without a covered neighbour)", () => {
    const tbl =
      `<a:tbl><a:tblGrid><a:gridCol w="914400"/><a:gridCol w="914400"/></a:tblGrid>` +
      // owner claims gridSpan=2 but its neighbour is a real (non-merged) cell.
      `<a:tr><a:tc gridSpan="2"><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>o</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>${cell}</a:tr></a:tbl>`;
    expect(parseTableGraphicFrameForTest(graphicFrame(tbl))).toBeUndefined();
  });
});
