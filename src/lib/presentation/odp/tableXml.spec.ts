import { describe, expect, it } from "vitest";
import { serializeOdpTableFrame } from "./serialize";
import { parseOdpTableFrameForTest } from "./reader";
import type { TableFrame } from "../model";
import { GOLDEN_TABLE, TABLE_GOLDEN, TABLE_GOLDEN_STYLES } from "./tableGolden";

const NS =
  `xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" ` +
  `xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" ` +
  `xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" ` +
  `xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"`;

function styleFor(name: string, family: string, props: string): string {
  return `<style:style style:name="${name}" style:family="${family}">${props}</style:style>`;
}
const COL = (name: string, w: string) => styleFor(name, "table-column", `<style:table-column-properties style:column-width="${w}"/>`);

describe("odp table golden (AC-16)", () => {
  it("serialises GOLDEN_TABLE to TABLE_GOLDEN + styles exactly (write direction)", () => {
    const frame = { ...GOLDEN_TABLE, id: "any" } as TableFrame;
    const out = serializeOdpTableFrame(frame);
    expect(out.frame).toBe(TABLE_GOLDEN);
    expect(out.autoStyles).toBe(TABLE_GOLDEN_STYLES);
  });

  it("reads TABLE_GOLDEN back to the clean GOLDEN_TABLE (read direction, AC-16b)", () => {
    const parsed = parseOdpTableFrameForTest(TABLE_GOLDEN, TABLE_GOLDEN_STYLES);
    expect(parsed).toBeDefined();
    const { id: _id, ...rest } = parsed!;
    expect(rest).toEqual(GOLDEN_TABLE);
  });
});

describe("odp table reader accepts flavoured variants (AC-16c)", () => {
  it("accepts number-columns-repeated, covered cells and a spanning cell", () => {
    const styles =
      COL("c", "3cm") +
      styleFor("r", "table-row", `<style:table-row-properties style:row-height="1cm"/>`);
    const frame =
      `<draw:frame ${NS} svg:x="0cm" svg:y="0cm" svg:width="6cm" svg:height="2cm">` +
      `<table:table>` +
      `<table:table-column table:style-name="c" table:number-columns-repeated="2"/>` +
      `<table:table-header-rows>` +
      `<table:table-row table:style-name="r">` +
      `<table:table-cell table:number-columns-spanned="2"><text:p>Head</text:p></table:table-cell>` +
      `<table:covered-table-cell/>` +
      `</table:table-row></table:table-header-rows>` +
      `<table:table-row table:style-name="r">` +
      `<table:table-cell><text:p>a</text:p></table:table-cell>` +
      `<table:table-cell><text:p>b</text:p></table:table-cell>` +
      `</table:table-row></table:table></draw:frame>`;
    const parsed = parseOdpTableFrameForTest(frame, styles);
    expect(parsed).toBeDefined();
    expect(parsed!.columnWidthsCm).toEqual([3, 3]);
    expect(parsed!.rows[0].cells[0].colSpan).toBe(2);
    expect(parsed!.rows[0].cells[1].covered).toBe(true);
    expect(parsed!.rows[0].isHeader).toBe(true);
    expect(parsed!.firstRow).toBe(true);
    expect(parsed!.rows[1].cells[1].paragraphs[0].runs[0].text).toBe("b");
  });
});

describe("odp table abort conditions map to undefined (AC-16d)", () => {
  const c = `<table:table-cell><text:p>x</text:p></table:table-cell>`;
  it("tolerates missing table-column declarations by inferring the count", () => {
    const frame = `<draw:frame ${NS} svg:x="0cm" svg:y="0cm" svg:width="6cm" svg:height="2cm"><table:table><table:table-row>${c}${c}</table:table-row></table:table></draw:frame>`;
    const parsed = parseOdpTableFrameForTest(frame);
    expect(parsed).toBeDefined();
    expect(parsed!.columnWidthsCm).toHaveLength(2);
  });
  it("aborts with no rows", () => {
    const frame = `<draw:frame ${NS} svg:x="0cm" svg:y="0cm" svg:width="6cm" svg:height="2cm"><table:table><table:table-column/></table:table></draw:frame>`;
    expect(parseOdpTableFrameForTest(frame)).toBeUndefined();
  });
  it("aborts when a row's cell count does not match the column count", () => {
    const frame = `<draw:frame ${NS} svg:x="0cm" svg:y="0cm" svg:width="6cm" svg:height="2cm"><table:table><table:table-column/><table:table-column/><table:table-row>${c}${c}${c}</table:table-row></table:table></draw:frame>`;
    expect(parseOdpTableFrameForTest(frame)).toBeUndefined();
  });
  it("aborts on an inconsistent merge map", () => {
    const frame = `<draw:frame ${NS} svg:x="0cm" svg:y="0cm" svg:width="6cm" svg:height="2cm"><table:table><table:table-column/><table:table-column/><table:table-row><table:table-cell table:number-columns-spanned="2"><text:p>o</text:p></table:table-cell>${c}</table:table-row></table:table></draw:frame>`;
    expect(parseOdpTableFrameForTest(frame)).toBeUndefined();
  });
});
