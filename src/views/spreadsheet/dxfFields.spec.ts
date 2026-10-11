import { describe, expect, it } from "vitest";
import { dxfAsStyle, dxfTouchedFields, fieldsAsDxf } from "./dxfFields";
import { fieldsOf } from "../../lib/spreadsheet/formatCells";
import type { CellBorderEdge } from "../../lib/spreadsheet/model";

function noEdges(): CellBorderEdge | null {
  return null;
}

describe("dxfAsStyle / dxfTouchedFields", () => {
  it("seeds a CellStyle and the touched set from a loaded Dxf", () => {
    const dxf = { font: { color: "#9C0006", bold: true, underline: "none" as const }, fill: { kind: "solid" as const, color: "#FFC7CE" } };
    const style = dxfAsStyle(dxf);
    expect(style.font).toEqual({ color: "#9C0006", bold: true });
    expect(style.fill).toEqual({ kind: "solid", color: "#FFC7CE" });
    const touched = dxfTouchedFields(dxf);
    expect(touched.has("fontColor")).toBe(true);
    expect(touched.has("fontStyle")).toBe(true);
    expect(touched.has("underline")).toBe(true);
    expect(touched.has("strike")).toBe(false);
  });

  it("an empty Dxf seeds an empty style and no touched fields", () => {
    expect(dxfAsStyle({})).toEqual({});
    expect(dxfTouchedFields({}).size).toBe(0);
  });
});

describe("fieldsAsDxf", () => {
  it("builds a Dxf from only the touched font fields", () => {
    const fields = fieldsOf({}, "Calibri", 11);
    const touched = new Set<keyof typeof fields>(["fontColor"]);
    const dxf = fieldsAsDxf({ ...fields, fontColor: "#9C0006" }, touched, noEdges, undefined, false);
    expect(dxf).toEqual({ font: { color: "#9C0006" } });
  });

  it("round-trips a loaded Dxf unchanged when nothing is re-touched (fill/border; font IS re-seeded as touched)", () => {
    const dxf = { font: { color: "#9C0006" }, fill: { kind: "solid" as const, color: "#FFC7CE" } };
    const style = dxfAsStyle(dxf);
    const fields = fieldsOf(style, "Calibri", 11);
    const touched = dxfTouchedFields(dxf);
    const rebuilt = fieldsAsDxf(fields, touched, (edge) => (style.borders?.[edge as "top" | "right" | "bottom" | "left"] ?? null), undefined, false);
    expect(rebuilt).toEqual(dxf);
  });

  it("includes only the edges edgeShown reports set", () => {
    const edge = { style: "thin" as const, color: "#9C0006" };
    const dxf = fieldsAsDxf(fieldsOf({}, "Calibri", 11), new Set(), (side) => (side === "top" ? edge : null), undefined, false);
    expect(dxf).toEqual({ borders: { top: edge } });
  });

  it("sets an explicit false/none when a field is touched to its off state", () => {
    const fields = { ...fieldsOf({}, "Calibri", 11), fontStyle: "regular" as const, underline: "none" as const, strike: false };
    const touched = new Set<keyof typeof fields>(["fontStyle", "underline", "strike"]);
    const dxf = fieldsAsDxf(fields, touched, noEdges, undefined, false);
    expect(dxf).toEqual({ font: { bold: false, italic: false, underline: "none", strike: false } });
  });

  it("includes the number format only when touched", () => {
    const fields = fieldsOf({}, "Calibri", 11);
    expect(fieldsAsDxf(fields, new Set(), noEdges, "0.00%", false)).toEqual({});
    expect(fieldsAsDxf(fields, new Set(), noEdges, "0.00%", true)).toEqual({ numberFormat: "0.00%" });
  });
});
