// SPEC-132 AC-3, AC-4, AC-16, AC-35, AC-47: the xlsx chart part on the wire —
// references read from Excel-shaped parts, the in-place patch (c:f texts and
// caches change, every other byte stays), and the goldens.
import { describe, expect, it } from "vitest";
import { CHART_CONTENT_TYPE, CHART_RELATIONSHIP, parseChartXml } from "../../core/chartPartOoxml";
import { MULTI_LEVEL_CHART_REASON, PIVOT_CHART_REASON, unsupportedChartVisual, type ChartCell } from "../charts";
import type { CellValue, SheetChartRefs } from "../model";
import {
  GOLDEN_CHART_DRAWING_RELATIONSHIP,
  GOLDEN_CHART_GRAPHIC_FRAME,
  GOLDEN_CHART_GRAPHIC_FRAME_INPUT,
  GOLDEN_CHART_OVERRIDE,
  GOLDEN_COLUMN_CHART_PART,
  GOLDEN_COLUMN_CHART_REFS,
  GOLDEN_COLUMN_CHART_VISUAL,
  GOLDEN_LINE_CHART_PART,
  GOLDEN_LINE_CHART_REFS,
  GOLDEN_LINE_CHART_VISUAL,
  GOLDEN_SCATTER_CHART_PART,
  GOLDEN_SCATTER_CHART_REFS,
  GOLDEN_SCATTER_CHART_VISUAL,
} from "./chartGolden";
import {
  chartGraphicFrameXml,
  chartReadOnlyReason,
  nextChartPartName,
  patchXlsxChartPart,
  readChartPartRefs,
  readXlsxChart,
  xlsxChartPartXml,
  xlsxChartWireRefs,
} from "./chartXml";

// --- Excel-shaped parts ----------------------------------------------------------------------

const C = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

const chartSpace = (inner: string, prefix = "c"): string =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n` +
  `<${prefix}:chartSpace xmlns:${prefix}="${C}" xmlns:a="${A}" xmlns:r="${R}">` +
  `<${prefix}:date1904 val="0"/><${prefix}:lang val="en-US"/><${prefix}:roundedCorners val="0"/>` +
  `<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">` +
  `<mc:Choice Requires="c14" xmlns:c14="http://schemas.microsoft.com/office/drawing/2007/8/2/chart"><c14:style val="102"/></mc:Choice>` +
  `<mc:Fallback><${prefix}:style val="2"/></mc:Fallback></mc:AlternateContent>` +
  inner +
  `<${prefix}:printSettings><${prefix}:headerFooter/><${prefix}:pageMargins b="0.75" l="0.7" r="0.7" t="0.75" header="0.3" footer="0.3"/>` +
  `<${prefix}:pageSetup/></${prefix}:printSettings></${prefix}:chartSpace>`;

const strCache = (labels: readonly string[]): string =>
  `<c:strCache><c:ptCount val="${labels.length}"/>${labels.map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join("")}</c:strCache>`;
const numCache = (values: readonly (number | null)[], formatCode = "General", ext = ""): string =>
  `<c:numCache><c:formatCode>${formatCode}</c:formatCode><c:ptCount val="${values.length}"/>` +
  values.map((v, i) => (v === null ? "" : `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`)).join("") +
  `${ext}</c:numCache>`;

const SERIES_EXT =
  `<c:extLst><c:ext uri="{02D57815-91ED-43cb-92C2-25804820EDAC}" xmlns:c15="http://schemas.microsoft.com/office/drawing/2012/chart">` +
  `<c15:datalabelsRange><c15:f>Data!$E$2:$E$5</c15:f><c15:dlblRangeCache><c:ptCount val="4"/></c15:dlblRangeCache></c15:datalabelsRange></c:ext>` +
  `<c:ext uri="{C3380CC4-5D6E-409C-BE32-E72D297353CC}" xmlns:c16="http://schemas.microsoft.com/office/drawing/2014/chart">` +
  `<c16:uniqueId val="{00000000-0001-0000-0000-000000000000}"/></c:ext></c:extLst>`;
const CACHE_EXT = `<c:extLst><c:ext uri="{9C2A4B0E-0000-4000-8000-000000000001}"><x:keep xmlns:x="urn:example:keep"/></c:ext></c:extLst>`;
const MONTHS = ["Jan", "Feb", "Mar", "Apr"];

function barSeries(index: number, col: string, name: string, values: readonly (number | null)[], formatCode = "General", ext = ""): string {
  return (
    `<c:ser><c:idx val="${index}"/><c:order val="${index}"/>` +
    `<c:tx><c:strRef><c:f>Data!$${col}$1</c:f>${strCache([name])}</c:strRef></c:tx>` +
    `<c:spPr><a:solidFill><a:schemeClr val="accent${index + 1}"/></a:solidFill><a:ln><a:noFill/></a:ln></c:spPr><c:invertIfNegative val="0"/>` +
    `<c:cat><c:strRef><c:f>Data!$A$2:$A$5</c:f>${strCache(MONTHS)}</c:strRef></c:cat>` +
    `<c:val><c:numRef><c:f>Data!$${col}$2:$${col}$5</c:f>${numCache(values, formatCode, ext)}</c:numRef></c:val>` +
    SERIES_EXT.replace("$E$", `$${col}$`).replace("0001-0000", `000${index + 1}-0000`) +
    `</c:ser>`
  );
}

const AXES =
  `<c:catAx><c:axId val="10"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/>` +
  `<c:numFmt formatCode="General" sourceLinked="1"/><c:crossAx val="20"/><c:crosses val="autoZero"/><c:auto val="1"/></c:catAx>` +
  `<c:valAx><c:axId val="20"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/>` +
  `<c:majorGridlines/><c:crossAx val="10"/><c:crosses val="autoZero"/></c:valAx>`;

function plot(typeXml: string, tail = `<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/>`): string {
  return `<c:chart><c:autoTitleDeleted val="0"/><c:plotArea><c:layout/>${typeXml}${AXES}</c:plotArea><c:legend><c:legendPos val="b"/><c:overlay val="0"/></c:legend>${tail}</c:chart>`;
}

/** A clustered column over Data!$A$1:$D$5 as Excel writes it. */
const COLUMN_PART = chartSpace(
  plot(
    `<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>` +
      barSeries(0, "B", "North", [10, 20, 30, 40]) +
      barSeries(1, "C", "South", [4, null, 6, 8], "0.0", CACHE_EXT) +
      barSeries(2, "D", "West", [1, 2, 3, 4]) +
      `<c:dLbls><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>` +
      `<c:gapWidth val="219"/><c:overlap val="-27"/><c:axId val="10"/><c:axId val="20"/></c:barChart>`,
  ),
);

const COLUMN_REFS: SheetChartRefs = {
  series: ["B", "C", "D"].map((col) => ({ name: `Data!$${col}$1`, categories: "Data!$A$2:$A$5", values: `Data!$${col}$2:$${col}$5` })),
  plotVisibleOnly: true,
};

/** A line over another sheet, 'Q1 Data'!$B$2:$B$13, named by a literal. */
function linePart(tail?: string): string {
  return chartSpace(
    plot(
      `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>` +
        `<c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:v>Revenue</c:v></c:tx><c:marker><c:symbol val="none"/></c:marker>` +
        `<c:cat><c:numRef><c:f>'Q1 Data'!$A$2:$A$13</c:f>${numCache([1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12])}</c:numRef></c:cat>` +
        `<c:val><c:numRef><c:f>'Q1 Data'!$B$2:$B$13</c:f>${numCache([5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16])}</c:numRef></c:val>` +
        `<c:smooth val="0"/></c:ser><c:marker val="1"/><c:axId val="10"/><c:axId val="20"/></c:lineChart>`,
      tail,
    ),
  );
}

const SCATTER_PART = chartSpace(
  plot(
    `<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>` +
      `<c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:strRef><c:f>Data!$B$1</c:f>${strCache(["y"])}</c:strRef></c:tx>` +
      `<c:xVal><c:numRef><c:f>Data!$A$2:$A$4</c:f>${numCache([1, 2, 4])}</c:numRef></c:xVal>` +
      `<c:yVal><c:numRef><c:f>Data!$B$2:$B$4</c:f>${numCache([1.5, 3, 7.25])}</c:numRef></c:yVal>` +
      `<c:smooth val="0"/></c:ser><c:axId val="10"/><c:axId val="20"/></c:scatterChart>`,
  ).replace("<c:catAx>", "<c:valAx>").replace("<c:auto val=\"1\"/></c:catAx>", "</c:valAx>"),
);

const cell = (value: CellValue, text: string = value === null ? "" : String(value), row = 0): ChartCell => ({ sheetId: "s1", row, col: 0, value, text, hidden: false });

// --- reading ---------------------------------------------------------------------------------

describe("readChartPartRefs — Excel-shaped parts", () => {
  it("reads every series' name, categories and values references, skipping c15 extension formulas", () => {
    expect(readChartPartRefs(COLUMN_PART)).toEqual({ refs: COLUMN_REFS, uncached: false });
  });

  it("reads quoted sheet names and leaves a literal c:tx/c:v name absent", () => {
    expect(readChartPartRefs(linePart())).toEqual({
      refs: { series: [{ categories: "'Q1 Data'!$A$2:$A$13", values: "'Q1 Data'!$B$2:$B$13" }], plotVisibleOnly: true },
      uncached: false,
    });
  });

  it("reads a scatter's c:xVal and c:yVal", () => {
    expect(readChartPartRefs(SCATTER_PART)?.refs.series).toEqual([{ name: "Data!$B$1", categories: "Data!$A$2:$A$4", values: "Data!$B$2:$B$4" }]);
  });

  it("reads plotVisOnly: 0 and false are false; absent, valueless or 1 are true", () => {
    const plotVisibleOnly = (tail: string): boolean | undefined => readChartPartRefs(linePart(tail))?.refs.plotVisibleOnly;
    expect(plotVisibleOnly(`<c:plotVisOnly val="0"/>`)).toBe(false);
    expect(plotVisibleOnly(`<c:plotVisOnly val="false"/><c:dispBlanksAs val="zero"/>`)).toBe(false);
    expect(plotVisibleOnly(`<c:plotVisOnly val="1"/>`)).toBe(true);
    expect(plotVisibleOnly(`<c:plotVisOnly/>`)).toBe(true);
    expect(plotVisibleOnly(`<c:dispBlanksAs val="gap"/>`)).toBe(true);
  });

  it("resolves the chart namespace from the declarations, whatever the prefix", () => {
    const renamed = COLUMN_PART.replaceAll("<c:", "<ch:").replaceAll("</c:", "</ch:").replace(`xmlns:c="${C}"`, `xmlns:ch="${C}"`);
    expect(renamed).not.toContain("<c:");
    expect(readChartPartRefs(renamed)).toEqual({ refs: COLUMN_REFS, uncached: false });
    // A `c:` prefix bound to another namespace is not the chart's.
    expect(readChartPartRefs(COLUMN_PART.replace(`xmlns:c="${C}"`, `xmlns:c="urn:not-a-chart"`))).toBeUndefined();
  });

  it("decodes entities before reading a reference", () => {
    const part = SCATTER_PART.replace("<c:f>Data!$B$2:$B$4</c:f>", "<c:f>'R&amp;D'!$B$2:$B$4</c:f>");
    expect(readChartPartRefs(part)?.refs.series[0]!.values).toBe("'R&D'!$B$2:$B$4");
  });

  it("flags a reference with no cache beside it", () => {
    const part = COLUMN_PART.replace(`<c:f>Data!$D$1</c:f>${strCache(["West"])}`, "<c:f>Data!$D$1</c:f>");
    expect(readChartPartRefs(part)).toEqual({ refs: COLUMN_REFS, uncached: true });
  });

  it("is undefined for a part naming no reference, a combo, or malformed XML", () => {
    const literal = chartSpace(
      plot(
        `<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:v>Lit</c:v></c:tx>` +
          `<c:cat><c:strLit><c:ptCount val="1"/><c:pt idx="0"><c:v>a</c:v></c:pt></c:strLit></c:cat>` +
          `<c:val><c:numLit><c:ptCount val="1"/><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val></c:ser><c:axId val="10"/><c:axId val="20"/></c:barChart>`,
      ),
    );
    expect(readChartPartRefs(literal)).toBeUndefined();
    expect(readChartPartRefs(COLUMN_PART.replace("</c:barChart>", "</c:barChart><c:lineChart/>"))).toBeUndefined();
    expect(readChartPartRefs("<c:chartSpace")).toBeUndefined();
  });
});

describe("chartReadOnlyReason", () => {
  const pivot = COLUMN_PART.replace("<c:chart>", "<c:pivotSource><c:name>[Book1.xlsx]Pivot!PivotTable1</c:name><c:fmtId val=\"0\"/></c:pivotSource><c:chart>");
  const multiLevel = COLUMN_PART.replace(
    `<c:cat><c:strRef><c:f>Data!$A$2:$A$5</c:f>${strCache(MONTHS)}</c:strRef></c:cat>`,
    `<c:cat><c:multiLvlStrRef><c:f>Data!$A$2:$B$5</c:f><c:multiLvlStrCache><c:ptCount val="4"/>` +
      `<c:lvl>${MONTHS.map((m, i) => `<c:pt idx="${i}"><c:v>${m}</c:v></c:pt>`).join("")}</c:lvl>` +
      `<c:lvl><c:pt idx="0"><c:v>H1</c:v></c:pt><c:pt idx="2"><c:v>H2</c:v></c:pt></c:lvl></c:multiLvlStrCache></c:multiLvlStrRef></c:cat>`,
  );

  it("names pivot charts, multi-level categories, and nothing else", () => {
    expect(chartReadOnlyReason(pivot)).toBe(PIVOT_CHART_REASON);
    expect(chartReadOnlyReason(multiLevel)).toBe(MULTI_LEVEL_CHART_REASON);
    expect(chartReadOnlyReason(pivot.replace(COLUMN_PART, multiLevel))).toBe(PIVOT_CHART_REASON);
    expect(chartReadOnlyReason(pivot.replace(/<c:cat>.*?<\/c:cat>/, multiLevel.match(/<c:cat><c:multiLvlStrRef>.*?<\/c:cat>/)![0]))).toBe(PIVOT_CHART_REASON);
    expect(chartReadOnlyReason(COLUMN_PART)).toBeUndefined();
  });

  it("reads a multi-level chart's references and paints its innermost level from the cache", () => {
    const read = readXlsxChart(multiLevel);
    expect(read.readOnlyReason).toBe(MULTI_LEVEL_CHART_REASON);
    expect(read.refs?.series[0]!.categories).toBe("Data!$A$2:$B$5");
    expect(read.visual.categories).toEqual(MONTHS);
    expect(readXlsxChart(pivot).readOnlyReason).toBe(PIVOT_CHART_REASON);
  });
});

describe("readXlsxChart", () => {
  it("parses a catalogue chart's visual from its caches with its references", () => {
    const read = readXlsxChart(COLUMN_PART);
    expect(read.visual.chartType).toBe("columnClustered");
    expect(read.visual.categories).toEqual(MONTHS);
    expect(read.visual.series.map((s) => [s.name, s.values])).toEqual([
      ["North", [10, 20, 30, 40]],
      ["South", [4, null, 6, 8]],
      ["West", [1, 2, 3, 4]],
    ]);
    expect(read.refs).toEqual(COLUMN_REFS);
    expect(read.uncached).toBeUndefined();
    expect(read.readOnlyReason).toBeUndefined();
  });

  it("keeps an out-of-catalogue chart (c:bar3DChart) as its part text: placeholder visual, no refs", () => {
    const bar3D = COLUMN_PART.replace("<c:barChart>", "<c:view3D/><c:bar3DChart>").replace("</c:barChart>", "</c:bar3DChart>");
    expect(readXlsxChart(bar3D)).toEqual({ visual: unsupportedChartVisual(bar3D) });
  });

  it("numbers the categories of a chart with references but no categories reference", () => {
    const part = COLUMN_PART.replaceAll(`<c:cat><c:strRef><c:f>Data!$A$2:$A$5</c:f>${strCache(MONTHS)}</c:strRef></c:cat>`, "");
    const read = readXlsxChart(part);
    expect(read.visual.categories).toEqual(["1", "2", "3", "4"]);
    expect(read.refs?.series.every((s) => s.categories === undefined)).toBe(true);
  });

  it("marks a chart with an uncached reference for a display refresh", () => {
    expect(readXlsxChart(COLUMN_PART.replace(numCache([1, 2, 3, 4]), "")).uncached).toBe(true);
  });
});

// --- patching ----------------------------------------------------------------------------------

/** The refs after a row inserted above the data (rows 1–5 → 2–6). */
const MOVED_REFS: SheetChartRefs = {
  series: ["B", "C", "D"].map((col) => ({ name: `Data!$${col}$2`, categories: "Data!$A$3:$A$6", values: `Data!$${col}$3:$${col}$6` })),
  plotVisibleOnly: true,
};

/** COLUMN_PART with exactly its modelled c:f texts moved down one row. */
function columnPartMoved(): string {
  let expected = COLUMN_PART.replaceAll("<c:f>Data!$A$2:$A$5</c:f>", "<c:f>Data!$A$3:$A$6</c:f>");
  for (const col of ["B", "C", "D"]) {
    expected = expected.replace(`<c:f>Data!$${col}$1</c:f>`, `<c:f>Data!$${col}$2</c:f>`).replace(`<c:f>Data!$${col}$2:$${col}$5</c:f>`, `<c:f>Data!$${col}$3:$${col}$6</c:f>`);
  }
  return expected;
}

/** A grid for cellsOf: a reference → its cells. */
function cellsFrom(table: Record<string, ChartCell[]>): (canonical: string) => ChartCell[] | null {
  return (canonical) => table[canonical] ?? null;
}

/** The cells COLUMN_PART's caches already hold. */
const SAME_CELLS: Record<string, ChartCell[]> = {
  "Data!$A$2:$A$5": MONTHS.map((m) => cell(m)),
  "Data!$B$1": [cell("North")],
  "Data!$C$1": [cell("South")],
  "Data!$D$1": [cell("West")],
  "Data!$B$2:$B$5": [10, 20, 30, 40].map((v) => cell(v)),
  "Data!$C$2:$C$5": [cell(4), cell(null), cell(6), cell(8)],
  "Data!$D$2:$D$5": [1, 2, 3, 4].map((v) => cell(v)),
};

describe("patchXlsxChartPart — path 2 leaves every other byte as it was", () => {
  it("returns the identical string when nothing changes", () => {
    expect(patchXlsxChartPart(COLUMN_PART, COLUMN_REFS)).toBe(COLUMN_PART);
    expect(patchXlsxChartPart(COLUMN_PART, COLUMN_REFS, cellsFrom(SAME_CELLS))).toBe(COLUMN_PART);
  });

  it("rewrites only the c:f texts whose reference changed (the c15:f data-label ranges stay)", () => {
    expect(patchXlsxChartPart(COLUMN_PART, MOVED_REFS)).toBe(columnPartMoved());
    const one: SheetChartRefs = { ...COLUMN_REFS, series: [COLUMN_REFS.series[0]!, { ...COLUMN_REFS.series[1], values: "Data!$C$2:$C$6" }, COLUMN_REFS.series[2]!] };
    expect(patchXlsxChartPart(COLUMN_PART, one)).toBe(COLUMN_PART.replace("<c:f>Data!$C$2:$C$5</c:f>", "<c:f>Data!$C$2:$C$6</c:f>"));
  });

  it("keeps an unchanged reference's bytes even when spelled differently", () => {
    const spelled = COLUMN_PART.replace("<c:f>Data!$B$1</c:f>", "<c:f>=Data!$B$1</c:f>");
    expect(patchXlsxChartPart(spelled, COLUMN_REFS)).toBe(spelled);
  });

  it("writes a reference XML-escaped, through toXlsxFormula", () => {
    const refs: SheetChartRefs = { ...COLUMN_REFS, series: [{ ...COLUMN_REFS.series[0], name: "'R&D'!$B$1" }, ...COLUMN_REFS.series.slice(1)] };
    expect(patchXlsxChartPart(COLUMN_PART, refs)).toBe(COLUMN_PART.replace("<c:f>Data!$B$1</c:f>", "<c:f>'R&amp;D'!$B$1</c:f>"));
  });

  it("rewrites caches from the cells: ptCount, a pt per non-blank cell, numbers raw, text formatted; formatCode and extLst kept", () => {
    const cells = {
      ...SAME_CELLS,
      "Data!$C$1": [cell("South & East")],
      "Data!$A$2:$A$5": [cell(45292, "1/1/2024"), cell("Feb"), cell(null), cell("Apr")],
      // A blank, a text and an error: no c:pt in a number cache.
      "Data!$C$2:$C$5": [cell(4.5), cell(null), cell("n/a"), cell({ error: "#N/A" } as unknown as CellValue, "#N/A"), cell(9)],
    };
    const patched = patchXlsxChartPart(COLUMN_PART, COLUMN_REFS, cellsFrom(cells));
    const expected = COLUMN_PART.replace(
      `<c:f>Data!$C$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>South</c:v></c:pt></c:strCache>`,
      `<c:f>Data!$C$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>South &amp; East</c:v></c:pt></c:strCache>`,
    )
      .replaceAll(
        `<c:ptCount val="4"/><c:pt idx="0"><c:v>Jan</c:v></c:pt><c:pt idx="1"><c:v>Feb</c:v></c:pt><c:pt idx="2"><c:v>Mar</c:v></c:pt><c:pt idx="3"><c:v>Apr</c:v></c:pt>`,
        `<c:ptCount val="4"/><c:pt idx="0"><c:v>1/1/2024</c:v></c:pt><c:pt idx="1"><c:v>Feb</c:v></c:pt><c:pt idx="3"><c:v>Apr</c:v></c:pt>`,
      )
      .replace(
        `<c:formatCode>0.0</c:formatCode><c:ptCount val="4"/><c:pt idx="0"><c:v>4</c:v></c:pt><c:pt idx="2"><c:v>6</c:v></c:pt><c:pt idx="3"><c:v>8</c:v></c:pt>${CACHE_EXT}`,
        `<c:formatCode>0.0</c:formatCode><c:ptCount val="5"/><c:pt idx="0"><c:v>4.5</c:v></c:pt><c:pt idx="4"><c:v>9</c:v></c:pt>${CACHE_EXT}`,
      );
    expect(expected).not.toBe(COLUMN_PART);
    expect(patched).toBe(expected);
    expect(readXlsxChart(patched).visual.series[1]!.values).toEqual([4.5, null, null, null, 9]);
  });

  it("moves references and refreshes caches together, keeping every other byte", () => {
    const moved: Record<string, ChartCell[]> = {};
    for (const [ref, cells] of Object.entries(SAME_CELLS)) {
      moved[ref.replace(/\$(\d+)/g, (_, row: string) => `$${Number(row) + 1}`)] = cells;
    }
    moved["Data!$B$3:$B$6"] = [10, 20, 30, 99].map((v) => cell(v));
    const expected = columnPartMoved().replace(`<c:pt idx="3"><c:v>40</c:v></c:pt>`, `<c:pt idx="3"><c:v>99</c:v></c:pt>`);
    expect(patchXlsxChartPart(COLUMN_PART, MOVED_REFS, cellsFrom(moved))).toBe(expected);
  });

  it("leaves a cache whose reference does not resolve, and every cache without cellsOf", () => {
    const partial: Record<string, ChartCell[]> = { ...SAME_CELLS, "Data!$B$2:$B$5": [cell(1)] };
    delete partial["Data!$D$2:$D$5"];
    expect(patchXlsxChartPart(COLUMN_PART, COLUMN_REFS, cellsFrom(partial))).toBe(
      COLUMN_PART.replace(`<c:ptCount val="4"/><c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="1"><c:v>20</c:v></c:pt><c:pt idx="2"><c:v>30</c:v></c:pt><c:pt idx="3"><c:v>40</c:v></c:pt>`, `<c:ptCount val="1"/><c:pt idx="0"><c:v>1</c:v></c:pt>`),
    );
    expect(patchXlsxChartPart(COLUMN_PART, MOVED_REFS)).toBe(columnPartMoved());
  });

  it("leaves equal numbers spelled differently alone", () => {
    const spelled = COLUMN_PART.replace(`<c:pt idx="0"><c:v>4</c:v></c:pt>`, `<c:pt idx="0"><c:v>4.0000000000000004</c:v></c:pt>`);
    expect(patchXlsxChartPart(spelled, COLUMN_REFS, cellsFrom(SAME_CELLS))).toBe(spelled);
  });

  it("gives an uncached reference a cache right after its c:f", () => {
    const uncached = COLUMN_PART.replace(`<c:f>Data!$D$1</c:f>${strCache(["West"])}`, "<c:f>Data!$D$1</c:f>").replace(
      `<c:f>Data!$D$2:$D$5</c:f>${numCache([1, 2, 3, 4])}`,
      "<c:f>Data!$D$2:$D$5</c:f>",
    );
    expect(patchXlsxChartPart(uncached, COLUMN_REFS)).toBe(uncached);
    expect(patchXlsxChartPart(uncached, COLUMN_REFS, cellsFrom(SAME_CELLS))).toBe(
      uncached
        .replace("<c:f>Data!$D$1</c:f>", `<c:f>Data!$D$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>West</c:v></c:pt></c:strCache>`)
        .replace("<c:f>Data!$D$2:$D$5</c:f>", `<c:f>Data!$D$2:$D$5</c:f>${numCache([1, 2, 3, 4])}`),
    );
  });

  it("writes new elements with the part's own prefix", () => {
    const renamed = (xml: string): string => xml.replaceAll("<c:", "<ch:").replaceAll("</c:", "</ch:").replace(`xmlns:c="${C}"`, `xmlns:ch="${C}"`);
    const uncached = COLUMN_PART.replace(`<c:f>Data!$D$2:$D$5</c:f>${numCache([1, 2, 3, 4])}`, "<c:f>Data!$D$2:$D$5</c:f>");
    const patched = patchXlsxChartPart(renamed(uncached), MOVED_REFS, cellsFrom({ "Data!$D$3:$D$6": [cell(7)] }));
    expect(patched).toContain(`<ch:f>Data!$D$3:$D$6</ch:f><ch:numCache><ch:formatCode>General</ch:formatCode><ch:ptCount val="1"/><ch:pt idx="0"><ch:v>7</ch:v></ch:pt></ch:numCache>`);
    expect(patched).not.toContain("<c:");
  });

  it("rewrites a multi-level reference's c:f but never its cache", () => {
    const part = COLUMN_PART.replace(
      `<c:cat><c:strRef><c:f>Data!$A$2:$A$5</c:f>${strCache(MONTHS)}</c:strRef></c:cat>`,
      `<c:cat><c:multiLvlStrRef><c:f>Data!$A$2:$B$5</c:f><c:multiLvlStrCache><c:ptCount val="1"/><c:lvl><c:pt idx="0"><c:v>Jan</c:v></c:pt></c:lvl></c:multiLvlStrCache></c:multiLvlStrRef></c:cat>`,
    );
    const refs: SheetChartRefs = { ...COLUMN_REFS, series: [{ ...COLUMN_REFS.series[0], categories: "Data!$A$3:$B$6" }, ...COLUMN_REFS.series.slice(1)] };
    const patched = patchXlsxChartPart(part, refs);
    expect(patched).toBe(part.replace("<c:f>Data!$A$2:$B$5</c:f>", "<c:f>Data!$A$3:$B$6</c:f>"));
    expect(patchXlsxChartPart(part, refs, cellsFrom(SAME_CELLS))).toBe(patched);
  });

  it("patches only the series the refs reach, and a scatter's xVal/yVal", () => {
    const first: SheetChartRefs = { series: [MOVED_REFS.series[0]!], plotVisibleOnly: true };
    expect(patchXlsxChartPart(COLUMN_PART, first)).toBe(
      COLUMN_PART.replace("<c:f>Data!$B$1</c:f>", "<c:f>Data!$B$2</c:f>")
        .replace("<c:f>Data!$A$2:$A$5</c:f>", "<c:f>Data!$A$3:$A$6</c:f>")
        .replace("<c:f>Data!$B$2:$B$5</c:f>", "<c:f>Data!$B$3:$B$6</c:f>"),
    );
    const scatter: SheetChartRefs = { series: [{ name: "Data!$B$1", categories: "Data!$A$2:$A$5", values: "Data!$B$2:$B$5" }], plotVisibleOnly: true };
    expect(patchXlsxChartPart(SCATTER_PART, scatter, cellsFrom({ "Data!$A$2:$A$5": [cell(1), cell(2), cell(4), cell(8)] }))).toBe(
      SCATTER_PART.replace("<c:f>Data!$A$2:$A$4</c:f>", "<c:f>Data!$A$2:$A$5</c:f>")
        .replace("<c:f>Data!$B$2:$B$4</c:f>", "<c:f>Data!$B$2:$B$5</c:f>")
        .replace(numCache([1, 2, 4]), numCache([1, 2, 4, 8])),
    );
  });

  it("returns a part it cannot read unchanged", () => {
    expect(patchXlsxChartPart("<c:chartSpace", MOVED_REFS)).toBe("<c:chartSpace");
  });
});

// --- goldens ---------------------------------------------------------------------------------

const GOLDENS = [
  { label: "a clustered column with every reference", visual: GOLDEN_COLUMN_CHART_VISUAL, refs: GOLDEN_COLUMN_CHART_REFS, part: GOLDEN_COLUMN_CHART_PART },
  { label: "a scatter (xVal/yVal, plotVisOnly 0)", visual: GOLDEN_SCATTER_CHART_VISUAL, refs: GOLDEN_SCATTER_CHART_REFS, part: GOLDEN_SCATTER_CHART_PART },
  { label: "a line with no categories reference and a literal name", visual: GOLDEN_LINE_CHART_VISUAL, refs: GOLDEN_LINE_CHART_REFS, part: GOLDEN_LINE_CHART_PART },
] as const;

describe("new chart parts equal the hand-authored goldens (AC-35)", () => {
  for (const golden of GOLDENS) {
    it(`xlsxChartPartXml writes ${golden.label}`, () => {
      expect(xlsxChartPartXml(golden.visual, golden.refs)).toBe(golden.part);
    });

    it(`reading ${golden.label} gives the visual and references back`, () => {
      expect(readXlsxChart(golden.part)).toEqual({ visual: golden.visual, refs: golden.refs });
      expect(readChartPartRefs(golden.part)).toEqual({ refs: golden.refs, uncached: false });
      expect(chartReadOnlyReason(golden.part)).toBeUndefined();
      expect(patchXlsxChartPart(golden.part, golden.refs)).toBe(golden.part);
    });
  }

  it("parseChartXml alone gives the visual back where every series has categories", () => {
    for (const golden of GOLDENS.slice(0, 2)) {
      const { chartType, categories, series, options } = golden.visual;
      expect(parseChartXml(golden.part)).toEqual({ chartType, categories, series, options });
    }
  });

  it("chartGraphicFrameXml writes the anchor's canonical graphic frame", () => {
    const { cNvPrId, name, relId } = GOLDEN_CHART_GRAPHIC_FRAME_INPUT;
    expect(chartGraphicFrameXml(cNvPrId, name, relId)).toBe(GOLDEN_CHART_GRAPHIC_FRAME);
    expect(chartGraphicFrameXml(5, 'Sales "Q1" & <Q2>', "rId7")).toContain(`<xdr:cNvPr id="5" name="Sales &quot;Q1&quot; &amp; &lt;Q2&gt;"/>`);
  });

  it("the relationship and Override goldens carry core's chart constants", () => {
    expect(GOLDEN_CHART_DRAWING_RELATIONSHIP).toContain(`Type="${CHART_RELATIONSHIP}"`);
    expect(GOLDEN_CHART_OVERRIDE).toContain(`ContentType="${CHART_CONTENT_TYPE}"`);
    expect(GOLDEN_CHART_OVERRIDE).toContain(`PartName="/${nextChartPartName(["xl/charts/chart1.xml", "xl/drawings/drawing1.xml"])}"`);
  });

  it("a part without refs is SPEC-052's literal-chart form", () => {
    expect(xlsxChartPartXml(GOLDEN_COLUMN_CHART_VISUAL, undefined)).toContain("<c:f>Sheet1!$B$1</c:f>");
  });
});

describe("xlsxChartWireRefs", () => {
  it("writes each canonical reference through toXlsxFormula, keeps absent slots absent and carries plotVisibleOnly", () => {
    expect(xlsxChartWireRefs(GOLDEN_LINE_CHART_REFS)).toEqual({
      series: [{ name: "'R&D'!$B$1", values: "'R&D'!$B$2:$B$4" }, { values: "'R&D'!$C$2:$C$4" }],
      plotVisibleOnly: true,
    });
    expect(xlsxChartWireRefs({ series: [{ values: "(Data!$B$2:$B$3,Data!$B$5:$B$6)" }], plotVisibleOnly: false })).toEqual({
      series: [{ values: "(Data!$B$2:$B$3,Data!$B$5:$B$6)" }],
      plotVisibleOnly: false,
    });
  });
});
