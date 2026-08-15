// SPEC-052: chart structure commands. Mirrors shapes.spec.ts's own fixture/
// testing conventions.
import { describe, it, expect } from "vitest";
import {
  chartAt,
  deleteChart,
  insertChart,
  insertChartFromTable,
  moveChart,
  setChartAltText,
  setChartData,
  setChartOptions,
  setChartSeriesColor,
  setChartSize,
  setChartWrap,
  type ChartLocation,
} from "./charts";
import {
  chartDataFromTableCells,
  NEW_CHART_HEIGHT_CM,
  NEW_CHART_WIDTH_CM,
  sampleChartCategories,
  sampleChartSeries,
} from "./core/charts";
import { caretSelection } from "./commands";
import { createDocument, isChart, type Chart, type DocumentModel, type Paragraph, type Table } from "./documentModel";
import type { CaretPosition } from "./textEditing";

function paragraph(text: string): Paragraph {
  return { type: "paragraph", runs: text.length > 0 ? [{ text, formatting: {} }] : [] };
}

function doc(...blocks: Array<Paragraph | Table | Chart>): DocumentModel {
  const model = createDocument("charts-fixture");
  model.sections = [{ blocks }];
  return model;
}

function caret(block: number, offset: number): CaretPosition {
  return { section: 0, block, offset };
}

function selectionAt(block: number, offset: number) {
  return caretSelection(caret(block, offset));
}

function baseChart(overrides: Partial<Chart> = {}): Chart {
  return {
    type: "chart",
    chartType: "columnClustered",
    categories: ["A", "B"],
    series: [{ name: "S", values: [1, 2], color: "#111111" }],
    options: { legend: "right", dataLabels: false, gridlines: true },
    widthCm: 12,
    heightCm: 7.5,
    wrap: "square",
    relativeToH: "column",
    relativeToV: "paragraph",
    offsetXCm: 0,
    offsetYCm: 0,
    ...overrides,
  };
}

function cell(text: string) {
  return { blocks: [paragraph(text)] };
}

function grid(rows: string[][]): Table {
  return {
    type: "table",
    rows: rows.map((cells) => ({ cells: cells.map(cell) })),
    columnWidthsCm: rows[0].map(() => 5),
  };
}

describe("insertChart (AC-2, AC-3)", () => {
  it("inserts a 12cm x 7.5cm chart with sample data at the caret", () => {
    const model = doc(paragraph(""));
    const { model: next, selection } = insertChart(model, selectionAt(0, 0), "columnClustered");
    const blocks = next.sections[0].blocks;
    expect(blocks).toHaveLength(2);
    expect(isChart(blocks[0])).toBe(true);
    const chart = blocks[0] as Chart;
    expect(chart.chartType).toBe("columnClustered");
    expect(chart.widthCm).toBe(NEW_CHART_WIDTH_CM);
    expect(chart.heightCm).toBe(NEW_CHART_HEIGHT_CM);
    expect(chart.categories).toEqual(sampleChartCategories());
    expect(chart.series.map((s) => s.name)).toEqual(sampleChartSeries().map((s) => s.name));
    expect(chart.sourcePart).toBeUndefined();
    expect(chart.edited).toBeUndefined();
    expect(selection.focus).toEqual(caret(1, 0));
  });

  it("splits a paragraph at a middle offset, chart between the two halves", () => {
    const model = doc(paragraph("hello world"));
    const { model: next } = insertChart(model, selectionAt(0, 5), "line");
    const blocks = next.sections[0].blocks;
    expect(blocks).toHaveLength(3);
    expect(blocks[0]).toMatchObject({ type: "paragraph" });
    expect(isChart(blocks[1])).toBe(true);
    expect((blocks[1] as Chart).chartType).toBe("line");
    expect(blocks[2]).toMatchObject({ type: "paragraph" });
  });
});

describe("chartAt", () => {
  it("returns the chart at a location, or null for anything else", () => {
    const model = doc(paragraph(""), baseChart());
    expect(chartAt(model, { section: 0, block: 1 })).not.toBeNull();
    expect(chartAt(model, { section: 0, block: 0 })).toBeNull();
  });
});

describe("setChartData (AC-7..AC-10)", () => {
  it("replaces categories/series and marks the chart edited", () => {
    const model = doc(baseChart());
    const location: ChartLocation = { section: 0, block: 0 };
    const { model: next } = setChartData(model, selectionAt(0, 0), location, ["X", "Y", "Z"], [
      { name: "New series", values: [1, null, 3], color: "#222222" },
    ]);
    const chart = chartAt(next, location)!;
    expect(chart.categories).toEqual(["X", "Y", "Z"]);
    expect(chart.series).toEqual([{ name: "New series", values: [1, null, 3], color: "#222222" }]);
    expect(chart.edited).toBe(true);
  });
});

describe("setChartOptions (AC-14..AC-20)", () => {
  it("replaces options and marks the chart edited", () => {
    const model = doc(baseChart());
    const location: ChartLocation = { section: 0, block: 0 };
    const { model: next } = setChartOptions(model, selectionAt(0, 0), location, {
      title: "Revenue",
      legend: "top",
      dataLabels: true,
      gridlines: false,
      valueAxisMin: 0,
      valueAxisMax: 100,
    });
    const chart = chartAt(next, location)!;
    expect(chart.options.title).toBe("Revenue");
    expect(chart.options.legend).toBe("top");
    expect(chart.options.valueAxisMin).toBe(0);
    expect(chart.edited).toBe(true);
  });
});

describe("setChartSeriesColor (AC-19)", () => {
  it("updates one series' colour and marks the chart edited", () => {
    const model = doc(baseChart({ series: [{ name: "A", values: [1], color: "#111111" }, { name: "B", values: [2], color: "#222222" }] }));
    const location: ChartLocation = { section: 0, block: 0 };
    const { model: next } = setChartSeriesColor(model, selectionAt(0, 0), location, 1, "#abcdef");
    const chart = chartAt(next, location)!;
    expect(chart.series[0].color).toBe("#111111");
    expect(chart.series[1].color).toBe("#abcdef");
    expect(chart.edited).toBe(true);
  });

  it("is a no-op for an out-of-range series index", () => {
    const model = doc(baseChart());
    const location: ChartLocation = { section: 0, block: 0 };
    const { model: next } = setChartSeriesColor(model, selectionAt(0, 0), location, 5, "#abcdef");
    expect(chartAt(next, location)!.edited).toBeUndefined();
  });
});

describe("placement commands never set `edited` (SPEC-052: 'only the data sheet and the options panel set it')", () => {
  const location: ChartLocation = { section: 0, block: 0 };

  it("setChartSize", () => {
    const model = doc(baseChart());
    const { model: next } = setChartSize(model, selectionAt(0, 0), location, 8, 5);
    const chart = chartAt(next, location)!;
    expect(chart.widthCm).toBe(8);
    expect(chart.heightCm).toBe(5);
    expect(chart.edited).toBeUndefined();
  });

  it("moveChart", () => {
    const model = doc(baseChart({ offsetXCm: 1, offsetYCm: 1 }));
    const { model: next } = moveChart(model, selectionAt(0, 0), location, 2, 3);
    const chart = chartAt(next, location)!;
    expect(chart.offsetXCm).toBe(3);
    expect(chart.offsetYCm).toBe(4);
    expect(chart.edited).toBeUndefined();
  });

  it("setChartWrap", () => {
    const model = doc(baseChart({ wrap: "square" }));
    const { model: next } = setChartWrap(model, selectionAt(0, 0), location, "behind");
    const chart = chartAt(next, location)!;
    expect(chart.wrap).toBe("behind");
    expect(chart.edited).toBeUndefined();
  });

  it("setChartAltText", () => {
    const model = doc(baseChart());
    const { model: next } = setChartAltText(model, selectionAt(0, 0), location, "A bar chart of revenue");
    const chart = chartAt(next, location)!;
    expect(chart.altText).toBe("A bar chart of revenue");
    expect(chart.edited).toBeUndefined();
  });
});

describe("deleteChart (AC-32)", () => {
  it("removes the chart block", () => {
    const model = doc(paragraph("before"), baseChart(), paragraph("after"));
    const { model: next } = deleteChart(model, selectionAt(1, 0), { section: 0, block: 1 });
    expect(next.sections[0].blocks).toHaveLength(2);
    expect(next.sections[0].blocks.some((b) => isChart(b))).toBe(false);
  });
});

describe("chartDataFromTableCells (AC-21, AC-22)", () => {
  it("reads series names from the header row and categories from the first column", () => {
    const cells = [
      ["", "Revenue", "Cost"],
      ["Q1", "10", "5"],
      ["Q2", "20", "8"],
    ];
    const result = chartDataFromTableCells(cells);
    expect(result.categories).toEqual(["Q1", "Q2"]);
    expect(result.series).toEqual([
      { name: "Revenue", values: [10, 20], color: expect.any(String) },
      { name: "Cost", values: [5, 8], color: expect.any(String) },
    ]);
    expect(result.unparsable).toEqual([]);
  });

  it("reports cells that could not be read as a number, leaving them null", () => {
    const cells = [
      ["", "Revenue"],
      ["Q1", "not a number"],
      ["Q2", "20"],
    ];
    const result = chartDataFromTableCells(cells);
    expect(result.series[0].values).toEqual([null, 20]);
    expect(result.unparsable).toEqual(["not a number"]);
  });

  it("treats an empty cell as a gap, not an unparsable cell", () => {
    const cells = [
      ["", "Revenue"],
      ["Q1", ""],
    ];
    const result = chartDataFromTableCells(cells);
    expect(result.series[0].values).toEqual([null]);
    expect(result.unparsable).toEqual([]);
  });

  it("returns empty data for a table too small to hold a header + data column", () => {
    expect(chartDataFromTableCells([["only"]])).toEqual({ categories: [], series: [], unparsable: [] });
    expect(chartDataFromTableCells([])).toEqual({ categories: [], series: [], unparsable: [] });
  });
});

describe("insertChartFromTable (AC-21)", () => {
  it("places a new chart immediately after the table with the table's data", () => {
    const table = grid([
      ["", "Revenue"],
      ["Q1", "10"],
      ["Q2", "20"],
    ]);
    const model = doc(table, paragraph("after"));
    const data = chartDataFromTableCells([
      ["", "Revenue"],
      ["Q1", "10"],
      ["Q2", "20"],
    ]);
    const { model: next } = insertChartFromTable(model, selectionAt(0, 0), { section: 0, block: 0 }, "columnClustered", data);
    const blocks = next.sections[0].blocks;
    expect(blocks).toHaveLength(3);
    expect(isChart(blocks[1])).toBe(true);
    const chart = blocks[1] as Chart;
    expect(chart.categories).toEqual(["Q1", "Q2"]);
    expect(chart.series[0].values).toEqual([10, 20]);
  });
});
