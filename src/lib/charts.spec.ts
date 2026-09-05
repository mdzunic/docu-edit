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
  parseChartNumber,
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

  it("stores a ThemeColorRef verbatim; untouched series keep their own references (SPEC-077 AC-23)", () => {
    const untouched = { theme: "accent4", lumMod: 0.6, lumOff: 0.4 } as const;
    const model = doc(
      baseChart({
        series: [
          { name: "A", values: [1], color: untouched },
          { name: "B", values: [2], color: "#222222" },
        ],
      }),
    );
    const location: ChartLocation = { section: 0, block: 0 };
    const { model: next } = setChartSeriesColor(model, selectionAt(0, 0), location, 1, { theme: "accent2" });
    const chart = chartAt(next, location)!;
    // The changed series stores the reference, not a resolved hex...
    expect(chart.series[1].color).toEqual({ theme: "accent2" });
    // ...and the untouched series keeps its very own reference object.
    expect(chart.series[0].color).toBe(untouched);
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
    // SPEC-077 AC-46: a new chart's series are seeded with theme accent
    // references, not literals, so a theme change recolours the chart.
    expect(result.series).toEqual([
      { name: "Revenue", values: [10, 20], color: { theme: "accent1" } },
      { name: "Cost", values: [5, 8], color: { theme: "accent2" } },
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
    // A header row with nothing under it has nothing to chart either.
    expect(chartDataFromTableCells([["", "Revenue"]])).toEqual({ categories: [], series: [], unparsable: [] });
  });

  it("treats a first row of nothing but numbers as data, naming the series itself", () => {
    const cells = [
      ["2021", "10", "5"],
      ["2022", "20", "8"],
    ];
    const result = chartDataFromTableCells(cells);
    // The first column stays the category column even when it holds numbers.
    expect(result.categories).toEqual(["2021", "2022"]);
    expect(result.series.map((s) => s.name)).toEqual(["Series 1", "Series 2"]);
    expect(result.series.map((s) => s.values)).toEqual([
      [10, 20],
      [5, 8],
    ]);
    expect(result.unparsable).toEqual([]);
  });

  it("reads the numbers people actually type into tables (AC-22)", () => {
    const cells = [
      ["", "Revenue", "Share"],
      ["Q1", "$1,250.50", "12.5%"],
      ["Q2", "1.250,5", "10 %"],
      ["Q3", "(1,000)", "0,125"],
      ["Q4", "1 000 kg", "n/a"],
    ];
    const result = chartDataFromTableCells(cells);
    expect(result.series[0].values).toEqual([1250.5, 1250.5, -1000, 1000]);
    expect(result.series[1].values).toEqual([12.5, 10, 0.125, null]);
    expect(result.unparsable).toEqual(["n/a"]);
  });
});

describe("parseChartNumber", () => {
  it.each([
    ["10", 10],
    [" 10 ", 10],
    ["1.5", 1.5],
    ["1,5", 1.5],
    ["1,000", 1000],
    ["12,500", 12500],
    ["1.500", 1.5],
    ["0,125", 0.125],
    ["1234,5", 1234.5],
    ["1,250.5", 1250.5],
    ["1.250,5", 1250.5],
    ["1 000", 1000],
    ["1\u00a0000", 1000],
    ["1'000", 1000],
    ["1,000,000", 1000000],
    ["1.000.000", 1000000],
    ["$10", 10],
    ["10 €", 10],
    ["€ 1.234,56", 1234.56],
    ["10%", 10],
    ["12.5%", 12.5],
    ["10 kg", 10],
    ["(1,000)", -1000],
    ["-5", -5],
    ["−5", -5],
    ["+5", 5],
    ["1.5e3", 1500],
  ])("reads %j as %d", (raw, expected) => {
    expect(parseChartNumber(raw)).toBe(expected);
  });

  it.each([["not a number"], [""], ["   "], ["abc"], ["1,2,3.4"], ["1,00,000"], ["n/a"], ["--"], ["1.2.3,4,5"]])(
    "rejects %j",
    (raw) => {
      expect(parseChartNumber(raw)).toBeNull();
    },
  );
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

const chartIn = (model: DocumentModel, block: number): Chart => {
  const found = chartAt(model, { section: 0, block });
  if (!found) {
    throw new Error(`no chart at block ${block}`);
  }
  return found;
};

const paragraphIn = (model: DocumentModel, block: number): Paragraph => {
  const found = model.sections[0].blocks[block];
  if (found.type !== "paragraph") {
    throw new Error(`no paragraph at block ${block}`);
  }
  return found;
};

// --- anchor pairing (SPEC-078) ----------------------------------------------

describe("chart anchor pairing (SPEC-078)", () => {
  it("insertChart pairs the chart with the caret paragraph", () => {
    const { model } = insertChart(doc(paragraph("Anchor")), selectionAt(0, 6), "columnClustered");
    expect(model.sections[0].blocks.map((b) => b.type)).toEqual(["paragraph", "chart", "paragraph"]);
    const ordinal = paragraphIn(model, 0).anchorOrdinal;
    expect(ordinal).toBeDefined();
    expect(chartIn(model, 1).anchorOrdinal).toBe(ordinal);
  });

  it("insertChartFromTable pairs with the paragraph following the table when there is one", () => {
    const table: Table = {
      type: "table",
      rows: [{ cells: [{ blocks: [paragraph("A")] }, { blocks: [paragraph("1")] }] }],
      columnWidthsCm: [5, 5],
    };
    const model = doc(table, paragraph("After"));
    const data = { categories: ["A"], series: [{ name: "S", values: [1], color: "#4e79a7" }], unparsable: [] };
    const { model: next } = insertChartFromTable(model, selectionAt(1, 0), { section: 0, block: 0 }, "barClustered", data);
    expect(next.sections[0].blocks.map((b) => b.type)).toEqual(["table", "chart", "paragraph"]);
    const ordinal = paragraphIn(next, 2).anchorOrdinal;
    expect(ordinal).toBeDefined();
    expect(chartIn(next, 1).anchorOrdinal).toBe(ordinal);
  });

  it("insertChartFromTable stays unpaired when nothing paragraph-shaped follows", () => {
    const table: Table = {
      type: "table",
      rows: [{ cells: [{ blocks: [paragraph("A")] }] }],
      columnWidthsCm: [5],
    };
    const { model: next } = insertChartFromTable(
      doc(table),
      selectionAt(0, 0),
      { section: 0, block: 0 },
      "barClustered",
      { categories: ["A"], series: [{ name: "S", values: [1], color: "#4e79a7" }], unparsable: [] },
    );
    expect(chartIn(next, 1).anchorOrdinal).toBeUndefined();
  });

  it("deleteChart releases the pairing when it was the last referrer", () => {
    const { model } = insertChart(doc(paragraph("Anchor")), selectionAt(0, 6), "columnClustered");
    const { model: next } = deleteChart(model, selectionAt(1, 0), { section: 0, block: 1 });
    expect(paragraphIn(next, 0).anchorOrdinal).toBeUndefined();
  });

  it("setChartWrap to inline drops and releases; back to floating re-pairs", () => {
    const { model } = insertChart(doc(paragraph("Anchor")), selectionAt(0, 6), "columnClustered");
    const location: ChartLocation = { section: 0, block: 1 };
    const { model: inline } = setChartWrap(model, selectionAt(1, 0), location, "inline");
    expect(chartIn(inline, 1).anchorOrdinal).toBeUndefined();
    expect(paragraphIn(inline, 0).anchorOrdinal).toBeUndefined();
    const { model: floating } = setChartWrap(inline, selectionAt(1, 0), location, "square");
    const ordinal = paragraphIn(floating, 0).anchorOrdinal;
    expect(ordinal).toBeDefined();
    expect(chartIn(floating, 1).anchorOrdinal).toBe(ordinal);
  });
});
