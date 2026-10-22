// SPEC-076 — how a table's page slice paints: the repeated header run's DOM
// contract (AC-7), its identity with the original (AC-2) and the banding the
// component resolves at paint time (AC-24, AC-29).
import { describe, expect, it } from "vitest";
import { mount } from "@vue/test-utils";
import DocTableBlock from "./DocTableBlock.vue";
import { caretFromRepeatedHeaderPoint, domPointForCaret } from "../lib/domCaret";
import type { NamedStyle, Table } from "../lib/documentModel";

const cell = (text: string) => ({
  blocks: [{ type: "paragraph" as const, runs: [{ text, formatting: {} }] }],
});

function table(): Table {
  return {
    type: "table",
    columnWidthsCm: [4, 4],
    rows: [
      { isHeader: true, cells: [cell("Name"), cell("Qty")] },
      { cells: [cell("a"), cell("1")] },
      { cells: [cell("b"), cell("2")] },
      { cells: [cell("c"), cell("3")] },
    ],
  };
}

function mountSlice(rowStart: number, rowEnd: number, styles: NamedStyle[] = []) {
  return mount(DocTableBlock, {
    props: { block: table(), section: 0, blockIndex: 0, rowStart, rowEnd, styles },
  });
}

describe("DocTableBlock header repetition (SPEC-076)", () => {
  it("draws the header once for a slice that starts at the table's top (AC-6)", () => {
    const w = mountSlice(0, -1);
    expect(w.findAll("tbody > tr")).toHaveLength(4);
    expect(w.findAll("[data-repeated-header]")).toHaveLength(0);
  });

  it("redraws the header above a continuation's first body row (AC-1)", () => {
    const w = mountSlice(2, 4);
    const rows = w.findAll("tbody > tr");
    expect(rows).toHaveLength(3);
    expect(rows[0].attributes("data-repeated-header")).toBe("true");
    expect(rows[0].text()).toContain("Name");
    expect(rows[1].text()).toContain("b");
  });

  it("marks the repeat aria-hidden, uneditable and unselectable (AC-7)", () => {
    const w = mountSlice(2, 4);
    const repeated = w.get("[data-repeated-header]");
    expect(repeated.attributes("aria-hidden")).toBe("true");
    expect(repeated.attributes("contenteditable")).toBe("false");
    // data-canvas-ignore is the canvas's own "walk past this" contract — it
    // is what keeps find, spell-check, word count and caret resolution out
    // of the repeat (AC-9, AC-10).
    expect(repeated.attributes("data-canvas-ignore")).toBe("repeated-header");
    expect(repeated.classes()).toContain("doc-table-row-repeated");
  });

  it("renders the repeat identically to the original: same text, spans and cell style (AC-2)", () => {
    const whole = mountSlice(0, -1);
    const continuation = mountSlice(2, 4);
    const original = whole.findAll("tbody > tr")[0].findAll("td");
    const repeat = continuation.findAll("[data-repeated-header] td");
    expect(repeat).toHaveLength(original.length);
    repeat.forEach((repeated, index) => {
      expect(repeated.text()).toBe(original[index].text());
      expect(repeated.attributes("style")).toBe(original[index].attributes("style"));
      expect(repeated.attributes("data-row")).toBe(original[index].attributes("data-row"));
      expect(repeated.attributes("data-col")).toBe(original[index].attributes("data-col"));
    });
  });

  it("keeps the repeat's cells addressed to the REAL header row, so a click on one maps to it (AC-8)", () => {
    const w = mountSlice(2, 4);
    const paragraph = w.get("[data-repeated-header] .doc-paragraph").element;
    const root = document.createElement("div");
    root.append(w.element);
    const walker = document.createTreeWalker(paragraph, NodeFilter.SHOW_TEXT);
    let text = walker.nextNode() as Text | null;
    while (text && text.data.length === 0) {
      text = walker.nextNode() as Text | null;
    }
    // Two characters into "Name" — the offset nearest a click there.
    const caret = caretFromRepeatedHeaderPoint(root, text as Text, 2);
    expect(caret).toEqual({ section: 0, block: 0, offset: 2, cell: { row: 0, col: 0, block: 0 } });
  });

  it("never places the caret inside a repeat, even when only the continuation is realized (AC-9)", () => {
    // The page holding the real header row is virtualized away: all the DOM
    // has of row 0 is the continuation's redrawn copy.
    const w = mountSlice(2, 4);
    const root = document.createElement("div");
    root.append(w.element);
    expect(root.querySelectorAll('[data-row="0"][data-col="0"]').length).toBeGreaterThan(0);
    const point = domPointForCaret(root, {
      section: 0,
      block: 0,
      offset: 0,
      cell: { row: 0, col: 0, block: 0 },
    });
    expect(point).toBeNull();
  });

  it("does not resolve a point outside a repeat (AC-8)", () => {
    const w = mountSlice(0, -1);
    const root = document.createElement("div");
    root.append(w.element);
    const paragraph = w.get(".doc-paragraph").element;
    expect(caretFromRepeatedHeaderPoint(root, paragraph, 0)).toBeNull();
  });
});

describe("DocTableBlock banding (SPEC-076)", () => {
  const banded: NamedStyle[] = [
    {
      id: "S",
      name: "Banded",
      type: "table",
      banding: { bandedRows: { size: 1, shading: "#f2f2f2" } },
    },
  ];

  function bandedTable(look: Table["look"]): Table {
    const built = table();
    built.styleId = "S";
    built.look = look;
    return built;
  }

  it("paints the style's row banding without writing it into any cell (AC-24, AC-26, AC-29)", () => {
    const model = bandedTable({ bandedRows: true });
    const w = mount(DocTableBlock, { props: { block: model, section: 0, blockIndex: 0, styles: banded } });
    const rows = w.findAll("tbody > tr");
    // Row 0 is the header run, so banding counts from row 1: band 1 (row 1)
    // takes the shading, band 2 (row 2) does not.
    expect(rows[1].findAll("td")[0].attributes("style")).toContain("rgb(242, 242, 242)");
    expect(rows[2].findAll("td")[0].attributes("style") ?? "").not.toContain("rgb(242, 242, 242)");
    expect(model.rows[1].cells[0].shadingFill).toBeUndefined();
  });

  it("paints nothing extra when the look does not enable banding (AC-23)", () => {
    const w = mount(DocTableBlock, {
      props: { block: bandedTable({}), section: 0, blockIndex: 0, styles: banded },
    });
    expect(w.findAll("tbody > tr")[1].findAll("td")[0].attributes("style") ?? "").not.toContain("rgb(242, 242, 242)");
  });
});
