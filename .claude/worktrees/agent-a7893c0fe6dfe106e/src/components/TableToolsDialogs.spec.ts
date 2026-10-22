// SPEC-076 — the three Table dialogs' own surfaces: the gallery's live
// previews and selection (AC-31, AC-32, AC-34), the separator dialog's live
// row/column preview and custom-separator validation (AC-53..AC-55), and the
// sort dialog's three key rows (AC-38).
import { describe, expect, it } from "vitest";
import { mount } from "@vue/test-utils";
import TableStyleGalleryDialog from "./TableStyleGalleryDialog.vue";
import TableSeparatorDialog from "./TableSeparatorDialog.vue";
import TableSortDialog from "./TableSortDialog.vue";
import type { Paragraph, Table } from "../lib/documentModel";

function cell(text: string) {
  return { blocks: [{ type: "paragraph" as const, runs: [{ text, formatting: {} }] }] };
}

function table(): Table {
  return {
    type: "table",
    columnWidthsCm: [4, 4],
    rows: [
      { isHeader: true, cells: [cell("Quarter"), cell("Revenue")] },
      { cells: [cell("Q1"), cell("120")] },
      { cells: [cell("Q2"), cell("138")] },
    ],
  };
}

describe("TableStyleGalleryDialog (AC-31, AC-32, AC-34, AC-35)", () => {
  it("offers the five built-ins plus None", () => {
    const w = mount(TableStyleGalleryDialog, { props: { table: table() } });
    expect(w.findAll('[role="option"]').map((tile) => tile.attributes("data-testid"))).toEqual([
      "table-style-none",
      "table-style-TableStylePlain",
      "table-style-TableStyleGrid",
      "table-style-TableStyleBandedRows",
      "table-style-TableStyleHeader",
      "table-style-TableStyleAccent",
    ]);
  });

  it("previews the SELECTED table's own content, not a generic sample (AC-32)", () => {
    const w = mount(TableStyleGalleryDialog, { props: { table: table() } });
    const tile = w.get('[data-testid="table-style-TableStyleAccent"]');
    expect(tile.text()).toContain("Quarter");
    expect(tile.text()).toContain("Q1");
    // …and with that look's shading actually painted (the built-in Accent
    // bands its rows), which a generic sample could not show.
    expect(tile.html()).toContain("rgb(217, 226, 243)");
  });

  it("marks the applied entry selected, and None when there is none (AC-34)", () => {
    const applied = table();
    applied.styleId = "TableStyleGrid";
    const w = mount(TableStyleGalleryDialog, { props: { table: applied } });
    expect(w.get('[data-testid="table-style-TableStyleGrid"]').attributes("aria-selected")).toBe("true");
    expect(w.get('[data-testid="table-style-none"]').attributes("aria-selected")).toBe("false");

    const plain = mount(TableStyleGalleryDialog, { props: { table: table() } });
    expect(plain.get('[data-testid="table-style-none"]').attributes("aria-selected")).toBe("true");
  });

  it("shows the six look checkboxes, checked from the table's own look (AC-35)", async () => {
    const looked = table();
    looked.look = { firstRow: true };
    const w = mount(TableStyleGalleryDialog, { props: { table: looked } });
    expect(w.findAll('[data-testid^="table-look-"]')).toHaveLength(6);
    expect((w.get('[data-testid="table-look-firstRow"]').element as HTMLInputElement).checked).toBe(true);
    expect((w.get('[data-testid="table-look-lastRow"]').element as HTMLInputElement).checked).toBe(false);
    await w.get('[data-testid="table-look-bandedRows"]').trigger("change");
    expect(w.emitted("toggle-look")).toEqual([["bandedRows"]]);
  });
});

describe("TableSeparatorDialog (AC-53, AC-54, AC-55, AC-61)", () => {
  const paragraphs: Paragraph[] = [
    { type: "paragraph", runs: [{ text: "a\tb\tc", formatting: {} }] },
    { type: "paragraph", runs: [{ text: "d,e", formatting: {} }] },
  ];

  it("offers Tab (default), Comma, Semicolon and Custom (AC-53)", () => {
    const w = mount(TableSeparatorDialog, { props: { paragraphs } });
    expect(w.findAll('[data-testid^="separator-"]').map((el) => el.attributes("data-testid"))).toEqual([
      "separator-tab",
      "separator-comma",
      "separator-semicolon",
      "separator-custom",
      "separator-custom-char",
    ]);
    expect((w.get('[data-testid="separator-tab"]').element as HTMLInputElement).checked).toBe(true);
  });

  it("previews the row and column count, live as the separator changes (AC-54)", async () => {
    const w = mount(TableSeparatorDialog, { props: { paragraphs } });
    expect(w.get('[data-testid="convert-preview"]').text()).toBe("2 rows × 3 columns");
    await w.get('[data-testid="separator-comma"]').setValue();
    expect(w.get('[data-testid="convert-preview"]').text()).toBe("2 rows × 2 columns");
  });

  it("refuses an empty, whitespace-only or multi-character custom separator (AC-55)", async () => {
    const w = mount(TableSeparatorDialog, { props: { paragraphs } });
    const error = () => w.find('[data-testid="separator-error"]').exists();
    await w.get('[data-testid="separator-custom"]').setValue();
    expect(error()).toBe(true);
    await w.get('[data-testid="separator-custom-char"]').setValue(" ");
    expect(error()).toBe(true);
    await w.get('[data-testid="separator-custom-char"]').setValue("||");
    expect(error()).toBe(true);
    await w.get('[data-testid="separator-custom-char"]').setValue("|");
    expect(error()).toBe(false);
    await w.get('[data-testid="convert-confirm"]').trigger("click");
    expect(w.emitted("apply")).toEqual([["custom", "|"]]);
  });

  it("shows no row/column preview in the table→text direction (AC-61)", () => {
    const w = mount(TableSeparatorDialog, { props: { direction: "table-to-text" } });
    expect(w.find('[data-testid="convert-preview"]').exists()).toBe(false);
    expect(w.findAll('[data-testid^="separator-"]').length).toBe(5);
  });
});

describe("TableSortDialog (AC-38, AC-39, AC-50)", () => {
  it("offers three key rows, each limited to the table's own columns (AC-38)", () => {
    const w = mount(TableSortDialog, { props: { table: table() } });
    expect(w.findAll('[data-testid^="sort-column-"]')).toHaveLength(3);
    // "(none)" plus one option per column.
    expect(w.get('[data-testid="sort-column-0"]').findAll("option")).toHaveLength(3);
    expect(w.findAll('[data-testid^="sort-type-"]')).toHaveLength(3);
    expect(w.findAll('[data-testid^="sort-direction-"]')).toHaveLength(3);
  });

  it("seeds key one from the caller and leaves keys two and three at (none) (AC-39)", () => {
    const w = mount(TableSortDialog, {
      props: {
        table: table(),
        initialKeys: [{ columnIndex: 1, direction: "asc", type: "text" }],
      },
    });
    expect((w.get('[data-testid="sort-column-0"]').element as HTMLSelectElement).value).toBe("1");
    expect((w.get('[data-testid="sort-column-1"]').element as HTMLSelectElement).value).toBe("-1");
    expect((w.get('[data-testid="sort-column-2"]').element as HTMLSelectElement).value).toBe("-1");
  });

  it("disables OK with the caller's stated reason (AC-49, AC-50)", () => {
    const w = mount(TableSortDialog, {
      props: { table: table(), blockedReason: "there is nothing to sort" },
    });
    expect(w.get('[data-testid="sort-blocked-reason"]').text()).toBe("there is nothing to sort");
    expect(w.get('[data-testid="sort-confirm"]').attributes("disabled")).toBeDefined();
  });

  it("re-asks the caller when the header toggle moves the sort range (AC-41)", async () => {
    const w = mount(TableSortDialog, { props: { table: table(), initialHasHeader: true } });
    await w.get('[data-testid="sort-has-header"]').setValue(false);
    expect(w.emitted("header-change")).toEqual([[false]]);
  });
});
