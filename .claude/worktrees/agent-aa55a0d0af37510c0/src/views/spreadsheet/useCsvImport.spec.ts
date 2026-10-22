// SPEC-126's import bar and dialog: the bar's sentence for every delimiter,
// encoding and locale, dismissal, the dialog preselected from the dialect
// with a live preview of the first 20 records, "Apply" re-reading the
// retained bytes in place while the workbook is unedited (same ids, still
// unedited, undo reset, download still the source bytes) and "Open as new
// workbook" once it is edited.
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { effectScope, nextTick, shallowRef, type EffectScope } from "vue";
import { flushPromises, mount } from "@vue/test-utils";
import CsvImportBar from "../../components/spreadsheet/CsvImportBar.vue";
import CsvImportDialog from "../../components/spreadsheet/CsvImportDialog.vue";
import { equalBytes } from "../../lib/core/equality";
import { csvSourceIntact, spreadsheetDownload } from "../../lib/spreadsheet/download";
import { readCsvBytes } from "../../lib/spreadsheet/csv/parse";
import { buildSpec126Semicolon1252CsvBytes, buildSpec126SepCsvBytes, buildSpec126UnicodeTsvBytes } from "../../lib/spreadsheet/fixtures/build";
import type { SpreadsheetModel } from "../../lib/spreadsheet/model";
import { CSV_PREVIEW_RECORDS, csvImportBarText, csvImportPreview, reimportedTitle, useCsvImport, type CsvImport } from "./useCsvImport";

const store = vi.hoisted(() => ({ storeSpreadsheet: vi.fn(), storeSpreadsheetEdit: vi.fn() }));
vi.mock("../../lib/documentStore", () => ({
  storeSpreadsheet: store.storeSpreadsheet,
  storeSpreadsheetEdit: (...args: unknown[]) => {
    store.storeSpreadsheetEdit(...args);
    return Promise.resolve();
  },
}));

let scope: EffectScope | null = null;

beforeEach(() => {
  localStorage.clear();
  store.storeSpreadsheet.mockClear();
  store.storeSpreadsheetEdit.mockClear();
});

afterEach(() => {
  scope?.stop();
  scope = null;
});

function german(): SpreadsheetModel {
  return readCsvBytes(buildSpec126Semicolon1252CsvBytes(), "csv", "Kunden", { userLocaleTag: "en-US" });
}

/** useCsvImport over `model`, with the view's book, undo and router faked. */
function harness(model: SpreadsheetModel | null) {
  const workbook = shallowRef<SpreadsheetModel | null>(model);
  const replace = vi.fn((next: SpreadsheetModel) => {
    workbook.value = next;
  });
  const reset = vi.fn();
  const push = vi.fn().mockResolvedValue(undefined);
  scope = effectScope();
  const csv = scope.run(() => useCsvImport({ book: { workbook, replace }, undo: { reset }, router: { push } }))!;
  return { csv, workbook, replace, reset, push };
}

describe("the bar's sentence", () => {
  it('reads "Opened as <delimiter>, <encoding>" plus the locale when it is not the user\'s', () => {
    expect(csvImportBarText({ delimiter: ";", encoding: "windows-1252", localeTag: "de-DE" }, "en-US")).toBe(
      "Opened as semicolon-separated, Windows-1252, numbers and dates read as German",
    );
    expect(csvImportBarText({ delimiter: ",", encoding: "utf-8", localeTag: "en-US" }, "en-US")).toBe("Opened as comma-separated, UTF-8");
    expect(csvImportBarText({ delimiter: "\t", encoding: "utf-16le", localeTag: "en-GB" }, "en-US")).toBe(
      "Opened as tab-separated, UTF-16, numbers and dates read as English (UK)",
    );
    expect(csvImportBarText({ delimiter: "|", encoding: "utf-16be", localeTag: "cs-CZ" }, "cs-CZ")).toBe("Opened as pipe-separated, UTF-16");
  });

  it("shows on a csv/tsv workbook until dismissed, and never on a package workbook", async () => {
    const { csv, workbook } = harness(german());
    expect(csv.barText.value).toBe("Opened as semicolon-separated, Windows-1252, numbers and dates read as German");
    csv.dismiss();
    expect(csv.barText.value).toBe("");
    // Another workbook on show brings the bar back.
    workbook.value = readCsvBytes(buildSpec126SepCsvBytes(), "csv", "Fruit", { userLocaleTag: "en-US" });
    await nextTick();
    expect(csv.barText.value).toBe("Opened as semicolon-separated, UTF-8");
    workbook.value = { ...workbook.value, id: "xlsx-book", format: "xlsx" };
    await nextTick();
    expect(csv.barText.value).toBe("");
  });

  it("renders the sentence with Change… and a dismiss button", async () => {
    const { csv } = harness(german());
    const wrapper = mount(CsvImportBar, { props: { csvImport: csv } });
    const bar = wrapper.get('[data-testid="csv-import-bar"]');
    expect(bar.text().replace(/\s+/g, " ")).toBe("Opened as semicolon-separated, Windows-1252, numbers and dates read as German · Change…");
    await wrapper.get('[data-testid="csv-import-change"]').trigger("click");
    expect(csv.dialogOpen.value).toBe(true);
    await wrapper.get('[data-testid="csv-import-dismiss"]').trigger("click");
    expect(wrapper.find('[data-testid="csv-import-bar"]').exists()).toBe(false);
    wrapper.unmount();
  });
});

describe("the dialog", () => {
  it("opens preselected from the dialect", () => {
    const { csv } = harness(german());
    csv.openDialog();
    expect(csv.dialogOpen.value).toBe(true);
    expect([csv.separator.value, csv.encoding.value, csv.localeTag.value]).toEqual(["semicolon", "windows-1252", "de-DE"]);
  });

  it("previews the first 20 records of the retained bytes, re-parsed on every change", async () => {
    const lines = Array.from({ length: 30 }, (_, index) => `r${index};a,b`).join("\n");
    const model = readCsvBytes(new TextEncoder().encode(lines), "csv", "Long", { userLocaleTag: "en-US" });
    const { csv } = harness(model);
    csv.openDialog();
    expect(CSV_PREVIEW_RECORDS).toBe(20);
    expect(csv.preview.value.records).toHaveLength(20);
    expect(csv.preview.value.records[0]).toEqual(["r0", "a,b"]);
    csv.separator.value = "comma";
    await nextTick();
    expect(csv.preview.value.records[0]).toEqual(["r0;a", "b"]);
    csv.encoding.value = "utf-16le";
    await nextTick();
    expect(csv.preview.value.records[0]).not.toEqual(["r0;a", "b"]);
  });

  it("shows why the chosen character set cannot read the file", () => {
    // UTF-16 read as UTF-8 is full of NULs: binary to the reader.
    expect(csvImportPreview(buildSpec126UnicodeTsvBytes(), "\t", "utf-8")).toEqual({
      records: [],
      error: "This file doesn't contain text we can read as CSV",
    });
  });

  it("leaves a sep= line out of the preview", () => {
    expect(csvImportPreview(buildSpec126SepCsvBytes(), ";", "utf-8").records[0]).toEqual(["Item", "Qty", "Price"]);
  });

  it("renders the three selects, the preview and Apply while unedited", async () => {
    const { csv } = harness(german());
    csv.openDialog();
    const wrapper = mount(CsvImportDialog, { props: { csvImport: csv }, attachTo: document.body });
    await flushPromises();
    const dialog = wrapper.get('[data-testid="csv-import-dialog"]');
    expect(dialog.text()).toContain("Text import options");
    const options = (testid: string) => wrapper.findAll(`[data-testid="${testid}"] option`).map((option) => option.text());
    expect(options("csv-import-separator")).toEqual(["Comma", "Semicolon", "Tab", "Pipe"]);
    expect(options("csv-import-encoding")).toEqual(["Unicode (UTF-8)", "Unicode (UTF-16 LE)", "Unicode (UTF-16 BE)", "Western (Windows-1252)"]);
    expect(options("csv-import-locale")).toEqual(["English (US)", "English (UK)", "German", "French", "Spanish", "Czech"]);
    expect((wrapper.get('[data-testid="csv-import-separator"]').element as HTMLSelectElement).value).toBe("semicolon");
    expect(wrapper.findAll('[data-testid="csv-import-preview-row"]')).toHaveLength(5);
    expect(wrapper.get('[data-testid="csv-import-preview"]').text()).toContain("Müller");
    await wrapper.get('[data-testid="csv-import-separator"]').setValue("comma");
    expect(csv.separator.value).toBe("comma");
    expect(wrapper.get('[data-testid="csv-import-preview-row"]').text()).toContain("Name;Zip;Amount");
    expect(wrapper.find('[data-testid="csv-import-apply"]').exists()).toBe(true);
    expect(wrapper.find('[data-testid="csv-import-open-new"]').exists()).toBe(false);
    wrapper.unmount();
  });
});

describe("Apply (unedited)", () => {
  it("replaces the first sheet's rows in place — same ids, still unedited, undo reset, download still the source bytes", async () => {
    const model = german();
    const sheetId = model.sheets[0]!.id;
    const { csv, workbook, replace, reset } = harness(model);
    csv.openDialog();
    csv.separator.value = "comma";
    csv.apply();
    expect(replace).toHaveBeenCalledTimes(1);
    const next = workbook.value!;
    expect(next).not.toBe(model);
    expect(next.id).toBe(model.id);
    expect(next.sheets[0]!.id).toBe(sheetId);
    expect(next.sheets[0]!.dirty).toBeUndefined();
    expect(next.structureDirty).toBeUndefined();
    expect(next.sheets[0]!.rows[0]!.cells[0]!.value).toBe("Name;Zip;Amount");
    expect(next.csv).toMatchObject({ delimiter: ",", encoding: "windows-1252", localeTag: "de-DE", sheetId });
    expect(reset).toHaveBeenCalledTimes(1);
    expect(store.storeSpreadsheetEdit).toHaveBeenCalledWith(next);
    expect(csv.dialogOpen.value).toBe(false);
    expect(csvSourceIntact(next)).toBe(true);
    expect(equalBytes((await spreadsheetDownload(next)).bytes, buildSpec126Semicolon1252CsvBytes())).toBe(true);
    // The bar reads the new dialect.
    expect(csv.barText.value).toBe("Opened as comma-separated, Windows-1252, numbers and dates read as German");
  });

  it("keeps the dialog open with the reason when the chosen options cannot read the file", () => {
    const { csv, replace } = harness(readCsvBytes(buildSpec126UnicodeTsvBytes(), "tsv", "Greetings", { userLocaleTag: "en-US" }));
    csv.openDialog();
    csv.encoding.value = "utf-8";
    csv.apply();
    expect(replace).not.toHaveBeenCalled();
    expect(csv.dialogOpen.value).toBe(true);
    expect(csv.error.value).toBe("This file doesn't contain text we can read as CSV");
  });
});

describe("Open as new workbook (edited)", () => {
  it("stores a new workbook titled \"<title> (re-imported)\" from the same bytes and navigates to it, leaving the edited one untouched", async () => {
    const model = german();
    model.sheets = [{ ...model.sheets[0]!, dirty: true, rows: [] }];
    const { csv, workbook, replace, push } = harness(model);
    expect(csv.edited.value).toBe(true);
    csv.openDialog();
    csv.localeTag.value = "en-US";
    // Apply is not the action once edited.
    csv.apply();
    expect(replace).not.toHaveBeenCalled();
    csv.openAsNew();
    expect(store.storeSpreadsheet).toHaveBeenCalledTimes(1);
    const fresh = store.storeSpreadsheet.mock.calls[0]![0] as SpreadsheetModel;
    expect(fresh.title).toBe("Kunden (re-imported)");
    expect(reimportedTitle("Kunden")).toBe("Kunden (re-imported)");
    expect(fresh.id).not.toBe(model.id);
    expect(fresh.csv).toMatchObject({ delimiter: ";", encoding: "windows-1252", localeTag: "en-US" });
    expect(fresh.sheets[0]!.rows.length).toBeGreaterThan(0);
    expect(push).toHaveBeenCalledWith(`/sheets/${fresh.id}`);
    expect(workbook.value).toBe(model);
    expect(model.sheets[0]!.rows).toEqual([]);
    expect(csv.dialogOpen.value).toBe(false);
  });

  it("shows Open as new workbook instead of Apply once the workbook is edited", async () => {
    const model = german();
    model.structureDirty = true;
    const { csv } = harness(model);
    csv.openDialog();
    const wrapper = mount(CsvImportDialog, { props: { csvImport: csv as CsvImport }, attachTo: document.body });
    await flushPromises();
    expect(wrapper.find('[data-testid="csv-import-apply"]').exists()).toBe(false);
    expect(wrapper.get('[data-testid="csv-import-open-new"]').text()).toBe("Open as new workbook");
    wrapper.unmount();
  });
});
