// SPEC-116 in the library: workbooks are first-class rows and cards — the
// XLSX/XLSM/XLTX/XLTM/ODS/OTS badges (format + package variant), the xlsx
// and ods chips right after odp (a variant filters with its format), the
// card's SheetMiniature and "N sheets" stat, and the shared search/sort/star/
// select/trash/download machinery over MIXED kinds with no new code path.
// SPEC-126 adds CSV and TSV as first-class workbook formats: their badges,
// chips right after ods, the card's miniature and "1 sheet", and bulk
// download of the exact source bytes. Same harness as HomeView.media.spec.ts.
import { beforeEach, afterEach, describe, it, expect, vi } from "vitest";
import { mount, flushPromises } from "@vue/test-utils";
import { createMemoryHistory, createRouter } from "vue-router";
import { equalBytes } from "../lib/core/equality";
import { unzip } from "../lib/core/zip";
import type { PackageVariant, SpreadsheetFormat } from "../lib/spreadsheet/model";

type FixtureName = "mini-book.xlsx" | "mini-book.ods" | "spec116-macros.xlsm" | "spec116-template.xltx";

/** One row per line of the Reference's "File names and labels" table. XLTM
 * and OTS have no fixture of their own: their variant is set on a read
 * model. */
const FLAVOURS: ReadonlyArray<{ title: string; fixture: FixtureName; variant?: PackageVariant; badge: string }> = [
  { title: "Budget", fixture: "mini-book.xlsx", badge: "XLSX" },
  { title: "Macros", fixture: "spec116-macros.xlsm", badge: "XLSM" },
  { title: "Invoice template", fixture: "spec116-template.xltx", badge: "XLTX" },
  { title: "Macro template", fixture: "spec116-macros.xlsm", variant: "macroEnabledTemplate", badge: "XLTM" },
  { title: "Stock", fixture: "mini-book.ods", badge: "ODS" },
  { title: "Stock template", fixture: "mini-book.ods", variant: "template", badge: "OTS" },
];

/** Unmounts each test's library (mounted into document.body). */
const cleanups: Array<() => void> = [];

beforeEach(() => {
  vi.resetModules();
  localStorage.clear();
  vi.doMock("../lib/idb", () => ({
    loadAllRecords: vi.fn().mockResolvedValue([]),
    putRecord: vi.fn().mockResolvedValue(undefined),
    deleteRecord: vi.fn().mockResolvedValue(undefined),
  }));
});

afterEach(() => {
  for (const cleanup of cleanups.splice(0)) {
    cleanup();
  }
  vi.doUnmock("../lib/idb");
  vi.doUnmock("../lib/core/triggerDownload");
  vi.restoreAllMocks();
});

async function loadModules() {
  const store = await import("../lib/documentStore");
  const { readSpreadsheetBytes } = await import("../lib/spreadsheet/open");
  const fixtures = await import("../lib/spreadsheet/fixtures/build");
  const { createBlankPresentation } = await import("../lib/presentation/slides");
  const { routes } = await import("../router");
  const { default: HomeView } = await import("./HomeView.vue");
  async function mountHome() {
    const router = createRouter({ history: createMemoryHistory(), routes });
    await router.push("/");
    const wrapper = mount(HomeView, { global: { plugins: [router] }, attachTo: document.body });
    cleanups.push(() => wrapper.unmount());
    // HomeView preloads the editor chunk on mount (SPEC-047 AC-3) and never
    // awaits it; join that in-flight import here so it never settles after
    // the file's environment is torn down (see HomeView.media.spec.ts).
    await import("./DocumentView.vue");
    await flushPromises();
    return { wrapper, router };
  }
  const builders: Record<FixtureName, () => Promise<Uint8Array>> = {
    "mini-book.xlsx": fixtures.buildMiniBookXlsxBytes,
    "mini-book.ods": fixtures.buildMiniBookOdsBytes,
    "spec116-macros.xlsm": fixtures.buildSpecMacrosXlsmBytes,
    "spec116-template.xltx": fixtures.buildSpecTemplateXltxBytes,
  };
  /** Reads a fixture into a workbook titled `title` (its package variant
   * optionally overridden) and stores it as last modified at `at`. */
  async function seedWorkbook(title: string, fixture: FixtureName, at: number, variant?: PackageVariant) {
    const bytes = await builders[fixture]();
    const format: SpreadsheetFormat = fixture.endsWith(".ods") ? "ods" : "xlsx";
    const model = await readSpreadsheetBytes(bytes, format, title);
    if (variant) {
      model.packageVariant = variant;
    }
    vi.spyOn(Date, "now").mockReturnValue(at);
    store.storeSpreadsheet(model);
    return { bytes, model };
  }
  /** Every flavour in FLAVOURS, newest first, beside a text document. */
  async function seedEveryFlavour() {
    seedText("doc-1", "Report", 10_000);
    for (const [index, flavour] of FLAVOURS.entries()) {
      await seedWorkbook(flavour.title, flavour.fixture, 9_000 - index * 100, flavour.variant);
    }
  }
  function seedText(id: string, title: string, at: number) {
    vi.spyOn(Date, "now").mockReturnValue(at);
    store.getOrCreateDocument(id).title = title;
  }
  function seedDeck(title: string, at: number) {
    const deck = createBlankPresentation();
    deck.title = title;
    vi.spyOn(Date, "now").mockReturnValue(at);
    store.storePresentation(deck);
    return deck;
  }
  /** SPEC-126: reads a delimited-text fixture into a workbook titled
   * `title` and stores it as last modified at `at`. */
  async function seedCsv(title: string, fixture: "csv" | "tsv", at: number) {
    const { readCsvBytes } = await import("../lib/spreadsheet/csv/parse");
    const bytes = fixture === "csv" ? fixtures.buildSpec126Semicolon1252CsvBytes() : fixtures.buildSpec126UnicodeTsvBytes();
    const model = readCsvBytes(bytes, fixture, title, { userLocaleTag: "en-US" });
    vi.spyOn(Date, "now").mockReturnValue(at);
    store.storeSpreadsheet(model);
    return { bytes, model };
  }
  return { store, seedWorkbook, seedEveryFlavour, seedText, seedDeck, seedCsv, mountHome };
}

async function selectAllRows(wrapper: { findAll: (selector: string) => Array<{ setValue: (value: boolean) => Promise<void> }> }) {
  for (const check of wrapper.findAll('[data-testid="row-select"]')) {
    await check.setValue(true);
  }
}

function rowTitles(wrapper: { findAll: (selector: string) => Array<{ text: () => string }> }): string[] {
  return wrapper.findAll('[data-testid="row-title"]').map((cell) => cell.text());
}

/** The downloaded package against its source: the same part names in the
 * same order, each part byte-identical (SPEC-008's relation for a package
 * whose ZIP container is rebuilt). */
async function expectSameParts(actual: Uint8Array, expected: Uint8Array): Promise<void> {
  const out = await unzip(actual);
  const source = await unzip(expected);
  expect([...out.keys()]).toEqual([...source.keys()]);
  for (const [name, data] of source) {
    expect(equalBytes(out.get(name) ?? new Uint8Array(0), data), name).toBe(true);
  }
}

describe("workbook rows and chips", () => {
  it("badges each workbook by format + package variant: XLSX, XLSM, XLTX, XLTM, ODS, OTS", async () => {
    const { seedEveryFlavour, mountHome } = await loadModules();
    await seedEveryFlavour();
    const { wrapper } = await mountHome();

    const badges = Object.fromEntries(
      wrapper
        .findAll('[data-testid="library-row"]')
        .map((row) => [row.get('[data-testid="row-title"]').text(), row.get('[data-testid="row-format"]').text()]),
    );
    expect(badges).toEqual({
      Report: "New",
      ...Object.fromEntries(FLAVOURS.map((flavour) => [flavour.title, flavour.badge])),
    });
  });

  it("puts the xlsx and ods chips right after odp, each counting and filtering its format's variants too", async () => {
    const { seedEveryFlavour, mountHome } = await loadModules();
    await seedEveryFlavour();
    const { wrapper } = await mountHome();

    const chipIds = wrapper.findAll('[data-testid^="filter-"]').map((chip) => chip.attributes("data-testid"));
    const odp = chipIds.indexOf("filter-odp");
    expect(odp).toBeGreaterThanOrEqual(0);
    expect(chipIds.slice(odp + 1, odp + 3)).toEqual(["filter-xlsx", "filter-ods"]);

    const xlsx = wrapper.get('[data-testid="filter-xlsx"]');
    const ods = wrapper.get('[data-testid="filter-ods"]');
    expect(xlsx.text()).toContain("XLSX");
    expect(xlsx.get(".chip-count").text()).toBe("4");
    expect(ods.text()).toContain("ODS");
    expect(ods.get(".chip-count").text()).toBe("2");

    await xlsx.trigger("click");
    expect(rowTitles(wrapper).toSorted()).toEqual(["Budget", "Invoice template", "Macro template", "Macros"]);
    await xlsx.trigger("click"); // off again
    await ods.trigger("click");
    expect(rowTitles(wrapper).toSorted()).toEqual(["Stock", "Stock template"]);
  });

  it('shows a workbook card with a SheetMiniature of model.preview and an "N sheets" stat', async () => {
    const { seedWorkbook, mountHome } = await loadModules();
    const { model } = await seedWorkbook("Budget", "mini-book.xlsx", 2000);
    const { wrapper } = await mountHome();
    await wrapper.get('[data-testid="view-grid"]').trigger("click");
    const card = () => wrapper.get(`[data-testid="library-row"][data-doc-id="${model.id}"]`);
    expect(card().classes()).toContain("library-card");

    // mini-book.xlsx has three sheets; the hidden one counts.
    await vi.waitFor(() => expect(card().get('[data-testid="row-counts"]').text()).toBe("3 sheets"), { timeout: 5_000 });
    await vi.waitFor(() => expect(card().find('[data-testid="sheet-miniature"]').exists()).toBe(true), { timeout: 5_000 });
    const cells = card()
      .findAll('[data-testid="sheet-miniature-cell"]')
      .map((cell) => cell.text());
    expect(cells).toContain("Revenue");
    // The miniature is the stored preview — the active sheet's top-left
    // corner, row by row — not a re-render of the rows.
    expect(cells).toEqual(model.preview?.cells.flat());
  });

  it("search, sort, starring and multi-select span mixed kinds with no special-casing", async () => {
    const { store, seedText, seedWorkbook, mountHome } = await loadModules();
    seedText("doc-1", "Quarterly Report", 3000);
    const { model } = await seedWorkbook("Quarterly Budget", "mini-book.xlsx", 2000);
    const { wrapper } = await mountHome();
    expect(rowTitles(wrapper)).toEqual(["Quarterly Report", "Quarterly Budget"]);

    // Subsequence search matches the workbook's title.
    await wrapper.get('[data-testid="library-search"]').setValue("qbud");
    expect(rowTitles(wrapper)).toEqual(["Quarterly Budget"]);
    await wrapper.get('[data-testid="library-search"]').setValue("");

    // Sorting by title orders it among the text documents.
    await wrapper.get('[data-testid="sort-key"]').setValue("title");
    expect(rowTitles(wrapper)).toEqual(["Quarterly Budget", "Quarterly Report"]);
    await wrapper.get('[data-testid="sort-key"]').setValue("modified");
    expect(rowTitles(wrapper)).toEqual(["Quarterly Report", "Quarterly Budget"]);

    // Starring the workbook row persists the flag and sorts it first.
    await wrapper.get(`[data-testid="library-row"][data-doc-id="${model.id}"] [data-testid="row-star"]`).trigger("click");
    expect(store.getStoredRecord(model.id)?.starred).toBe(true);
    expect(rowTitles(wrapper)).toEqual(["Quarterly Budget", "Quarterly Report"]);

    // Both rows select; the bulk bar counts them.
    await selectAllRows(wrapper);
    expect(wrapper.get('[data-testid="selection-count"]').text()).toContain("2");
  });
});

describe("bulk actions over mixed kinds", () => {
  it("bulk download sends each kind down its own path, the workbook part for part", async () => {
    const downloads: Array<{ bytes: Uint8Array; name: string; type: string }> = [];
    vi.doMock("../lib/core/triggerDownload", () => ({
      triggerBlobDownload: (bytes: Uint8Array, name: string, type: string) => {
        downloads.push({ bytes, name, type });
      },
    }));
    const { seedText, seedDeck, seedWorkbook, mountHome } = await loadModules();
    seedText("doc-1", "Report", 3000);
    seedDeck("Pitch", 2500);
    const { bytes } = await seedWorkbook("Budget", "mini-book.xlsx", 2000);
    const { wrapper } = await mountHome();

    await selectAllRows(wrapper);
    expect(wrapper.get('[data-testid="selection-count"]').text()).toContain("3");
    await wrapper.get('[data-testid="bulk-download"]').trigger("click");
    await vi.waitFor(() => expect(downloads).toHaveLength(3), { timeout: 10_000 });

    const byName = new Map(downloads.map((download) => [download.name, download]));
    expect([...byName.keys()].toSorted()).toEqual(["Budget.xlsx", "Pitch.pptx", "Report.docx"]);
    const workbook = byName.get("Budget.xlsx");
    expect(workbook?.type).toBe("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    await expectSameParts(workbook?.bytes ?? new Uint8Array(0), bytes);
  });

  it("bulk Move to trash and its Undo take every kind in one batch, the workbook restored intact", async () => {
    const { store, seedText, seedDeck, seedWorkbook, mountHome } = await loadModules();
    seedText("doc-1", "Report", 3000);
    const deck = seedDeck("Pitch", 2500);
    const { model } = await seedWorkbook("Budget", "mini-book.xlsx", 2000);
    const { wrapper } = await mountHome();

    await selectAllRows(wrapper);
    await wrapper.get('[data-testid="bulk-trash"]').trigger("click");
    expect(store.listDocuments()).toEqual([]);
    expect(store.listTrashedDocuments().map((entry) => entry.kind).toSorted()).toEqual([
      "presentation",
      "spreadsheet",
      "text",
    ]);

    await wrapper.get('[data-testid="toast-action"]').trigger("click");
    expect(store.listDocuments().map((entry) => entry.id).toSorted()).toEqual(["doc-1", deck.id, model.id].toSorted());
    const restored = store.getSpreadsheet(model.id);
    expect(restored?.sheets).toHaveLength(3);
    expect(restored?.sheets.map((sheet) => sheet.name)).toEqual(["Summary", "Data", "Hidden"]);
    expect(restored?.sheets.map((sheet) => sheet.rows.length)).toEqual([11, 500, 2]);
  });
});

describe("csv and tsv workbooks (SPEC-126)", () => {
  it("badges them CSV and TSV and puts a csv and a tsv chip right after ods, each counting and filtering its format", async () => {
    const { seedCsv, seedWorkbook, mountHome } = await loadModules();
    await seedWorkbook("Stock", "mini-book.ods", 4000);
    await seedCsv("Kunden", "csv", 3000);
    await seedCsv("Greetings", "tsv", 2000);
    const { wrapper } = await mountHome();

    const badges = Object.fromEntries(
      wrapper
        .findAll('[data-testid="library-row"]')
        .map((row) => [row.get('[data-testid="row-title"]').text(), row.get('[data-testid="row-format"]').text()]),
    );
    expect(badges).toEqual({ Stock: "ODS", Kunden: "CSV", Greetings: "TSV" });

    const chipIds = wrapper.findAll('[data-testid^="filter-"]').map((chip) => chip.attributes("data-testid"));
    const ods = chipIds.indexOf("filter-ods");
    expect(chipIds.slice(ods + 1, ods + 3)).toEqual(["filter-csv", "filter-tsv"]);
    const csv = wrapper.get('[data-testid="filter-csv"]');
    expect(csv.text()).toContain("CSV");
    expect(csv.get(".chip-count").text()).toBe("1");
    expect(wrapper.get('[data-testid="filter-tsv"] .chip-count').text()).toBe("1");
    await csv.trigger("click");
    expect(rowTitles(wrapper)).toEqual(["Kunden"]);
    await csv.trigger("click");
    await wrapper.get('[data-testid="filter-tsv"]').trigger("click");
    expect(rowTitles(wrapper)).toEqual(["Greetings"]);
  });

  it('shows a csv workbook\'s card with the SheetMiniature and "1 sheet"', async () => {
    const { seedCsv, mountHome } = await loadModules();
    const { model } = await seedCsv("Kunden", "csv", 2000);
    const { wrapper } = await mountHome();
    await wrapper.get('[data-testid="view-grid"]').trigger("click");
    const card = () => wrapper.get(`[data-testid="library-row"][data-doc-id="${model.id}"]`);
    await vi.waitFor(() => expect(card().get('[data-testid="row-counts"]').text()).toBe("1 sheet"), { timeout: 5_000 });
    await vi.waitFor(() => expect(card().find('[data-testid="sheet-miniature"]').exists()).toBe(true), { timeout: 5_000 });
    const cells = card()
      .findAll('[data-testid="sheet-miniature-cell"]')
      .map((cell) => cell.text());
    expect(cells).toContain("Müller");
    expect(cells).toContain("01234");
    // Displayed in the user's locale, read in the file's.
    expect(cells).toContain("1234.5");
  });

  it("searches, stars, trashes, restores and bulk-downloads them on the shared paths — the download the exact source bytes", async () => {
    const downloads: Array<{ bytes: Uint8Array; name: string; type: string }> = [];
    vi.doMock("../lib/core/triggerDownload", () => ({
      triggerBlobDownload: (bytes: Uint8Array, name: string, type: string) => {
        downloads.push({ bytes, name, type });
      },
    }));
    const { store, seedText, seedCsv, mountHome } = await loadModules();
    seedText("doc-1", "Report", 3000);
    const { model, bytes } = await seedCsv("Kunden", "csv", 2000);
    const { wrapper } = await mountHome();

    await wrapper.get('[data-testid="library-search"]').setValue("kdn");
    expect(rowTitles(wrapper)).toEqual(["Kunden"]);
    await wrapper.get('[data-testid="library-search"]').setValue("");
    await wrapper.get(`[data-testid="library-row"][data-doc-id="${model.id}"] [data-testid="row-star"]`).trigger("click");
    expect(store.getStoredRecord(model.id)?.starred).toBe(true);
    expect(rowTitles(wrapper)).toEqual(["Kunden", "Report"]);

    await selectAllRows(wrapper);
    await wrapper.get('[data-testid="bulk-download"]').trigger("click");
    await vi.waitFor(() => expect(downloads).toHaveLength(2), { timeout: 10_000 });
    const csv = downloads.find((download) => download.name === "Kunden.csv");
    expect(csv?.type).toBe("text/csv");
    expect(equalBytes(csv?.bytes ?? new Uint8Array(0), bytes)).toBe(true);

    await selectAllRows(wrapper);
    await wrapper.get('[data-testid="bulk-trash"]').trigger("click");
    expect(store.listTrashedDocuments().map((entry) => entry.id)).toContain(model.id);
    await wrapper.get('[data-testid="toast-action"]').trigger("click");
    const restored = store.getSpreadsheet(model.id);
    expect(restored?.format).toBe("csv");
    expect(restored?.csv).toEqual(model.csv);
    expect(restored?.sheets[0]?.rows).toHaveLength(5);
  });
});
