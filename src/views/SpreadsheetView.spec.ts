// The workbook view at /sheets/:id (SPEC-116 "The read-only grid"): SPEC-058's
// app bar with the workbook's title, badge and format label, Download (the
// file that came in, part for part) and SPEC-060's Close; the not-found
// state; the view's own kind hand-off; the sheet tabs (the ARIA tablist
// pattern, and Ctrl+PageUp/PageDown from the grid); the DOM-virtualised
// read-only grid — its ARIA, headers, the General display and alignment, and
// a million-row sheet's bounded DOM; foreign sheets; and the reader-version
// upgrade path through the view. Same harness as MediaView.spec.ts: fresh
// modules per test (the store is module state), idb mocked, the view mounted
// at its route through a memory router — attached to the document so focus
// is real.
import { afterEach, beforeEach, describe, expect, it, vi, type MockInstance } from "vitest";
import { flushPromises, mount } from "@vue/test-utils";
import { createMemoryHistory, createRouter, type Router } from "vue-router";
import { equalBytes } from "../lib/core/equality";
import { unzip } from "../lib/core/zip";
import { parseA1 } from "../lib/spreadsheet/address";
import { SPREADSHEET_READER_VERSION } from "../lib/spreadsheet/open";
import type {
  Cell,
  CellValue,
  PackageVariant,
  Row,
  Sheet,
  SpreadsheetFormat,
  SpreadsheetModel,
} from "../lib/spreadsheet/model";

type FixtureName = "mini-book.xlsx" | "mini-book.ods" | "spec116-macros.xlsm" | "spec116-template.xltx";

/** What a test reads off a rendered wrapper (VueWrapper and DOMWrapper alike). */
interface Finder {
  findAll(selector: string): Array<{ text(): string; attributes(key: string): string | undefined }>;
}

/** Every download the view triggers (triggerBlobDownload is mocked). */
const downloads: Array<{ bytes: Uint8Array; name: string; type: string }> = [];
/** Unmounts each test's view (mounted into document.body). */
const cleanups: Array<() => void> = [];

beforeEach(() => {
  vi.resetModules();
  localStorage.clear();
  downloads.length = 0;
  vi.doMock("../lib/idb", () => ({
    loadAllRecords: vi.fn().mockResolvedValue([]),
    putRecord: vi.fn().mockResolvedValue(undefined),
    deleteRecord: vi.fn().mockResolvedValue(undefined),
  }));
  vi.doMock("../lib/core/triggerDownload", () => ({
    triggerBlobDownload: (bytes: Uint8Array, name: string, type: string) => {
      downloads.push({ bytes, name, type });
    },
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
  const { routes } = await import("../router");
  const { default: SpreadsheetView } = await import("./SpreadsheetView.vue");
  const builders: Record<FixtureName, () => Promise<Uint8Array>> = {
    "mini-book.xlsx": fixtures.buildMiniBookXlsxBytes,
    "mini-book.ods": fixtures.buildMiniBookOdsBytes,
    "spec116-macros.xlsm": fixtures.buildSpecMacrosXlsmBytes,
    "spec116-template.xltx": fixtures.buildSpecTemplateXltxBytes,
  };
  /** A fixture read as the open funnel reads it, titled by its base name. */
  async function readFixture(name: FixtureName) {
    const bytes = await builders[name]();
    const format: SpreadsheetFormat = name.endsWith(".ods") ? "ods" : "xlsx";
    const model = await readSpreadsheetBytes(bytes, format, name.replace(/\.[^.]+$/, ""));
    return { bytes, model };
  }
  /** Mounts the view at /sheets/:id; `beforeMount` runs once the route is
   * entered and before the view loads. */
  async function mountSheets(id: string, beforeMount?: (router: Router) => void) {
    const router = createRouter({ history: createMemoryHistory(), routes });
    await router.push(`/sheets/${id}`);
    beforeMount?.(router);
    const wrapper = mount(SpreadsheetView, { global: { plugins: [router] }, attachTo: document.body });
    cleanups.push(() => wrapper.unmount());
    await flushPromises();
    return { wrapper, router };
  }
  return { store, readFixture, mountSheets };
}

/** SPEC-117's required layout fields at Excel's defaults: 64 px columns and
 * 15 pt (20 px) rows, nothing merged. */
const DEFAULT_LAYOUT = { columns: [], defaultColWidthPx: 64, defaultRowHeightPt: 15, merges: [] } satisfies Partial<Sheet>;

/** A grid sheet from [row, col, value] triples (0-based) — sparse, rows and
 * cells ascending, as the model's invariants require. */
function gridSheet(
  name: string,
  cells: ReadonlyArray<readonly [number, number, CellValue]>,
  visibility: Sheet["visibility"] = "visible",
): Sheet {
  const rows: Row[] = [];
  for (const [index, col, value] of cells.toSorted((a, b) => a[0] - b[0] || a[1] - b[1])) {
    const last = rows.at(-1);
    if (last?.index === index) {
      last.cells.push({ col, value });
    } else {
      rows.push({ index, cells: [{ col, value }] });
    }
  }
  return { id: crypto.randomUUID(), name, visibility, kind: "grid", rows, ...DEFAULT_LAYOUT };
}

/** A foreign sheet (chartsheet, dialogsheet, macrosheet): a placeholder, its
 * part carried verbatim. */
function foreignSheet(name: string, partName: string): Sheet {
  return { id: crypto.randomUUID(), name, visibility: "visible", kind: "foreign", rows: [], partName, ...DEFAULT_LAYOUT };
}

/** A hand-built workbook with no source package (`opaque` is empty, so it
 * declares no part content types and nothing here downloads it), at the
 * current reader version so no upgrade runs. */
function handBuiltBook(sheets: Sheet[], activeSheetIndex = 0): SpreadsheetModel {
  return {
    kind: "spreadsheet",
    id: crypto.randomUUID(),
    title: "Hand built",
    format: "xlsx",
    readerVersion: SPREADSHEET_READER_VERSION,
    sheets,
    activeSheetIndex,
    dateSystem: "1900",
    opaque: [],
    styles: [{}],
    defaultStyleId: 0,
  };
}

/** Five sheets, three of them visible: very-hidden "Secret" and hidden
 * "Stash" sit between the visible ones. */
function bookWithHiddenSheets(): SpreadsheetModel {
  return handBuiltBook([
    gridSheet("One", [[0, 0, "one"]]),
    gridSheet("Secret", [[0, 0, "secret"]], "veryHidden"),
    gridSheet("Two", [[0, 0, "two"]]),
    gridSheet("Stash", [[0, 0, "stash"]], "hidden"),
    gridSheet("Three", [[0, 0, "three"]]),
  ]);
}

const MILLION = 1_000_000;

/** The million-row sheet: rows 0..999,999, each holding its own index in
 * column A, with the last 50 rows filled across B..Z as well. */
function millionRowSheet(): Sheet {
  const rows: Row[] = [];
  for (let index = 0; index < MILLION; index += 1) {
    const cells: Cell[] = [{ col: 0, value: index }];
    if (index >= MILLION - 50) {
      for (let col = 1; col <= 25; col += 1) {
        cells.push({ col, value: col });
      }
    }
    rows.push({ index, cells });
  }
  return { id: crypto.randomUUID(), name: "Big", visibility: "visible", kind: "grid", rows, ...DEFAULT_LAYOUT };
}

/** The tab strip, one "name:aria-selected:tabindex" string per tab. */
function tabStrip(wrapper: Finder): string[] {
  return wrapper
    .findAll('[data-testid="sheet-tab"]')
    .map((tab) => `${tab.text()}:${tab.attributes("aria-selected")}:${tab.attributes("tabindex")}`);
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

describe("SpreadsheetView app bar (SPEC-058, SPEC-060)", () => {
  // The Reference's "File names and labels" table. The XLTM and OTS rows have
  // no fixture of their own: their variant is set on a read model.
  const LABELS: ReadonlyArray<{ fixture: FixtureName; variant?: PackageVariant; badge: string; label: string }> = [
    { fixture: "mini-book.xlsx", badge: "XLSX", label: "Excel workbook (.xlsx)" },
    { fixture: "spec116-macros.xlsm", badge: "XLSM", label: "Excel macro-enabled workbook (.xlsm)" },
    { fixture: "spec116-template.xltx", badge: "XLTX", label: "Excel template (.xltx)" },
    {
      fixture: "spec116-macros.xlsm",
      variant: "macroEnabledTemplate",
      badge: "XLTM",
      label: "Excel macro-enabled template (.xltm)",
    },
    { fixture: "mini-book.ods", badge: "ODS", label: "OpenDocument Spreadsheet (.ods)" },
    { fixture: "mini-book.ods", variant: "template", badge: "OTS", label: "OpenDocument Spreadsheet template (.ots)" },
  ];

  it.each(LABELS)("shows the title, the $badge badge and the label $label", async ({ fixture, variant, badge, label }) => {
    const { store, readFixture, mountSheets } = await loadModules();
    const { model } = await readFixture(fixture);
    if (variant) {
      model.packageVariant = variant;
    }
    store.storeSpreadsheet(model);
    const { wrapper } = await mountSheets(model.id);

    expect(wrapper.get('[data-testid="document-title"]').text()).toBe(model.title);
    expect(wrapper.get('[data-testid="format-badge"]').text()).toBe(badge);
    expect(wrapper.get('[data-testid="spreadsheet-format"]').text()).toBe(label);
  });

  const DOWNLOADS: ReadonlyArray<{ fixture: FixtureName; type: string }> = [
    { fixture: "mini-book.xlsx", type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" },
    { fixture: "mini-book.ods", type: "application/vnd.oasis.opendocument.spreadsheet" },
  ];

  it.each(DOWNLOADS)(
    "Download hands back $fixture as it came in: one file, named and typed by its format, every part byte-identical in source order",
    async ({ fixture, type }) => {
      const { store, readFixture, mountSheets } = await loadModules();
      const { bytes, model } = await readFixture(fixture);
      store.storeSpreadsheet(model);
      const { wrapper } = await mountSheets(model.id);

      await wrapper.get('[data-testid="spreadsheet-download"]').trigger("click");
      await vi.waitFor(() => expect(downloads).toHaveLength(1));
      await flushPromises();
      expect(downloads).toHaveLength(1);
      const [download] = downloads;
      expect(download?.name).toBe(fixture);
      expect(download?.type).toBe(type);
      await expectSameParts(download?.bytes ?? new Uint8Array(0), bytes);
    },
  );

  it("Close returns to the library with the workbook's id in the router state", async () => {
    const { store, readFixture, mountSheets } = await loadModules();
    const { model } = await readFixture("mini-book.xlsx");
    store.storeSpreadsheet(model);
    const { wrapper, router } = await mountSheets(model.id);
    const push = vi.spyOn(router, "push");

    await wrapper.get('[data-testid="close-document"]').trigger("click");
    expect(push).toHaveBeenCalledWith({ path: "/", state: { closedDocId: model.id } });
    await vi.waitFor(() => expect(router.currentRoute.value.path).toBe("/"), { timeout: 20_000 });
    expect(router.options.history.state).toMatchObject({ closedDocId: model.id });
  });
});

describe("SpreadsheetView load states", () => {
  it("shows the not-found state for an id that names nothing", async () => {
    const { mountSheets } = await loadModules();
    const { wrapper } = await mountSheets("nope");
    expect(wrapper.find('[data-testid="spreadsheet-not-found"]').exists()).toBe(true);
    expect(wrapper.find('[data-testid="spreadsheet-grid"]').exists()).toBe(false);
    expect(wrapper.find('[data-testid="sheet-tabs"]').exists()).toBe(false);
  });

  it("replaces a text document's id with /doc/:id from its own load path", async () => {
    const { store, mountSheets } = await loadModules();
    const replaceSpies: Array<MockInstance<Router["replace"]>> = [];
    // The route is entered while "t1" names nothing — the router's guard lets
    // an unknown id through — so the hand-off below is the view's own.
    const { router } = await mountSheets("t1", (entered) => {
      store.getOrCreateDocument("t1");
      replaceSpies.push(vi.spyOn(entered, "replace"));
    });
    await vi.waitFor(() => expect(router.currentRoute.value.path).toBe("/doc/t1"), { timeout: 20_000 });
    const targets = (replaceSpies[0]?.mock.calls ?? []).map(([to]) => router.resolve(to).path);
    expect(targets.length).toBeGreaterThan(0);
    expect(new Set(targets)).toEqual(new Set(["/doc/t1"]));
  });
});

describe("SpreadsheetView sheet tabs", () => {
  it("has one tab per visible sheet — none for a hidden or very-hidden sheet", async () => {
    const { store, readFixture, mountSheets } = await loadModules();
    const { model } = await readFixture("mini-book.xlsx");
    store.storeSpreadsheet(model);
    const { wrapper } = await mountSheets(model.id);
    const tabs = wrapper.get('[data-testid="sheet-tabs"]');
    expect(tabs.attributes("role")).toBe("tablist");
    expect(tabs.findAll('[data-testid="sheet-tab"]').map((tab) => tab.attributes("role"))).toEqual(["tab", "tab"]);
    // mini-book.xlsx's third sheet, "Hidden", is state="hidden".
    expect(tabStrip(wrapper).map((tab) => tab.split(":")[0])).toEqual(["Summary", "Data"]);

    const book = bookWithHiddenSheets();
    store.storeSpreadsheet(book);
    const other = await mountSheets(book.id);
    expect(tabStrip(other.wrapper).map((tab) => tab.split(":")[0])).toEqual(["One", "Two", "Three"]);
  });

  it("opens on the file's active sheet, the only tab in the Tab order", async () => {
    const { store, readFixture, mountSheets } = await loadModules();
    const xlsx = (await readFixture("mini-book.xlsx")).model;
    store.storeSpreadsheet(xlsx);
    const first = await mountSheets(xlsx.id);
    expect(tabStrip(first.wrapper)).toEqual(["Summary:true:0", "Data:false:-1"]);
    expect(first.wrapper.get('[data-testid="spreadsheet-grid"]').attributes("aria-label")).toBe("Summary");
    // The selected tab controls the panel, which it labels.
    const selected = first.wrapper.get('[data-testid="sheet-tab"][aria-selected="true"]');
    const panel = first.wrapper.get('[role="tabpanel"]');
    expect(selected.attributes("aria-controls")).toBe(panel.attributes("id"));
    expect(panel.attributes("aria-labelledby")).toBe(selected.attributes("id"));

    // mini-book.ods: "Inventory" is the ActiveTable — and the FIRST tab,
    // because the hidden "Archive" table comes before it.
    const ods = (await readFixture("mini-book.ods")).model;
    store.storeSpreadsheet(ods);
    const second = await mountSheets(ods.id);
    expect(tabStrip(second.wrapper)).toEqual(["Inventory:true:0", "Totals:false:-1"]);
    const grid = second.wrapper.get('[data-testid="spreadsheet-grid"]');
    expect(grid.attributes("aria-label")).toBe("Inventory");
    expect(grid.get('[data-ref="A1"]').text()).toBe("Item");
  });

  it("switches the grid when a tab is clicked", async () => {
    const { store, readFixture, mountSheets } = await loadModules();
    const { model } = await readFixture("mini-book.xlsx");
    store.storeSpreadsheet(model);
    const { wrapper } = await mountSheets(model.id);
    expect(wrapper.get('[data-ref="A1"]').text()).toBe("Q3 report");

    await wrapper.findAll('[data-testid="sheet-tab"]')[1]?.trigger("click");
    expect(tabStrip(wrapper)).toEqual(["Summary:false:-1", "Data:true:0"]);
    const grid = wrapper.get('[data-testid="spreadsheet-grid"]');
    expect(grid.attributes("aria-label")).toBe("Data");
    expect(grid.get('[data-ref="A1"]').text()).toBe("ID");
    expect(grid.get('[data-ref="B2"]').text()).toBe("1.5");
  });

  it("moves selection and focus with ArrowRight/ArrowLeft/Home/End inside the tablist, wrapping at the ends", async () => {
    const { store, mountSheets } = await loadModules();
    const book = bookWithHiddenSheets();
    store.storeSpreadsheet(book);
    const { wrapper } = await mountSheets(book.id);
    const tabs = () => wrapper.findAll('[data-testid="sheet-tab"]');
    (wrapper.get('[data-testid="sheet-tab"][aria-selected="true"]').element as HTMLElement).focus();

    const steps: ReadonlyArray<readonly [key: string, lands: string]> = [
      ["ArrowRight", "Two"],
      ["ArrowRight", "Three"],
      ["ArrowRight", "One"], // wraps past the last tab
      ["ArrowLeft", "Three"], // wraps back past the first
      ["Home", "One"],
      ["End", "Three"],
      ["ArrowLeft", "Two"],
    ];
    for (const [key, lands] of steps) {
      const focused = tabs().find((tab) => tab.element === document.activeElement);
      expect(focused, `focus is on a tab before ${key}`).toBeDefined();
      await focused?.trigger("keydown", { key });
      await flushPromises();
      expect(tabStrip(wrapper).filter((tab) => tab.endsWith(":true:0")), `${key} → ${lands}`).toEqual([`${lands}:true:0`]);
      expect(document.activeElement?.textContent?.trim(), `${key} moves focus to ${lands}`).toBe(lands);
      const grid = wrapper.get('[data-testid="spreadsheet-grid"]');
      expect(grid.attributes("aria-label")).toBe(lands);
      expect(grid.get('[data-ref="A1"]').text()).toBe(lands.toLowerCase());
    }
  });

  it("switches to the next and previous visible sheet on Ctrl+PageDown and Ctrl+PageUp from the grid", async () => {
    const { store, mountSheets } = await loadModules();
    const book = bookWithHiddenSheets();
    store.storeSpreadsheet(book);
    const { wrapper } = await mountSheets(book.id);
    const grid = () => wrapper.get('[data-testid="spreadsheet-grid"]');

    // A plain PageDown is the grid's own scrolling; it never switches sheets.
    await grid().trigger("keydown", { key: "PageDown" });
    expect(grid().attributes("aria-label")).toBe("One");

    const steps: ReadonlyArray<readonly [key: string, lands: string]> = [
      ["PageDown", "Two"], // very-hidden "Secret" is skipped
      ["PageDown", "Three"], // hidden "Stash" is skipped
      ["PageUp", "Two"],
      ["PageUp", "One"],
    ];
    for (const [key, lands] of steps) {
      await grid().trigger("keydown", { key, ctrlKey: true });
      expect(grid().attributes("aria-label"), `Ctrl+${key} → ${lands}`).toBe(lands);
      expect(grid().get('[data-ref="A1"]').text()).toBe(lands.toLowerCase());
      expect(tabStrip(wrapper)).toContain(`${lands}:true:0`);
    }
  });
});

describe("SpreadsheetView read-only grid", () => {
  async function mountMiniBook() {
    const { store, readFixture, mountSheets } = await loadModules();
    const { model } = await readFixture("mini-book.xlsx");
    store.storeSpreadsheet(model);
    return (await mountSheets(model.id)).wrapper;
  }

  it("is a read-only ARIA grid of the whole sheet, headers included, named by the sheet", async () => {
    const wrapper = await mountMiniBook();
    const grid = wrapper.get('[data-testid="spreadsheet-grid"]');
    expect(grid.attributes("role")).toBe("grid");
    expect(grid.attributes("aria-readonly")).toBe("true");
    expect(grid.attributes("aria-rowcount")).toBe("1048577");
    expect(grid.attributes("aria-colcount")).toBe("16385");
    expect(grid.attributes("aria-label")).toBe("Summary");
  });

  it("letters its column headers and numbers its row headers", async () => {
    const wrapper = await mountMiniBook();
    const columns = wrapper.findAll('[role="columnheader"]');
    expect(columns.slice(0, 6).map((header) => header.text())).toEqual(["A", "B", "C", "D", "E", "F"]);
    expect(columns[0]?.attributes("aria-colindex")).toBe("2");
    const rows = wrapper.findAll('[role="rowheader"]');
    expect(rows.slice(0, 6).map((header) => header.text())).toEqual(["1", "2", "3", "4", "5", "6"]);
    expect(rows[0]?.attributes("aria-colindex")).toBe("1");
  });

  it("renders each stored cell as a sheet-cell gridcell at its address, ARIA indices 1-based past the headers", async () => {
    const wrapper = await mountMiniBook();
    const b2 = wrapper.get('[data-ref="B2"]');
    expect(b2.attributes("data-testid")).toBe("sheet-cell");
    expect(b2.attributes("role")).toBe("gridcell");
    expect(b2.attributes("aria-rowindex")).toBe("3");
    expect(b2.attributes("aria-colindex")).toBe("3");
    expect(b2.text()).toBe("Revenue");

    const cells = wrapper.findAll('[data-testid="sheet-cell"]');
    expect(cells.length).toBeGreaterThan(0);
    for (const cell of cells) {
      const at = parseA1(cell.attributes("data-ref") ?? "");
      expect(at, cell.attributes("data-ref")).not.toBeNull();
      expect(cell.attributes("role")).toBe("gridcell");
      expect(cell.attributes("aria-rowindex")).toBe(String((at?.row ?? -9) + 2));
      expect(cell.attributes("aria-colindex")).toBe(String((at?.col ?? -9) + 2));
    }
  });

  it("shows every value through the General display", async () => {
    const wrapper = await mountMiniBook();
    const shown = (ref: string) => wrapper.get(`[data-ref="${ref}"]`).element.textContent;
    expect(shown("A1")).toBe("Q3 report"); // rich text: its runs' text
    expect(shown("B4")).toBe("950.5");
    expect(shown("B10")).toBe("1.23457E+11");
    expect(shown("B11")).toBe("1.234E-10");
    expect(shown("B8")).toBe("TRUE");
    expect(shown("C8")).toBe("FALSE");
    expect(shown("B6")).toBe("#DIV/0!");
    expect(shown("B7")).toBe("45573"); // a date is its serial until SPEC-118
  });

  it("lays its cells out on the sheet's GridMetrics: <cols> widths through the MDW formula, rows at the sheet's default height", async () => {
    // SPEC-117 replaced SPEC-116's uniform 64 × 20 px lattice. mini-book's
    // Summary sheet: column A is `width="14.6328125"` (Aptos Narrow 11, MDW
    // 7 → 102 px), B–E the 64 px default, and every row the sheetFormatPr's
    // 14.5 pt (19 px).
    const wrapper = await mountMiniBook();
    const element = (ref: string) => wrapper.get(`[data-ref="${ref}"]`).element as HTMLElement;
    const left = (ref: string) => Number.parseFloat(element(ref).style.left);
    const width = (ref: string) => Number.parseFloat(element(ref).style.width);
    const top = (ref: string) => Number.parseFloat((element(ref).closest('[role="row"]') as HTMLElement).style.top);
    expect(width("A3")).toBe(102);
    expect(width("B2")).toBe(64);
    expect(left("C2") - left("B2")).toBe(64);
    expect(left("E3") - left("A3")).toBe(102 + 3 * 64);
    expect(top("B3") - top("B2")).toBe(19);
    expect(top("B11") - top("B2")).toBe(9 * 19);
  });

  it("sizes its scroll extent 100 rows past the used range, 1,000 at least, growing by 100 within one screen of the end", async () => {
    const { store, mountSheets } = await loadModules();
    const book = handBuiltBook([
      gridSheet("Tall", [
        [0, 0, "top"],
        [4_999, 0, "bottom"],
      ]),
      gridSheet("Short", [[0, 0, "only"]]),
    ]);
    store.storeSpreadsheet(book);
    const { wrapper } = await mountSheets(book.id);
    const canvasHeight = () => (wrapper.get(".grid-canvas").element as HTMLElement).style.height;
    const scrollTo = async (scrollTop: number) => {
      const grid = wrapper.get('[data-testid="spreadsheet-grid"]');
      (grid.element as HTMLElement).scrollTop = scrollTop;
      await grid.trigger("scroll");
    };
    expect(canvasHeight()).toBe(`${(4_999 + 100) * 20}px`);

    await scrollTo(1_000); // nowhere near the end
    expect(canvasHeight()).toBe(`${(4_999 + 100) * 20}px`);
    await scrollTo((4_999 + 100) * 20 - 700); // the very end
    expect(canvasHeight()).toBe(`${(4_999 + 200) * 20}px`);
    await scrollTo((4_999 + 100) * 20 - 700); // no longer within a screen of the new end
    expect(canvasHeight()).toBe(`${(4_999 + 200) * 20}px`);

    await wrapper.findAll('[data-testid="sheet-tab"]')[1]?.trigger("click");
    expect(canvasHeight()).toBe(`${1_000 * 20}px`);
  });

  it("aligns numbers right, text left, and booleans and errors centred", async () => {
    const wrapper = await mountMiniBook();
    const classes = (ref: string) => wrapper.get(`[data-ref="${ref}"]`).classes();
    expect(classes("B3")).toContain("is-number");
    expect(classes("B3")).not.toContain("is-text");
    expect(classes("B2")).toContain("is-text");
    expect(classes("B8")).toContain("is-centred");
    expect(classes("B6")).toContain("is-centred");
  });

  it("keeps fewer than 2,000 cells mounted for a million-row sheet, on open and scrolled to its last used row", async () => {
    // A 1280 × 720 grid: 700 px of rows under the 20 px letter row.
    vi.spyOn(Element.prototype, "clientWidth", "get").mockReturnValue(1280);
    vi.spyOn(Element.prototype, "clientHeight", "get").mockReturnValue(720);
    const { store, mountSheets } = await loadModules();
    const book = handBuiltBook([millionRowSheet()]);
    store.storeSpreadsheet(book);
    const { wrapper } = await mountSheets(book.id);
    const cellCount = () => wrapper.findAll('[data-testid="sheet-cell"]').length;

    expect(wrapper.get('[data-ref="A1"]').text()).toBe("0");
    expect(cellCount()).toBeLessThan(2_000);
    // The spacer stops growing at MAX_SCROLL_EXTENT_PX…
    const grid = wrapper.get('[data-testid="spreadsheet-grid"]');
    const canvas = grid.get(".grid-canvas").element as HTMLElement;
    expect(canvas.style.height).toBe("15000000px");

    // …and scroll offsets map onto rows proportionally beyond it. The extent
    // runs 100 rows past the last used row (0-based 999,999), so scroll to
    // where that row sits near the middle of the 700 px of rows.
    const rowPx = 20;
    const viewPx = 700;
    const lastUsedRow = MILLION - 1;
    const totalPx = (lastUsedRow + 100) * rowPx;
    const range = 15_000_000 - viewPx;
    const scrollTop = Math.round(((lastUsedRow * rowPx - 300) / (totalPx - viewPx)) * range);
    (grid.element as HTMLElement).scrollTop = scrollTop;
    await grid.trigger("scroll");
    await flushPromises();

    const last = wrapper.get('[data-ref="A1000000"]');
    expect(last.text()).toBe("999999");
    expect(wrapper.get('[data-ref="M1000000"]').text()).toBe("12");
    expect(cellCount()).toBeLessThan(2_000);
    // The row is placed inside the scrolled viewport, not merely mounted.
    // (SPEC-117: the grid's content box no longer scrolls — rows are placed
    // relative to the visible band, which here starts under the letter row.)
    const rowTop = Number.parseFloat((last.element.closest('[role="row"]') as HTMLElement).style.top);
    expect(rowTop).toBeGreaterThanOrEqual(0);
    expect(rowTop + rowPx).toBeLessThanOrEqual(viewPx);
  });
});

describe("SpreadsheetView foreign sheets", () => {
  const FOREIGN: ReadonlyArray<{ part: string; message: string }> = [
    { part: "xl/chartsheets/sheet1.xml", message: "This chart sheet can't be shown yet" },
    { part: "xl/dialogsheets/sheet1.xml", message: "This sheet type can't be shown yet" },
  ];

  it.each(FOREIGN)("shows a sheet-foreign panel for $part in place of the grid", async ({ part, message }) => {
    const { store, mountSheets } = await loadModules();
    const book = handBuiltBook([foreignSheet("Odd one", part), gridSheet("Data", [[0, 0, "data"]])]);
    store.storeSpreadsheet(book);
    const { wrapper } = await mountSheets(book.id);

    expect(tabStrip(wrapper)).toEqual(["Odd one:true:0", "Data:false:-1"]);
    expect(wrapper.get('[data-testid="sheet-foreign"]').text()).toBe(message);
    expect(wrapper.find('[data-testid="spreadsheet-grid"]').exists()).toBe(false);

    await wrapper.findAll('[data-testid="sheet-tab"]')[1]?.trigger("click");
    expect(wrapper.find('[data-testid="sheet-foreign"]').exists()).toBe(false);
    expect(wrapper.get('[data-testid="spreadsheet-grid"]').get('[data-ref="A1"]').text()).toBe("data");
  });
});

describe("SpreadsheetView reader versions", () => {
  it("re-reads a workbook stored by an older reader on open, keeping its id, title and sheet ids", async () => {
    const { store, readFixture, mountSheets } = await loadModules();
    const { model } = await readFixture("mini-book.xlsx");
    // A version-0 record whose rows were never extracted: only the retained
    // source package can bring them back.
    const stale: SpreadsheetModel = {
      ...model,
      readerVersion: 0,
      sheets: model.sheets.map((sheet) => ({ ...sheet, rows: [] })),
    };
    store.storeSpreadsheet(stale);
    const { wrapper } = await mountSheets(stale.id);

    // SPEC-117 mounts every position in the window, so B2 exists (empty)
    // before the upgrade lands: wait for its value.
    await vi.waitFor(() => expect(wrapper.find('[data-ref="B2"]').text()).toBe("Revenue"), { timeout: 10_000 });
    const upgraded = store.getSpreadsheet(stale.id);
    expect(upgraded?.readerVersion).toBe(SPREADSHEET_READER_VERSION);
    expect(upgraded?.id).toBe(stale.id);
    expect(upgraded?.title).toBe("mini-book");
    expect(upgraded?.sheets.map((sheet) => sheet.id)).toEqual(stale.sheets.map((sheet) => sheet.id));
    expect(upgraded?.sheets.map((sheet) => sheet.rows.length)).toEqual(model.sheets.map((sheet) => sheet.rows.length));
  });
});
