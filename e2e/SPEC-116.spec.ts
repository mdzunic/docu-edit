import { test, expect, type Page } from "@playwright/test";
import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { equalBytes } from "../src/lib/core/equality";
import { unzip } from "../src/lib/core/zip";
import { zip } from "../src/lib/core/zipWriter";

// SPEC-116 — workbooks in the library: the full journey (AC-82..84). Drop
// mini-book.xlsx (Excel-shaped) → its library row carries an XLSX badge and
// its card reads "3 sheets" → open it: B2 reads "Revenue", a number sits
// right-aligned → the second tab shows its own values → Close → drop
// mini-book.ods (LibreOffice-shaped) → it opens on the file's active sheet,
// the first TAB (the hidden first table has none) → Download returns every
// part byte-identical → /doc/<the workbook id> lands on /sheets/<id> → the
// ODS chip filters the library to that one row.

const fixturesDir = join(dirname(fileURLToPath(import.meta.url)), "fixtures");
const xlsxFixture = join(fixturesDir, "mini-book.xlsx");
const odsFixture = join(fixturesDir, "mini-book.ods");

const XLSX_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
const ODS_TYPE = "application/vnd.oasis.opendocument.spreadsheet";

async function dropFile(page: Page, name: string, mime: string, path: string) {
  const transfer = await page.evaluateHandle(
    (item) => {
      const dt = new DataTransfer();
      const bytes = Uint8Array.from(atob(item.base64), (c) => c.charCodeAt(0));
      dt.items.add(new File([bytes], item.name, { type: item.mime }));
      return dt;
    },
    { name, mime, base64: readFileSync(path).toString("base64") },
  );
  await page.dispatchEvent("body", "drop", { dataTransfer: transfer });
}

function rowByBadge(page: Page, badge: string) {
  return page.getByTestId("library-row").filter({ has: page.getByTestId("row-format").getByText(badge, { exact: true }) });
}

function cell(page: Page, ref: string) {
  return page.locator(`[data-testid="sheet-cell"][data-ref="${ref}"]`);
}

/** Whether the workbook's shell and all its sheet-rows records have reached
 * IndexedDB — the write-through runs at idle, and the journey reloads. */
async function storedWithRows(page: Page, id: string, sheets: number): Promise<boolean> {
  return page.evaluate(
    async ({ id, sheets }) => {
      const db = await new Promise<IDBDatabase>((resolve, reject) => {
        // No explicit version: the app has already opened the DB at its
        // current schema version by this point.
        const request = indexedDB.open("docu-edit");
        request.onsuccess = () => resolve(request.result);
        request.onerror = () => reject(request.error);
      });
      const keys = await new Promise<IDBValidKey[]>((resolve, reject) => {
        const request = db.transaction("documents").objectStore("documents").getAllKeys();
        request.onsuccess = () => resolve(request.result);
        request.onerror = () => reject(request.error);
      });
      db.close();
      const names = keys.map(String);
      return names.includes(id) && names.filter((key) => key.startsWith(`sheetrows:${id}:`)).length === sheets;
    },
    { id, sheets },
  );
}

test("workbooks are first-class: drop, open, tabs, download, kind guard, filter (AC-82..84)", async ({ page }) => {
  await page.goto("/");
  await expect(page.getByTestId("app-bar")).toBeVisible();

  // Drop the Excel workbook: every open path funnels through
  // openAnyDocumentFile, so the drop opens the workbook view.
  await dropFile(page, "mini-book.xlsx", XLSX_TYPE, xlsxFixture);
  await expect(page).toHaveURL(/\/sheets\/.+/);
  await expect(page.getByTestId("spreadsheet-grid")).toBeVisible();
  await expect(page.getByTestId("spreadsheet-format")).toHaveText("Excel workbook (.xlsx)");

  // Close (SPEC-060) → a library row with an XLSX badge; its card reads
  // "3 sheets" (the hidden one counts).
  await page.getByTestId("close-document").click();
  await expect(page.getByTestId("library")).toBeVisible();
  await expect(rowByBadge(page, "XLSX")).toHaveCount(1);
  await page.getByTestId("view-grid").click();
  await expect(rowByBadge(page, "XLSX").getByTestId("row-counts")).toHaveText("3 sheets");
  await expect(rowByBadge(page, "XLSX").getByTestId("sheet-miniature")).toBeVisible();

  // Open it: B2's cell reads "Revenue", and a number cell is right-aligned.
  await rowByBadge(page, "XLSX").dblclick();
  await expect(page).toHaveURL(/\/sheets\/.+/);
  await expect(cell(page, "B2")).toHaveText("Revenue");
  await expect(cell(page, "B3")).toHaveText("1200");
  await expect(cell(page, "B3")).toHaveCSS("text-align", "right");
  await expect(cell(page, "A3")).toHaveCSS("text-align", "left");
  const grid = page.getByTestId("spreadsheet-grid");
  await expect(grid).toHaveAttribute("role", "grid");
  await expect(grid).toHaveAttribute("aria-label", "Summary");

  // Two tabs — the hidden third sheet has none; the second shows its values.
  const tabs = page.getByTestId("sheet-tab");
  await expect(tabs).toHaveText(["Summary", "Data"]);
  await expect(tabs.nth(0)).toHaveAttribute("aria-selected", "true");
  await tabs.nth(1).click();
  await expect(tabs.nth(1)).toHaveAttribute("aria-selected", "true");
  await expect(grid).toHaveAttribute("aria-label", "Data");
  await expect(cell(page, "A1")).toHaveText("ID");
  await expect(cell(page, "B3")).toHaveText("3");
  await expect(cell(page, "C3")).toHaveText("even");

  // Close → drop the LibreOffice workbook. It opens on the file's active
  // sheet ("Inventory"), which is the first tab: the hidden "Archive" table
  // before it has none.
  await page.getByTestId("close-document").click();
  await expect(page.getByTestId("library")).toBeVisible();
  await dropFile(page, "mini-book.ods", ODS_TYPE, odsFixture);
  await expect(page).toHaveURL(/\/sheets\/.+/);
  const odsId = new URL(page.url()).pathname.split("/").pop()!;
  await expect(page.getByTestId("spreadsheet-format")).toHaveText("OpenDocument Spreadsheet (.ods)");
  await expect(tabs).toHaveText(["Inventory", "Totals"]);
  await expect(tabs.first()).toHaveAttribute("aria-selected", "true");
  await expect(cell(page, "A1")).toHaveText("Item");
  await expect(cell(page, "A2")).toHaveText("Widget");

  // Download: every source part comes back byte-identical, in order (the ZIP
  // container itself is rebuilt), mimetype first.
  const downloadPromise = page.waitForEvent("download");
  await page.getByTestId("spreadsheet-download").click();
  const download = await downloadPromise;
  expect(download.suggestedFilename()).toBe("mini-book.ods");
  const saved = await unzip(new Uint8Array(readFileSync(await download.path())));
  const source = await unzip(new Uint8Array(readFileSync(odsFixture)));
  expect([...saved.keys()]).toEqual([...source.keys()]);
  expect([...saved.keys()][0]).toBe("mimetype");
  for (const [name, bytes] of source) {
    expect(equalBytes(saved.get(name)!, bytes), name).toBe(true);
  }

  // /doc/<the workbook id> — a full load, so the record comes back from
  // IndexedDB with its rows rejoined — lands on /sheets/<id>.
  await expect.poll(() => storedWithRows(page, odsId, 3)).toBe(true);
  await page.goto(`/doc/${odsId}`);
  await expect(page).toHaveURL(new RegExp(`/sheets/${odsId}$`));
  await expect(cell(page, "A2")).toHaveText("Widget");

  // Filtering by the ODS chip shows exactly the one row.
  await page.getByTestId("close-document").click();
  await expect(page.getByTestId("library")).toBeVisible();
  await page.getByTestId("filter-ods").click();
  await expect(page.getByTestId("library-row")).toHaveCount(1);
  await expect(page.getByTestId("library-row").getByTestId("row-format")).toHaveText("ODS");
});

/** A 1,000,000-row workbook in a few hundred bytes: LibreOffice's repeat
 * attributes expand into 999,950 one-cell rows and then 50 rows of 30 cells
 * (1,001,450 cells, under the 2,000,000 ceiling) — so the last used row's
 * neighbourhood is as full as a viewport can get. */
async function millionRowOds(): Promise<Uint8Array> {
  const encoder = new TextEncoder();
  const cell = (value: number, repeat = 1) =>
    `<table:table-cell office:value-type="float" office:value="${value}"${repeat > 1 ? ` table:number-columns-repeated="${repeat}"` : ""}/>`;
  const content =
    `<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" ` +
    `xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" office:version="1.3"><office:body><office:spreadsheet>` +
    `<table:table table:name="Million"><table:table-row table:number-rows-repeated="999950">${cell(1)}</table:table-row>` +
    `<table:table-row table:number-rows-repeated="50">${cell(7, 30)}</table:table-row></table:table>` +
    `</office:spreadsheet></office:body></office:document-content>`;
  return zip([
    { name: "mimetype", data: encoder.encode("application/vnd.oasis.opendocument.spreadsheet"), store: true },
    { name: "content.xml", data: encoder.encode(content) },
  ]);
}

test("a 1,000,000-row sheet scrolled to its last used row keeps fewer than 2,000 cells mounted (AC-69)", async ({ page }) => {
  await page.goto("/");
  await expect(page.getByTestId("app-bar")).toBeVisible();
  const transfer = await page.evaluateHandle(
    (base64) => {
      const dt = new DataTransfer();
      const bytes = Uint8Array.from(atob(base64), (c) => c.charCodeAt(0));
      dt.items.add(new File([bytes], "million.ods", { type: "application/vnd.oasis.opendocument.spreadsheet" }));
      return dt;
    },
    Buffer.from(await millionRowOds()).toString("base64"),
  );
  await page.dispatchEvent("body", "drop", { dataTransfer: transfer });
  await expect(page).toHaveURL(/\/sheets\/.+/);
  const grid = page.getByTestId("spreadsheet-grid");
  await expect(cell(page, "A1")).toHaveText("1");

  // Scroll to the last used row (1,000,000). The extent runs past it, and
  // above 15,000,000 px the scroll offset maps onto rows proportionally — so
  // aim the viewport's middle at that row through the same mapping.
  await grid.evaluate((element) => {
    const rowHeight = 20;
    const view = element.clientHeight - rowHeight;
    const canvas = parseFloat((element.querySelector(".grid-canvas") as HTMLElement).style.height);
    const totalPx = (999_999 + 100) * rowHeight;
    const virtual = 999_999 * rowHeight - view / 2;
    element.scrollTop = Math.round((virtual / (totalPx - view)) * (canvas - view));
  });
  await expect(cell(page, "A1000000")).toBeVisible();
  await expect(cell(page, "P1000000")).toHaveText("7");
  expect(await page.getByTestId("sheet-cell").count()).toBeLessThan(2000);
  await expect(page.locator('[role="rowheader"]').filter({ hasText: /^1000000$/ })).toBeVisible();
});
