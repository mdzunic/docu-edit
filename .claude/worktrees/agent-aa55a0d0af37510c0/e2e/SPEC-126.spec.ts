import { test, expect, type Page } from "@playwright/test";
import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { equalBytes } from "../src/lib/core/equality";

// SPEC-126 — CSV and TSV (AC-76..82). Drop spec126-semicolon-1252.csv (a
// German Excel export) → a library row with a CSV badge → open it: the import
// bar reads "Opened as semicolon-separated, Windows-1252, numbers and dates
// read as German · Change…", A2 reads "Müller", B2 "01234" left-aligned (a
// leading zero stays text), C2 "1234.5" right-aligned (read in German, shown
// in the user's locale) → Change… → Comma → Apply: A1 reads the whole header
// line → Change… → Semicolon → Apply → download: the bytes are the fixture's →
// type 99 into C2 → download: the fixture with exactly 1234,5 replaced by 99 →
// type =1+1 into D2 → download: the losses toast. Then mini-book.xlsx → the
// download menu's "Download a copy as CSV (current sheet)": a BOM, CRLF
// records, and one record holding "Revenue".

const fixturesDir = join(dirname(fileURLToPath(import.meta.url)), "fixtures");
const csvFixture = join(fixturesDir, "spec126-semicolon-1252.csv");
const xlsxFixture = join(fixturesDir, "mini-book.xlsx");

const CSV_LOSS_NOTICE = "Formatting, formulas and other sheets aren't saved in CSV — download a copy as .xlsx to keep them";

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

/** Clicks the centre of a cell — an empty cell has no element of its own,
 * so the point comes from its column and row headers. */
async function clickCell(page: Page, ref: string): Promise<void> {
  const [, letters, digits] = /^([A-Z]+)(\d+)$/.exec(ref)!;
  const column = await page.getByTestId("sheet-col-header").filter({ hasText: new RegExp(`^${letters}$`) }).boundingBox();
  const row = await page.getByTestId("sheet-row-header").filter({ hasText: new RegExp(`^${digits}$`) }).boundingBox();
  await page.mouse.click(column!.x + column!.width / 2, row!.y + row!.height / 2);
  await expect(page.getByTestId("name-box")).toHaveValue(ref);
}

/** Types an entry into a cell and commits it with Enter. */
async function enter(page: Page, ref: string, text: string): Promise<void> {
  await clickCell(page, ref);
  await page.keyboard.type(text);
  await page.keyboard.press("Enter");
}

/** Downloads through `trigger`; resolves to the file's bytes. */
async function downloadBytes(page: Page, trigger: () => Promise<void>): Promise<Uint8Array> {
  const downloadPromise = page.waitForEvent("download");
  await trigger();
  const download = await downloadPromise;
  return new Uint8Array(readFileSync(await download.path()));
}

/** Change… → choose a separator → Apply. */
async function reimport(page: Page, separator: string): Promise<void> {
  await page.getByTestId("csv-import-change").click();
  const dialog = page.getByTestId("csv-import-dialog");
  await expect(dialog).toBeVisible();
  await expect(dialog.getByTestId("csv-import-preview")).toBeVisible();
  await dialog.getByTestId("csv-import-separator").selectOption({ label: separator });
  await dialog.getByTestId("csv-import-apply").click();
  await expect(dialog).toHaveCount(0);
}

test("a German CSV opens, re-imports, downloads byte-for-byte and saves in its own dialect; any sheet copies as CSV (AC-76..82)", async ({ page }) => {
  const source = new Uint8Array(readFileSync(csvFixture));

  await page.goto("/");
  await expect(page.getByTestId("app-bar")).toBeVisible();
  await dropFile(page, "spec126-semicolon-1252.csv", "text/csv", csvFixture);
  await expect(page).toHaveURL(/\/sheets\/.+/);
  await expect(page.getByTestId("spreadsheet-format")).toHaveText("Comma-separated values (.csv)");

  // A library row with a CSV badge; open it.
  await page.getByTestId("close-document").click();
  await expect(page.getByTestId("library")).toBeVisible();
  await expect(rowByBadge(page, "CSV")).toHaveCount(1);
  await rowByBadge(page, "CSV").dblclick();
  await expect(page).toHaveURL(/\/sheets\/.+/);

  await expect(page.getByTestId("csv-import-bar")).toHaveText(
    "Opened as semicolon-separated, Windows-1252, numbers and dates read as German · Change…",
  );
  await expect(cell(page, "A2")).toHaveText("Müller");
  await expect(cell(page, "B2")).toHaveText("01234");
  await expect(cell(page, "B2")).toHaveCSS("text-align", "left");
  await expect(cell(page, "C2")).toHaveText("1234.5");
  await expect(cell(page, "C2")).toHaveCSS("text-align", "right");

  // Re-read the retained bytes split on commas: the header line is one field.
  await reimport(page, "Comma");
  await expect(cell(page, "A1")).toHaveText("Name;Zip;Amount");
  await expect(page.getByTestId("csv-import-bar")).toContainText("Opened as comma-separated");

  // Back to semicolons: still unedited, so the download is the source bytes.
  await reimport(page, "Semicolon");
  await expect(cell(page, "A1")).toHaveText("Name");
  const untouched = await downloadBytes(page, () => page.getByTestId("spreadsheet-download").click());
  expect(equalBytes(untouched, source)).toBe(true);

  // One field edited: the file in its own dialect, that field replaced.
  await enter(page, "C2", "99");
  await expect(cell(page, "C2")).toHaveText("99");
  const edited = await downloadBytes(page, () => page.getByTestId("spreadsheet-download").click());
  const expected = Buffer.from(Buffer.from(source).toString("latin1").replace("1234,5", "99"), "latin1");
  expect(equalBytes(edited, new Uint8Array(expected))).toBe(true);
  await expect(page.getByTestId("toast").filter({ hasText: CSV_LOSS_NOTICE })).toHaveCount(0);

  // A formula is something CSV cannot carry: the download still happens, with
  // the losses toast.
  await enter(page, "D2", "=1+1");
  await expect(cell(page, "D2")).toHaveText("2");
  const lossy = await downloadBytes(page, () => page.getByTestId("spreadsheet-download").click());
  expect(lossy.length).toBeGreaterThan(0);
  await expect(page.getByTestId("toast").filter({ hasText: CSV_LOSS_NOTICE })).toHaveCount(1);

  // Any workbook: the current sheet as CSV UTF-8.
  await page.goto("/");
  await page.getByTestId("open-document-input").setInputFiles(xlsxFixture);
  await expect(page).toHaveURL(/\/sheets\/.+/);
  await expect(page.getByTestId("csv-import-bar")).toHaveCount(0);
  await page.getByTestId("spreadsheet-download-menu").click();
  const copy = await downloadBytes(page, () => page.getByTestId("spreadsheet-download-copy-csv").click());
  expect([...copy.subarray(0, 3)]).toEqual([0xef, 0xbb, 0xbf]);
  const text = new TextDecoder().decode(copy.subarray(3));
  expect(text.endsWith("\r\n")).toBe(true);
  const records = text.slice(0, -2).split("\r\n");
  expect(records.every((record) => !record.includes("\n") || record.includes('"'))).toBe(true);
  expect(records.some((record) => record.includes("Revenue"))).toBe(true);
});
