import { test, expect, type Page } from "@playwright/test";
import { existsSync } from "node:fs";
import { execFileSync } from "node:child_process";
import path from "node:path";
import { fileURLToPath } from "node:url";

// SPEC-137 — large workbooks: functional behaviour (AC-61..64), in the
// normal `chromium` project — no timing assertions (those live in
// e2e/SPEC-137.perf.spec.ts). Open spec137-library.xlsx (100,000 cells)
// through the library picker, scroll to a distant row and check the grid's
// rendering ceiling holds, zoom out until the overview canvas takes over,
// zoom back in, edit a cell, reload and confirm both the edit and the
// chunked storage layout survived, then reopen from the library and check
// the storage panel reports a real size.
//
// Several of the features this spec exercises are being built by other
// SPEC-137 workstreams in parallel (sheet-overview-canvas, data-calc-state,
// the chunked storage layout, lazy residency, the open-progress indicator) —
// this file is written against the shared CONTRACT's names, not against
// what happens to be implemented in this worktree today.

const REPO_ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), "..");
const LARGE_FIXTURE_PATH = path.join(REPO_ROOT, "e2e", "fixtures", "spec137-large.xlsx");
const LIBRARY_FIXTURE_PATH = path.join(REPO_ROOT, "e2e", "fixtures", "spec137-library.xlsx");

function buildFixturesIfMissing(): void {
  if (existsSync(LARGE_FIXTURE_PATH) && existsSync(LIBRARY_FIXTURE_PATH)) {
    return;
  }
  execFileSync("npx", ["vite-node", "scripts/build-spec137-fixtures.ts"], { cwd: REPO_ROOT, stdio: "inherit" });
  if (!existsSync(LARGE_FIXTURE_PATH) || !existsSync(LIBRARY_FIXTURE_PATH)) {
    throw new Error("SPEC-137 fixture build did not produce both e2e/fixtures/spec137-*.xlsx files");
  }
}

function cell(page: Page, ref: string) {
  return page.locator(`[data-testid="sheet-cell"][data-ref="${ref}"]`);
}

async function waitForFirstPaintedCell(page: Page, timeout = 30_000): Promise<void> {
  await page.waitForFunction(
    () => {
      const grid = document.querySelector('[data-testid="spreadsheet-grid"]');
      if (!grid) {
        return false;
      }
      for (const el of grid.querySelectorAll(".sheet-cell")) {
        if ((el.textContent ?? "").trim().length > 0) {
          return true;
        }
      }
      return false;
    },
    { timeout },
  );
}

async function setZoom(page: Page, percent: number): Promise<void> {
  const field = page.getByTestId("zoom-percent");
  await field.click();
  await field.fill(String(percent));
  await field.press("Enter");
}

interface StorageCheck {
  hasChunkRecord: boolean;
  shellSheetsRowsEmpty: boolean;
}

/** Reads the "documents" IndexedDB store directly (database "docu-edit"):
 * whether at least one `sheetrows:<docId>:<sheetId>:<4-digit chunk>` chunk
 * record exists, and whether the shell record's `model.sheets` all have
 * `rows: []` (SPEC-137's storage layout — the shell holds no cell byte). */
async function storageCheck(page: Page, docId: string): Promise<StorageCheck> {
  return page.evaluate(async (id) => {
    const db = await new Promise<IDBDatabase>((resolve, reject) => {
      const request = indexedDB.open("docu-edit");
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error);
    });
    const store = db.transaction("documents").objectStore("documents");
    const keys = await new Promise<IDBValidKey[]>((resolve, reject) => {
      const request = store.getAllKeys();
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error);
    });
    const shell = await new Promise<unknown>((resolve, reject) => {
      const request = store.get(id);
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error);
    });
    db.close();
    const chunkPattern = new RegExp(`^sheetrows:${id}:[^:]+:\\d{4}$`);
    const hasChunkRecord = keys.map(String).some((key) => chunkPattern.test(key));
    const sheets = (shell as { model?: { sheets?: { rows?: unknown[] }[] } } | undefined)?.model?.sheets;
    const shellSheetsRowsEmpty = Array.isArray(sheets) && sheets.length > 0 && sheets.every((sheet) => Array.isArray(sheet.rows) && sheet.rows.length === 0);
    return { hasChunkRecord, shellSheetsRowsEmpty };
  }, docId);
}

test.beforeAll(() => {
  buildFixturesIfMissing();
});

test("a large workbook stays within the grid's rendering ceiling, degrades to the overview at low zoom, and persists in chunks (AC-61..64)", async ({ page }) => {
  test.setTimeout(60_000);

  await page.goto("/");
  await page.getByTestId("open-document-input").waitFor({ state: "attached" });
  const navigated = page.waitForURL(/\/sheets\/.+/, { timeout: 30_000 });
  await page.getByTestId("open-document-input").setInputFiles(LIBRARY_FIXTURE_PATH);
  await navigated;
  await waitForFirstPaintedCell(page);
  const docId = new URL(page.url()).pathname.split("/").pop()!;

  // Scroll to row 4,000 (the name box's Go To): the rendering ceiling holds
  // even at a distant row (AC-61).
  await page.getByTestId("name-box").fill("A4000");
  await page.keyboard.press("Enter");
  await expect(cell(page, "A4000")).toBeVisible();
  const sheetCellCount = await page.locator(".sheet-cell").count();
  expect(sheetCellCount, "AC-61: GRID_MAX_RENDERED_CELLS").toBeLessThanOrEqual(5000);
  const gridNodeCount = await page.locator('[data-testid="spreadsheet-grid"] *').count();
  expect(gridNodeCount, "AC-61: GRID_MAX_DOM_NODES").toBeLessThanOrEqual(12000);

  // Zoom to 10%: the overview canvas takes over, and only the active cell
  // keeps a DOM cell element (AC-62).
  await setZoom(page, 10);
  await expect(page.getByTestId("sheet-overview-canvas")).toBeVisible();
  await expect(page.locator(".sheet-cell")).toHaveCount(1);

  // Back to 100%, then edit B2 (AC-63).
  await setZoom(page, 100);
  await page.getByTestId("name-box").fill("B2");
  await page.keyboard.press("Enter");
  await page.keyboard.type("42");
  await page.keyboard.press("Enter");
  await expect(cell(page, "B2")).toHaveText("42");

  // Reload: not resident any more. The edit survived, and storage holds the
  // chunk records plus a shell with no cell bytes (AC-63).
  await page.reload();
  await waitForFirstPaintedCell(page);
  await expect(cell(page, "B2")).toHaveText("42");

  const storage = await storageCheck(page, docId);
  expect(storage.hasChunkRecord, "AC-63: a sheetrows:<docId>:<sheetId>:<chunk4> record exists").toBe(true);
  expect(storage.shellSheetsRowsEmpty, "AC-63: the shell's sheets all have rows: []").toBe(true);

  // Open it from the library again: it renders (AC-64).
  await page.getByTestId("close-document").click();
  await expect(page.getByTestId("library")).toBeVisible();
  await page.getByTestId("library-row").first().dblclick();
  await expect(page).toHaveURL(new RegExp(`/sheets/${docId}$`));
  await waitForFirstPaintedCell(page);
  await expect(cell(page, "B2")).toHaveText("42");

  // The library storage panel lists it with a non-zero size (AC-64).
  await page.getByTestId("close-document").click();
  await expect(page.getByTestId("library")).toBeVisible();
  await expect(page.getByTestId("storage-panel")).toBeVisible();
  const largestDocText = await page.getByTestId("largest-doc").first().textContent();
  expect(largestDocText ?? "", "AC-64: the storage panel lists a non-zero size").toMatch(/[1-9]/);
});
