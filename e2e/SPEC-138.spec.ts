import { test, expect, type Locator, type Page } from "@playwright/test";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

// SPEC-138 — autosave, drafts and version history for workbooks, as one
// session (AC-63): open mini-book.xlsx, type into B3 on "Data", reload — the
// value is there with no recovery bar; plant a newer draft the way a crash
// leaves one (SPEC-010's page.evaluate + indexedDB technique: a `draft:<id>`
// shell whose `rowKeys` manifest names one draft chunk record, keyed by
// draftRowsKey, and the committed chunks for the rest) and reopen — the bar
// shows over the drafted cell, the name box reads the drafted active cell, and
// Discard brings the committed value back; edit and download (a commit-point
// snapshot), open the history pane, name that version, preview it (the grid
// read-only), leave the preview, edit again and restore the named version —
// the grid shows the restored value and the list gains the restore's two
// snapshots; then a second tab on the same workbook is read-only under the
// conflict bar until the first closes, when it takes over and typing works.

const FIXTURE = join(dirname(fileURLToPath(import.meta.url)), "fixtures", "mini-book.xlsx");
const SAVED = "All changes saved";

function cell(page: Page, ref: string): Locator {
  return page.locator(`[data-testid="sheet-cell"][data-ref="${ref}"]`);
}

function tab(page: Page, name: string): Locator {
  return page.getByTestId("sheet-tab").filter({ hasText: new RegExp(`^\\s*${name}\\s*$`) });
}

/** Clicks the centre of a cell from its column and row headers. */
async function clickCell(page: Page, ref: string): Promise<void> {
  const [, letters, digits] = /^([A-Z]+)(\d+)$/.exec(ref)!;
  const column = await page.getByTestId("sheet-col-header").filter({ hasText: new RegExp(`^${letters}$`) }).boundingBox();
  const row = await page.getByTestId("sheet-row-header").filter({ hasText: new RegExp(`^${digits}$`) }).boundingBox();
  await page.mouse.click(column!.x + column!.width / 2, row!.y + row!.height / 2);
  await expect(page.getByTestId("name-box")).toHaveValue(ref);
}

async function showData(page: Page): Promise<void> {
  await tab(page, "Data").click();
  await expect(tab(page, "Data")).toHaveAttribute("aria-selected", "true");
}

/** Types an entry into a cell of the sheet on show and commits it. */
async function enterAt(page: Page, ref: string, text: string): Promise<void> {
  await clickCell(page, ref);
  await page.keyboard.type(text);
  await page.keyboard.press("Enter");
  await expect(cell(page, ref)).toHaveText(text);
}

async function waitForSaved(page: Page): Promise<void> {
  await expect(page.getByTestId("save-state")).toHaveText(SAVED, { timeout: 15_000 });
  await expect(page.getByTestId("save-state")).toHaveClass(/sheet-save-state/);
}

/** Writes a draft newer than the committed workbook straight into the app's
 * store: the committed shell read back, its manifest's keys listed in a
 * `rowKeys` manifest — committed chunk keys, except "Data"'s first chunk,
 * whose rows are copied under its draft key with B3 rewritten — and a view
 * that opens on "Data" with C4 active. What an unclean close leaves behind. */
async function plantNewerDraft(page: Page, id: string, value: string): Promise<void> {
  await page.evaluate(
    async ([docId, drafted]) => {
      const db = await new Promise<IDBDatabase>((resolve, reject) => {
        const request = indexedDB.open("docu-edit"); // the app already created it at its current version
        request.onsuccess = () => resolve(request.result);
        request.onerror = () => reject(request.error);
      });
      const read = (key: string) =>
        new Promise<Record<string, unknown>>((resolve, reject) => {
          const request = db.transaction("documents", "readonly").objectStore("documents").get(key);
          request.onsuccess = () => resolve(request.result as Record<string, unknown>);
          request.onerror = () => reject(request.error);
        });
      type Shell = { model: { sheets: { id: string; name: string }[] }; sheetStore: { chunks: Record<string, number[]> } };
      const committed = (await read(docId)) as unknown as Shell;
      const dataSheet = committed.model.sheets.find((sheet) => sheet.name === "Data")!;
      const chunkKey = (sheetId: string, chunk: number) => `sheetrows:${docId}:${sheetId}:${String(chunk).padStart(4, "0")}`;
      const rowKeys: Record<string, string[]> = {};
      for (const sheet of committed.model.sheets) {
        rowKeys[sheet.id] = (committed.sheetStore.chunks[sheet.id] ?? []).map((chunk) => chunkKey(sheet.id, chunk));
      }
      const first = chunkKey(dataSheet.id, 0);
      const draftKey = `draft${first}`;
      rowKeys[dataSheet.id] = [draftKey, ...rowKeys[dataSheet.id]!.slice(1)];
      const chunk = (await read(first)) as { rows: { index: number; cells: { col: number; value: unknown }[] }[] };
      const rows = chunk.rows.map((row) =>
        row.index === 2 ? { ...row, cells: row.cells.map((c) => (c.col === 1 ? { col: 1, value: drafted } : c)) } : row,
      );
      const draft = {
        id: `draft:${docId}`,
        docId,
        savedAt: Date.now() + 1000,
        model: committed.model,
        caret: null,
        scroll: 0,
        rowKeys,
        sourceKey: `sheetsource:${docId}`,
        sheetView: { sheetId: dataSheet.id, active: { row: 3, col: 2 }, topRow: 0, leftCol: 0 },
      };
      await new Promise<void>((resolve, reject) => {
        const tx = db.transaction("documents", "readwrite");
        tx.objectStore("documents").put({ ...chunk, id: draftKey, rows });
        tx.objectStore("documents").put(draft);
        tx.oncomplete = () => resolve();
        tx.onerror = () => reject(tx.error);
      });
      db.close();
    },
    [id, value] as const,
  );
}

test("a workbook autosaves, recovers a crash's draft, snapshots, names, previews and restores its versions, and locks a second tab", async ({ page }) => {
  await page.goto("/");
  await page.getByTestId("open-document-input").setInputFiles(FIXTURE);
  await expect(page).toHaveURL(/\/sheets\/.+/);
  const id = page.url().match(/\/sheets\/([^/?#]+)/)?.[1] ?? "";
  expect(id).not.toBe("");

  // Type into B3 on "Data"; the bar reads Saving… then All changes saved.
  await showData(page);
  await expect(cell(page, "B3")).toHaveText("3");
  await enterAt(page, "B3", "777");
  await waitForSaved(page);
  await expect(page.getByTestId("save-state")).toHaveAttribute("data-save-state", "saved");

  // Reload: the value is there, and nothing asks to be recovered.
  await page.reload();
  await showData(page);
  await expect(cell(page, "B3")).toHaveText("777");
  await expect(page.getByTestId("recovery-bar")).toBeHidden();
  await waitForSaved(page);

  // A crash leaves a draft newer than the committed workbook. Plant one from
  // the library (closing the workbook cleanly first), then reopen it.
  await page.goto("/");
  await expect(page.getByTestId("library")).toBeVisible();
  await plantNewerDraft(page, id, "4242");
  await page.goto(`/sheets/${id}`);
  const bar = page.getByTestId("recovery-bar");
  await expect(bar).toBeVisible();
  await expect(bar).toContainText("Recovered unsaved changes from");
  await expect(page.getByTestId("save-state")).toHaveText("Not saved yet");
  await expect(tab(page, "Data")).toHaveAttribute("aria-selected", "true");
  await expect(cell(page, "B3")).toHaveText("4242");
  await expect(page.getByTestId("name-box")).toHaveValue("C4");

  // Discard: the draft goes and the committed value returns.
  await page.getByTestId("recovery-discard").click();
  await expect(bar).toBeHidden();
  await expect(cell(page, "B3")).toHaveText("777");

  // Edit, then download: a commit point, so a snapshot.
  await enterAt(page, "B4", "888");
  const download = page.waitForEvent("download");
  await page.getByTestId("spreadsheet-download").click();
  await download;

  // The history pane lists it; name it.
  await page.getByTestId("sheet-history-toggle").click();
  const panel = page.getByTestId("sheet-history-panel");
  await expect(panel).toBeVisible();
  const versions = panel.getByTestId("history-version");
  await expect(versions).toHaveCount(1);
  await expect(versions.first()).toContainText("3 sheets ·");
  await versions.first().getByTestId("history-name-input").fill("Before rewrite");
  await versions.first().getByTestId("history-name-submit").click();
  const named = versions.filter({ hasText: "Before rewrite" });
  await expect(named.getByTestId("history-name")).toHaveText("Before rewrite");

  // Preview it: the grid shows the version, read-only; then back.
  await named.getByTestId("history-preview").click();
  const banner = panel.getByTestId("history-preview-banner");
  await expect(banner).toContainText("Previewing an earlier version, read-only");
  await expect(page.getByTestId("spreadsheet-grid")).toHaveAttribute("aria-readonly", "true");
  await expect(cell(page, "B4")).toHaveText("888");
  await panel.getByTestId("history-exit-preview").click();
  await expect(banner).toBeHidden();
  await expect(page.getByTestId("spreadsheet-grid")).not.toHaveAttribute("aria-readonly", "true");

  // Edit again, then restore the named version.
  await enterAt(page, "B4", "999");
  await named.getByTestId("history-restore").click();
  await expect(cell(page, "B4")).toHaveText("888");
  await expect(page.getByTestId("toast-action")).toHaveText("Undo restore");
  await expect(versions).toHaveCount(3);
  await expect(panel.getByTestId("history-trigger")).toHaveText(["After restore", "Before restore"]);
  await waitForSaved(page);

  // A second tab on the same workbook is read-only.
  const second = await page.context().newPage();
  await second.goto(`/sheets/${id}`);
  const conflict = second.getByTestId("sheet-conflict-bar");
  await expect(conflict).toHaveText("This workbook is open in another tab");
  await expect(second.getByTestId("spreadsheet-grid")).toHaveAttribute("aria-readonly", "true");
  await expect(second.getByTestId("save-state")).toHaveCount(0);
  await showData(second);
  await clickCell(second, "B5");
  const before = await cell(second, "B5").textContent();
  await second.keyboard.type("55");
  await second.keyboard.press("Enter");
  await expect(cell(second, "B5")).toHaveText(before ?? "");

  // The first tab closes: the second takes over, no reload.
  await page.close();
  await expect(conflict).toBeHidden({ timeout: 15_000 });
  await expect(second.getByTestId("spreadsheet-grid")).not.toHaveAttribute("aria-readonly", "true");
  await expect(cell(second, "B4")).toHaveText("888");
  await enterAt(second, "B5", "55");
  await waitForSaved(second);
});
