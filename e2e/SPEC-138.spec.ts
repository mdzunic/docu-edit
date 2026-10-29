import { test, expect, type Download, type Locator, type Page } from "@playwright/test";
import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { equalBytes } from "../src/lib/core/equality";
import { unzip } from "../src/lib/core/zip";

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

/** A recorded sheetView without its sheet id (SPEC-138's SheetViewState). */
type PlantedView = { active: { row: number; col: number }; topRow: number; leftCol: number };

/** Writes a draft newer than the committed workbook straight into the app's
 * store: the committed shell read back, its manifest's keys listed in a
 * `rowKeys` manifest — committed chunk keys, except "Data"'s first chunk,
 * whose rows are copied under its draft key with B3 rewritten — and a view
 * that opens on "Data" (by default with C4 active, scrolled to the top). What
 * an unclean close leaves behind. */
async function plantNewerDraft(page: Page, id: string, value: string, view: PlantedView = { active: { row: 3, col: 2 }, topRow: 0, leftCol: 0 }): Promise<void> {
  await page.evaluate(
    async ([docId, drafted, draftView]) => {
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
        sheetView: { sheetId: dataSheet.id, ...draftView },
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
    [id, value, view] as const,
  );
}

/** The first row and column the grid shows past its headers, as the header
 * labels read ("47", "J") — what a reopened workbook must scroll back to. */
async function topLeftShown(page: Page): Promise<{ row: string; col: string }> {
  return page.getByTestId("spreadsheet-grid").evaluate((grid) => {
    const headers = (testId: string) =>
      [...grid.querySelectorAll<HTMLElement>(`[data-testid="${testId}"]`)].map((el) => ({ text: el.textContent?.trim() ?? "", box: el.getBoundingClientRect() }));
    const rows = headers("sheet-row-header");
    const cols = headers("sheet-col-header");
    // Just past the column letters and the row numbers.
    const y = Math.max(...cols.map((col) => col.box.bottom)) + 1;
    const x = Math.max(...rows.map((row) => row.box.right)) + 1;
    return {
      row: rows.find((row) => row.box.top <= y && y < row.box.bottom)?.text ?? "",
      col: cols.find((col) => col.box.left <= x && x < col.box.right)?.text ?? "",
    };
  });
}

/** Scrolls the grid by (dx, dy) px a screen at a time, as a user's scrolling
 * does — each step's scroll event grows the extent the next step needs. */
async function scrollGrid(page: Page, dx: number, dy: number): Promise<void> {
  const grid = page.getByTestId("spreadsheet-grid");
  const target = await grid.evaluate((el, [x, y]) => ({ left: el.scrollLeft + x, top: el.scrollTop + y }), [dx, dy] as const);
  await expect
    .poll(() =>
      grid.evaluate((el, to) => {
        el.scrollLeft = Math.min(to.left, el.scrollLeft + el.clientWidth / 2);
        el.scrollTop = Math.min(to.top, el.scrollTop + el.clientHeight / 2);
        return el.scrollLeft >= to.left && el.scrollTop >= to.top;
      }, target),
    )
    .toBe(true);
}

test("a reopened workbook scrolls back to its committed top-left row and column (AC-8)", async ({ page }) => {
  test.setTimeout(120_000); // a long session; 35–45 s alone, more beside other files at four workers
  await page.goto("/");
  await page.getByTestId("open-document-input").setInputFiles(FIXTURE);
  await expect(page).toHaveURL(/\/sheets\/.+/);
  const id = page.url().match(/\/sheets\/([^/?#]+)/)?.[1] ?? "";
  await showData(page);
  await expect(cell(page, "B3")).toHaveText("3");

  // Scroll well away from A1 on both axes and edit a cell the scroll shows;
  // then scroll on, leaving the active cell out of view.
  await scrollGrid(page, 1100, 900);
  await expect.poll(() => topLeftShown(page)).not.toEqual({ row: "1", col: "A" });
  const edited = await topLeftShown(page);
  const column = String.fromCharCode(edited.col.charCodeAt(0) + 1);
  await enterAt(page, `${column}${Number(edited.row) + 2}`, "5150");
  await waitForSaved(page);
  await scrollGrid(page, 200, 1200);
  await expect.poll(async () => (await topLeftShown(page)).row).not.toBe(edited.row);
  const shown = await topLeftShown(page);
  expect(Number(shown.row)).toBeGreaterThan(Number(edited.row) + 30);
  expect(shown.col > "J").toBe(true); // past the columns a reopened grid's first extent reaches

  // Close (the clean close commits, awaited) and reopen, then reload: the same
  // top-left row and column each time, not the active cell scrolled into view.
  await page.getByTestId("close-document").click();
  await expect(page.getByTestId("library")).toBeVisible();
  for (let reopen = 0; reopen < 2; reopen += 1) {
    if (reopen === 0) {
      await page.goto(`/sheets/${id}`);
    } else {
      await page.reload();
    }
    await expect(tab(page, "Data")).toHaveAttribute("aria-selected", "true");
    await expect(page.getByTestId("recovery-bar")).toBeHidden();
    await waitForSaved(page);
    await expect.poll(() => topLeftShown(page)).toEqual(shown);
    await page.waitForTimeout(500); // nothing scrolls it away once settled
    expect(await topLeftShown(page)).toEqual(shown);
    await expect(page.getByTestId("name-box")).toHaveValue(`${column}${Number(edited.row) + 3}`);
  }
});

test("a recovered draft, Discard, and a committed view the model's own view fields disagree with each restore their sheet, active cell and top-left (AC-8, AC-27, AC-29)", async ({ page }) => {
  test.setTimeout(120_000);
  const id = await openFixture(page, "mini-book.xlsx");
  // The committed view: "Summary", D5 active, at the top left.
  await tab(page, "Summary").click();
  await expect(tab(page, "Summary")).toHaveAttribute("aria-selected", "true");
  await clickCell(page, "D5");
  await page.getByTestId("close-document").click();
  await expect(page.getByTestId("library")).toBeVisible();

  // A crash's draft scrolled on "Data": K54 active, J47 the top-left cell.
  await plantNewerDraft(page, id, "5151", { active: { row: 53, col: 10 }, topRow: 46, leftCol: 9 });
  await page.goto(`/sheets/${id}`);
  await expect(page.getByTestId("recovery-bar")).toBeVisible();
  await expect(tab(page, "Data")).toHaveAttribute("aria-selected", "true");
  await expect(page.getByTestId("name-box")).toHaveValue("K54");
  await expect.poll(() => topLeftShown(page)).toEqual({ row: "47", col: "J" });

  // Discard: the committed view comes back with the committed value.
  await page.getByTestId("recovery-discard").click();
  await expect(page.getByTestId("recovery-bar")).toBeHidden();
  await expect(tab(page, "Summary")).toHaveAttribute("aria-selected", "true");
  await expect(page.getByTestId("name-box")).toHaveValue("D5");
  await expect.poll(() => topLeftShown(page)).toEqual({ row: "1", col: "A" });
  await showData(page);
  await expect(cell(page, "B3")).toHaveText("3");

  // A commit whose view the model's own view fields disagree with (as when
  // view-state writes were held back behind edits): the recorded view wins on
  // reopen — its sheet, its active cell, its top-left row and column.
  await page.getByTestId("close-document").click();
  await expect(page.getByTestId("library")).toBeVisible();
  await page.evaluate(async (docId) => {
    const db = await new Promise<IDBDatabase>((resolve, reject) => {
      const request = indexedDB.open("docu-edit");
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error);
    });
    const store = () => db.transaction("documents", "readwrite").objectStore("documents");
    const record = await new Promise<{ model: { sheets: { id: string; name: string }[] }; sheetView?: unknown }>((resolve, reject) => {
      const request = store().get(docId);
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error);
    });
    const data = record.model.sheets.find((sheet) => sheet.name === "Data")!;
    record.sheetView = { sheetId: data.id, active: { row: 3, col: 2 }, topRow: 46, leftCol: 9 };
    await new Promise<void>((resolve, reject) => {
      const request = store().put(record);
      request.onsuccess = () => resolve();
      request.onerror = () => reject(request.error);
    });
    db.close();
  }, id);
  await page.goto(`/sheets/${id}`);
  await expect(tab(page, "Data")).toHaveAttribute("aria-selected", "true");
  await expect(page.getByTestId("recovery-bar")).toBeHidden();
  await waitForSaved(page);
  await expect(page.getByTestId("name-box")).toHaveValue("C4");
  await expect.poll(() => topLeftShown(page)).toEqual({ row: "47", col: "J" });
});

test("a workbook autosaves, recovers a crash's draft, snapshots, names, previews and restores its versions, and locks a second tab", async ({ page }) => {
  test.setTimeout(120_000); // a long session; 35–45 s alone, more beside other files at four workers
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
  // Undo has nothing to undo after a restore, and the panel says so (AC-55).
  await expect(panel).toContainText("Restoring replaces the workbook and clears undo — the version saved just before is at the top of the list");
  // (Twice: a stack that survived would take B4 back past both edits to the file's value.)
  await clickCell(page, "B4");
  await page.keyboard.press("ControlOrMeta+z");
  await page.keyboard.press("ControlOrMeta+z");
  await expect(cell(page, "B4")).toHaveText("888");
  await expect(page.getByTestId("sheet-menu-undo")).toBeDisabled();

  // A second tab on the same workbook is read-only.
  const second = await page.context().newPage();
  await second.goto(`/sheets/${id}`);
  const conflict = second.getByTestId("sheet-conflict-bar");
  await expect(conflict).toHaveText("This workbook is open in another tab");
  await expect(second.getByTestId("spreadsheet-grid")).toHaveAttribute("aria-readonly", "true");
  await expect(second.getByTestId("save-state")).toHaveCount(0);
  // Every edit command is disabled: the format strip and its merge split
  // button, Add sheet, the menus.
  await expect(second.getByTestId("sheet-bold")).toBeDisabled();
  await expect(second.getByTestId("sheet-merge")).toBeDisabled();
  await expect(second.getByTestId("sheet-merge-menu-button")).toBeDisabled();
  await expect(second.getByTestId("sheet-add")).toBeDisabled();
  await second.getByRole("menubar").getByText("Format", { exact: true }).click();
  await expect(second.getByTestId("sheet-menu-format-cells")).toBeDisabled();
  await second.keyboard.press("Escape");
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
  await expect(second.getByTestId("sheet-bold")).toBeEnabled();
  await expect(second.getByTestId("sheet-merge")).toBeEnabled();
  await expect(cell(second, "B4")).toHaveText("888");
  await enterAt(second, "B5", "55");
  await waitForSaved(second);
});

// --- QA (SPEC-138): the criteria the session above passes through quickly,
// each asserted on its own workbook.

/** Every key the documents store holds for a document (committed, draft,
 * chunk and source records), with the id written as "<doc>". */
async function documentKeys(page: Page, id: string): Promise<string[]> {
  return page.evaluate(async (docId) => {
    const db = await new Promise<IDBDatabase>((resolve, reject) => {
      const request = indexedDB.open("docu-edit");
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error);
    });
    const keys = await new Promise<IDBValidKey[]>((resolve, reject) => {
      const request = db.transaction("documents", "readonly").objectStore("documents").getAllKeys();
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error);
    });
    db.close();
    return keys.map(String).filter((key) => key.includes(docId)).map((key) => key.replace(docId, "<doc>"));
  }, id);
}

/** One record of the app's IndexedDB store. */
async function storedRecord(page: Page, store: string, key: string): Promise<unknown> {
  return page.evaluate(
    async ([storeName, recordKey]) => {
      const db = await new Promise<IDBDatabase>((resolve, reject) => {
        const request = indexedDB.open("docu-edit");
        request.onsuccess = () => resolve(request.result);
        request.onerror = () => reject(request.error);
      });
      const record = await new Promise<unknown>((resolve, reject) => {
        const request = db.transaction(storeName, "readonly").objectStore(storeName).get(recordKey);
        request.onsuccess = () => resolve(request.result);
        request.onerror = () => reject(request.error);
      });
      db.close();
      return record;
    },
    [store, key] as const,
  );
}

/** Every key of the versions store that belongs to a document. */
async function versionKeys(page: Page, id: string): Promise<string[]> {
  return page.evaluate(async (docId) => {
    const db = await new Promise<IDBDatabase>((resolve, reject) => {
      const request = indexedDB.open("docu-edit");
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error);
    });
    const keys = await new Promise<IDBValidKey[]>((resolve, reject) => {
      const request = db.transaction("versions", "readonly").objectStore("versions").getAllKeys();
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error);
    });
    db.close();
    return keys.map(String).filter((key) => key.includes(docId));
  }, id);
}

async function openFixture(page: Page, file: string): Promise<string> {
  await page.goto("/");
  await page.getByTestId("open-document-input").setInputFiles(join(dirname(FIXTURE), file));
  await expect(page).toHaveURL(/\/sheets\/.+/);
  await waitForSaved(page);
  return page.url().match(/\/sheets\/([^/?#]+)/)?.[1] ?? "";
}

async function downloadVia(page: Page, trigger: Locator): Promise<Download> {
  const download = page.waitForEvent("download");
  await trigger.click();
  return download;
}

/** Asserts two packages hold the same parts, in the same order, byte for byte. */
async function expectSameParts(actual: Download, expected: Uint8Array): Promise<void> {
  const got = await unzip(new Uint8Array(readFileSync((await actual.path())!)));
  const want = await unzip(expected);
  expect([...got.keys()]).toEqual([...want.keys()]);
  for (const [name, bytes] of want) {
    expect(equalBytes(got.get(name)!, bytes), name).toBe(true);
  }
}

async function openHistory(page: Page): Promise<Locator> {
  const toggle = page.getByTestId("sheet-history-toggle");
  if ((await toggle.getAttribute("aria-pressed")) !== "true") {
    await toggle.click();
  }
  const panel = page.getByTestId("sheet-history-panel");
  await expect(panel).toBeVisible();
  return panel;
}

test("drafts go on the 2 s cadence as bulk-free shells over draft chunk records, and a commit clears them (AC-9, AC-11, AC-12, AC-15, AC-17, AC-18, AC-23, AC-36, AC-40)", async ({ page }) => {
  const id = await openFixture(page, "mini-book.xlsx");
  const badge = page.getByTestId("save-state");
  // Record every text the badge shows from here on.
  await badge.evaluate((el) => {
    const seen: string[] = [];
    (window as unknown as { qaBadge: string[] }).qaBadge = seen;
    new MutationObserver(() => seen.push(el.textContent?.trim() ?? "")).observe(el, { subtree: true, childList: true, characterData: true });
  });
  const shown = () => page.evaluate(() => (window as unknown as { qaBadge: string[] }).qaBadge.slice());

  // Switching sheets and moving the selection is not an edit.
  await showData(page);
  await clickCell(page, "C7");
  await page.keyboard.press("ArrowDown");
  await tab(page, "Summary").click();
  await clickCell(page, "B2");
  await page.waitForTimeout(2500);
  expect(await shown()).toEqual([]);
  expect((await documentKeys(page, id)).filter((key) => key.startsWith("draft"))).toEqual([]);

  // An edit goes on the draft cadence: Saving…, then All changes saved.
  await showData(page);
  await enterAt(page, "B8", "8080");
  await waitForSaved(page);
  expect(await shown()).toEqual(expect.arrayContaining(["Saving…", SAVED]));

  const keys = await documentKeys(page, id);
  const draftChunks = keys.filter((key) => key.startsWith("draftsheetrows:<doc>:"));
  expect(draftChunks).toHaveLength(1);
  expect(keys).toContain(draftChunks[0]!.slice("draft".length)); // the committed chunk it shadows
  expect(keys.filter((key) => key.startsWith("draftsheetsource:"))).toEqual([]);

  type Shell = {
    model: { opaque?: unknown[]; sheets: { id: string; name: string; rows: unknown[]; sourceXml?: string }[] };
    rowKeys: Record<string, string[]>;
    sourceKey: string;
    sheetView: { sheetId: string; active: { row: number; col: number }; topRow: number; leftCol: number };
  };
  const draft = (await storedRecord(page, "documents", `draft:${id}`)) as Shell;
  for (const sheet of draft.model.sheets) {
    expect(sheet.rows, sheet.name).toEqual([]);
    expect(sheet.sourceXml, sheet.name).toBeUndefined();
  }
  expect(draft.model.opaque ?? []).toEqual([]);
  expect(draft.sourceKey).toBe(`sheetsource:${id}`);
  const data = draft.model.sheets.find((sheet) => sheet.name === "Data")!;
  expect(draft.rowKeys[data.id]![0]).toBe(`draftsheetrows:${id}:${data.id}:0000`);
  for (const sheet of draft.model.sheets.filter((s) => s.id !== data.id)) {
    for (const key of draft.rowKeys[sheet.id] ?? []) {
      expect(key.startsWith("sheetrows:"), key).toBe(true);
    }
  }
  expect(draft.sheetView).toEqual({ sheetId: data.id, active: { row: 8, col: 1 }, topRow: 0, leftCol: 0 });

  // Download is a commit point: the drafts go, and a version lands in the
  // versions store with its sources as content-addressed blobs.
  await downloadVia(page, page.getByTestId("spreadsheet-download"));
  await expect.poll(async () => (await documentKeys(page, id)).filter((key) => key.startsWith("draft"))).toEqual([]);
  await expect.poll(async () => (await versionKeys(page, id)).filter((key) => key.startsWith(`version:${id}:`)).length).toBe(1);
  const blobs = (await versionKeys(page, id)).filter((key) => key.startsWith("blob:"));
  expect(blobs.length).toBeGreaterThan(0);
  for (const key of blobs) {
    expect(key).toMatch(new RegExp(`^blob:${id}:[0-9a-f]{64}$`));
  }
});

test("the history pane lists, names, filters, deletes, previews and restores versions, and a restore keeps every part of the file (AC-50..AC-54, AC-56, AC-57)", async ({ page }) => {
  test.setTimeout(120_000);
  const source = new Uint8Array(readFileSync(FIXTURE));
  await openFixture(page, "mini-book.xlsx");
  const panel = await openHistory(page);
  const versions = panel.getByTestId("history-version");

  // An edit undone, then a download: the snapshot is the file as opened.
  await showData(page);
  await enterAt(page, "B3", "1");
  await page.keyboard.press("ControlOrMeta+z");
  await expect(cell(page, "B3")).toHaveText("3");
  const snapshot = await downloadVia(page, page.getByTestId("spreadsheet-download"));
  await expectSameParts(snapshot, source);
  const snapshotBytes = new Uint8Array(readFileSync((await snapshot.path())!));
  await expect(versions).toHaveCount(1);
  await expect(versions.first()).toContainText("3 sheets · 1,535 cells");
  await expect(versions.first()).toContainText("You");

  // A filled cell, then the same cell cleared: newest first, with the deltas.
  await enterAt(page, "E2", "x1");
  await downloadVia(page, page.getByTestId("spreadsheet-download"));
  await expect(versions).toHaveCount(2);
  await expect(versions.first()).toContainText("1,536 cells · +1 cell");
  await clickCell(page, "E2");
  await page.keyboard.press("Delete");
  await expect(cell(page, "E2")).toHaveText("");
  await downloadVia(page, page.getByTestId("spreadsheet-download"));
  await expect(versions).toHaveCount(3);
  await expect(versions.first()).toContainText("1,535 cells · −1 cell");

  // Name, filter to named versions, remove the name; delete a version.
  const plusOne = versions.nth(1);
  await plusOne.getByTestId("history-name-input").fill("QA named");
  await plusOne.getByTestId("history-name-submit").click();
  await expect(plusOne.getByTestId("history-name")).toHaveText("QA named");
  await panel.getByTestId("history-named-only").check();
  await expect(versions).toHaveCount(1);
  await panel.getByTestId("history-named-only").uncheck();
  await expect(versions).toHaveCount(3);
  await plusOne.getByRole("button", { name: "Remove name" }).click();
  await expect(panel.getByTestId("history-name").filter({ hasText: "QA named" })).toHaveCount(0);
  await versions.first().getByTestId("history-delete").click();
  await expect(versions).toHaveCount(2);

  // The live workbook: a sheet added, an edit on Summary, C3 selected.
  await page.getByTestId("sheet-add").click();
  await expect(page.getByTestId("sheet-tab")).toHaveCount(3);
  await tab(page, "Summary").click();
  await enterAt(page, "C12", "u1");
  await clickCell(page, "C3");

  // Preview the snapshot: its own tabs, on its own active sheet, read-only.
  const original = versions.last();
  await original.getByTestId("history-preview").click();
  await expect(panel.getByTestId("history-preview-banner")).toContainText("Previewing an earlier version, read-only");
  await expect(panel.getByTestId("history-exit-preview")).toHaveText("Back to current");
  await expect(page.getByTestId("spreadsheet-grid")).toHaveAttribute("aria-readonly", "true");
  await expect(page.getByTestId("sheet-tab")).toHaveText([/Summary/, /Data/]);
  await expect(tab(page, "Data")).toHaveAttribute("aria-selected", "true");
  await expect(cell(page, "E2")).toHaveText("");
  await expect(page.getByTestId("sheet-merge")).toBeDisabled();

  // Back to current: the live tabs, selection and undo stack are untouched.
  await panel.getByTestId("history-exit-preview").click();
  await expect(panel.getByTestId("history-preview-banner")).toBeHidden();
  await expect(page.getByTestId("spreadsheet-grid")).not.toHaveAttribute("aria-readonly", "true");
  await expect(page.getByTestId("sheet-tab")).toHaveCount(3);
  await expect(tab(page, "Summary")).toHaveAttribute("aria-selected", "true");
  await expect(page.getByTestId("name-box")).toHaveValue("C3");
  await clickCell(page, "C3");
  await page.keyboard.press("ControlOrMeta+z");
  await expect(cell(page, "C12")).toHaveText("");
  await page.keyboard.press("ControlOrMeta+Shift+z");
  await expect(cell(page, "C12")).toHaveText("u1");

  // Edit two sheets, then restore the snapshot, and undo that restore.
  await enterAt(page, "H2", "qa-h2");
  await showData(page);
  await enterAt(page, "B3", "31337");
  await original.getByTestId("history-restore").click();
  await expect(page.getByTestId("toast")).toContainText("Version restored");
  await expect(page.getByTestId("sheet-tab")).toHaveCount(2);
  await expect(versions).toHaveCount(4);
  await page.getByTestId("toast-action").click(); // Undo restore
  await expect(page.getByTestId("sheet-tab")).toHaveCount(3);
  await showData(page);
  await expect(cell(page, "B3")).toHaveText("31337");
  await expect(versions).toHaveCount(6);

  // Restore it again: the download is the snapshot's, part for part — and so
  // is the version's own download.
  await original.getByTestId("history-restore").click();
  await expect(page.getByTestId("sheet-tab")).toHaveText([/Summary/, /Data/]);
  await showData(page);
  await expect(cell(page, "B3")).toHaveText("3");
  await waitForSaved(page);
  await expectSameParts(await downloadVia(page, page.getByTestId("spreadsheet-download")), snapshotBytes);
  await expect(original.getByTestId("history-download")).toHaveText("Download .xlsx");
  const versionDownload = await downloadVia(page, original.getByTestId("history-download"));
  expect(versionDownload.suggestedFilename()).toBe("mini-book.xlsx");
  await expectSameParts(versionDownload, snapshotBytes);
});

test("an .ods workbook restored from a snapshot downloads part for part as its source (AC-57)", async ({ page }) => {
  const source = new Uint8Array(readFileSync(join(dirname(FIXTURE), "mini-book.ods")));
  await openFixture(page, "mini-book.ods");
  const panel = await openHistory(page);
  await enterAt(page, "B3", "1");
  await page.keyboard.press("ControlOrMeta+z");
  await expectSameParts(await downloadVia(page, page.getByTestId("spreadsheet-download")), source);
  await expect(panel.getByTestId("history-version")).toHaveCount(1);

  await enterAt(page, "B3", "31337");
  await tab(page, "Totals").click();
  await enterAt(page, "C2", "qa-c2");
  await page.getByTestId("sheet-add").click();
  await expect(page.getByTestId("sheet-tab")).toHaveCount(3);
  await panel.getByTestId("history-version").last().getByTestId("history-restore").click();
  await expect(page.getByTestId("sheet-tab")).toHaveText([/Inventory/, /Totals/]);
  await waitForSaved(page);
  const restored = await downloadVia(page, page.getByTestId("spreadsheet-download"));
  expect(restored.suggestedFilename()).toBe("mini-book.ods");
  await expectSameParts(restored, source);
  await expect(panel.getByTestId("history-version").last().getByTestId("history-download")).toHaveText("Download .ods");
});

test("a locked tab keeps its writing history actions and the merge split button disabled until the lock grants (AC-33, AC-35)", async ({ page }) => {
  const id = await openFixture(page, "mini-book.xlsx");
  await showData(page);
  await enterAt(page, "B4", "444");
  await downloadVia(page, page.getByTestId("spreadsheet-download"));

  const second = await page.context().newPage();
  await second.goto(`/sheets/${id}`);
  await expect(second.getByTestId("sheet-conflict-bar")).toHaveText("This workbook is open in another tab");
  const panel = await openHistory(second);
  const version = panel.getByTestId("history-version").first();
  await expect(version.getByTestId("history-restore")).toBeDisabled();
  await expect(version.getByTestId("history-name-submit")).toBeDisabled();
  await expect(version.getByTestId("history-delete")).toBeDisabled();
  await expect(version.getByTestId("history-preview")).toBeEnabled();
  await expect(version.getByTestId("history-download")).toBeEnabled();
  await expect(second.getByTestId("sheet-merge")).toBeDisabled();
  await expect(second.getByTestId("sheet-merge-menu-button")).toBeDisabled();
  // Not even a forced click opens the merge warning or the merge menu.
  await second.getByTestId("name-box").fill("A10:B11");
  await second.getByTestId("name-box").press("Enter");
  await second.getByTestId("sheet-merge").click({ force: true });
  await second.getByTestId("sheet-merge-menu-button").click({ force: true });
  await expect(second.getByTestId("sheet-merge-confirm")).toHaveCount(0);
  await expect(second.getByTestId("sheet-merge-centre")).toHaveCount(0);

  // Close (not a tab close) in the first tab hands the lock over.
  await page.getByTestId("close-document").click();
  await expect(page.getByTestId("library")).toBeVisible();
  await expect(second.getByTestId("sheet-conflict-bar")).toBeHidden({ timeout: 15_000 });
  await expect(second.getByTestId("spreadsheet-grid")).not.toHaveAttribute("aria-readonly", "true");
  await expect(version.getByTestId("history-restore")).toBeEnabled();
  await expect(second.getByTestId("sheet-merge")).toBeEnabled();
  await second.getByTestId("sheet-merge-menu-button").click();
  for (const item of ["sheet-merge-centre", "sheet-merge-across", "sheet-merge-cells", "sheet-unmerge"]) {
    await expect(second.getByTestId(item)).toBeEnabled();
  }
  await second.keyboard.press("Escape");
});

test("a failed draft write shows the error badge with one download link that runs the workbook download (AC-30, AC-31)", async ({ page }) => {
  await openFixture(page, "mini-book.xlsx");
  const badge = page.getByTestId("save-state");
  await expect(badge).toHaveClass(/sheet-save-state/);
  await expect(badge).toHaveAttribute("data-save-state", "saved");
  // Storage fills up: every documents-store put throws.
  await page.evaluate(() => {
    const put = IDBObjectStore.prototype.put;
    (window as unknown as { qaPut: typeof put }).qaPut = put;
    IDBObjectStore.prototype.put = function (this: IDBObjectStore, ...args: Parameters<typeof put>) {
      if (this.name === "documents") {
        throw new DOMException("Quota exceeded", "QuotaExceededError");
      }
      return put.apply(this, args);
    };
  });
  await showData(page);
  await enterAt(page, "B6", "6060");
  await expect(badge).toHaveText("Not saved — browser storage is full", { timeout: 15_000 });
  await expect(badge).toHaveAttribute("data-save-state", "error");
  await expect(page.getByTestId("sheet-save-state-download")).toHaveCount(1);
  const download = await downloadVia(page, page.getByTestId("sheet-save-state-download"));
  expect(download.suggestedFilename()).toBe("mini-book.xlsx");

  // Storage frees up: the next write clears the error.
  await page.evaluate(() => {
    IDBObjectStore.prototype.put = (window as unknown as { qaPut: typeof IDBObjectStore.prototype.put }).qaPut;
  });
  await enterAt(page, "B7", "7070");
  await waitForSaved(page);
  await expect(page.getByTestId("sheet-save-state-download")).toHaveCount(0);
});

test("a workbook over WORKBOOK_HISTORY_MAX_CELLS keeps no history and still commits (AC-44, AC-46, AC-47)", async ({ page }) => {
  test.setTimeout(120_000);
  // 26,000 rows × 10 columns = 260,000 cells, over the 250,000 limit.
  const lines: string[] = [];
  for (let row = 0; row < 26_000; row += 1) {
    lines.push(Array.from({ length: 10 }, (_, col) => String(row * 10 + col)).join(","));
  }
  await page.goto("/");
  await page.getByTestId("open-document-input").setInputFiles({ name: "big.csv", mimeType: "text/csv", buffer: Buffer.from(lines.join("\n")) });
  await expect(page).toHaveURL(/\/sheets\/.+/, { timeout: 60_000 });
  const id = page.url().match(/\/sheets\/([^/?#]+)/)?.[1] ?? "";
  await waitForSaved(page);
  await enterAt(page, "A1", "99");
  await waitForSaved(page);
  await downloadVia(page, page.getByTestId("spreadsheet-download"));
  const panel = await openHistory(page);
  await expect(panel.getByTestId("history-paused-notice")).toHaveText("Version history off — this workbook is larger than the 50 MB history budget");
  await expect(panel.getByTestId("history-version")).toHaveCount(0);
  // The refusal never blocks the commit: the drafts are gone.
  await expect.poll(async () => (await documentKeys(page, id)).filter((key) => key.startsWith("draft"))).toEqual([]);
  await expect(page.getByTestId("save-state")).toHaveText(SAVED);
});
