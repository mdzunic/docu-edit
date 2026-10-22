import { test, expect, type Page } from "@playwright/test";
import { MOD, LINE_START, LINE_END, extend } from "./keys";

// SPEC-020 — clear formatting and the format painter. Selection setup and
// model reads follow the same patterns as e2e/SPEC-012.spec.ts and
// e2e/SPEC-019.spec.ts.

async function createDocument(page: Page): Promise<string> {
  await page.goto("/");
  await page.getByTestId("new-document").click();
  await expect(page).toHaveURL(/\/doc\/.+/);
  const url = page.url();
  return url.slice(url.lastIndexOf("/") + 1);
}

const SAVED = "All changes saved";

async function waitForSaved(page: Page): Promise<void> {
  await expect(page.getByTestId("save-state")).toHaveText(SAVED);
}

const content = (page: Page) => page.getByTestId("document-content");
const pickerCurrent = (page: Page) => page.getByTestId("style-picker-current");
const clearFormattingButton = (page: Page) => page.getByTestId("clear-formatting-button");
const painterButton = (page: Page) => page.getByTestId("format-painter-button");

// SPEC-049 moved Clear formatting and the Format painter from the toolbar into
// the Format menu of the application menu bar. The menu title carries
// mousedown-prevent, so opening the menu preserves the editor's selection —
// arming the painter still captures the source paragraph exactly as before,
// and the painter item keeps the menu open so its double-click-to-lock
// gesture survives (data-menu-keep-open). The armed state now reports as
// aria-checked (menuitemcheckbox) rather than aria-pressed.
async function openFormatMenu(page: Page): Promise<void> {
  await page.getByTestId("menu-title-format").click();
}

async function armPainter(page: Page): Promise<void> {
  await openFormatMenu(page);
  await painterButton(page).click();
}

async function lockPainter(page: Page): Promise<void> {
  await openFormatMenu(page);
  await painterButton(page).dblclick();
}

interface ModelRun {
  text: string;
  formatting: Record<string, unknown>;
}

interface ModelBlock {
  runs: ModelRun[];
  formatting?: Record<string, unknown>;
  list?: { type: string; level: number };
  styleId?: string;
}

interface ModelShape {
  sections: Array<{ blocks: ModelBlock[] }>;
}

/** The newest persisted model (draft beats committed when newer) — same
 * IndexedDB read as e2e/SPEC-012.spec.ts. */
async function readLatestModel(page: Page, id: string): Promise<ModelShape> {
  return page.evaluate(async (docId) => {
    const db = await new Promise<IDBDatabase>((resolve, reject) => {
      const request = indexedDB.open("docu-edit"); // no explicit version: the app has already opened the DB at its current schema version by this point
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error);
    });
    const get = (key: string) =>
      new Promise<Record<string, never> | undefined>((resolve, reject) => {
        const tx = db.transaction("documents", "readonly");
        const rq = tx.objectStore("documents").get(key);
        rq.onsuccess = () => resolve(rq.result);
        rq.onerror = () => reject(rq.error);
      });
    const committed = (await get(docId)) as
      | { lastModified: number; model: unknown }
      | undefined;
    const draft = (await get("draft:" + docId)) as
      | { savedAt: number; model: unknown }
      | undefined;
    db.close();
    if (draft && (!committed || draft.savedAt > committed.lastModified)) {
      return draft.model;
    }
    return committed?.model ?? null;
  }, id) as Promise<ModelShape>;
}

function block(model: ModelShape, index: number): ModelBlock {
  return model.sections[0].blocks[index];
}

/** Selects a substring of the current line by character offsets, via
 * Home + arrow keys (the technique e2e/SPEC-012.spec.ts uses). */
async function selectRange(page: Page, from: number, to: number): Promise<void> {
  await page.keyboard.press(LINE_START);
  for (let i = 0; i < from; i++) {
    await page.keyboard.press("ArrowRight");
  }
  await page.keyboard.down("Shift");
  for (let i = from; i < to; i++) {
    await page.keyboard.press("ArrowRight");
  }
  await page.keyboard.up("Shift");
}

/** Selects an entire (single-line) paragraph, start to end. */
async function selectWholeLine(page: Page): Promise<void> {
  await page.keyboard.press(LINE_START);
  await page.keyboard.press(extend(LINE_END));
}

/** Polls the persisted model until `read` matches, for the assertions that
 * follow a format-painter application.
 *
 * `waitForSaved` is NOT a barrier there. The painter applies from a
 * `selectionchange`, which lands a task or two after the keystroke that made
 * the selection; until it does, no edit has happened yet, so `save-state`
 * still legitimately reads "All changes saved" from the PREVIOUS edit.
 * `waitForSaved` returns on that stale text and the one-shot `readLatestModel`
 * after it runs before the painter has touched anything. The window is real —
 * measured ~30ms on a warm dev server, reproducible on a clean checkout — and
 * a fast machine loses the race often enough to red four tests in this file.
 * Polling keeps the same expectation and drops the timing assumption. */
function polledModel<T>(page: Page, id: string, read: (model: ModelShape) => T) {
  return expect.poll(async () => read(await readLatestModel(page, id)));
}

test("the Format menu offers Clear formatting and Format painter (AC-1, AC-9)", async ({ page }) => {
  await createDocument(page);
  await openFormatMenu(page);
  await expect(clearFormattingButton(page)).toBeVisible();
  await expect(painterButton(page)).toBeVisible();
  await expect(painterButton(page)).toHaveAttribute("role", "menuitemcheckbox");
});

test("the Clear formatting button and Ctrl+\\ strip character formatting and reset the paragraph to Normal, without losing content or list membership (AC-1, AC-2, AC-3, AC-4, AC-6)", async ({
  page,
}) => {
  const id = await createDocument(page);
  await content(page).click();
  await page.keyboard.type("Styled list item");
  await page.getByTestId("list-bulleted").click();
  await selectWholeLine(page);
  await page.keyboard.press(`${MOD}+b`);
  await selectRange(page, 0, 6);
  await page.keyboard.press(`${MOD}+i`);
  await page.keyboard.press(`${MOD}+Alt+1`); // Heading 1 (SPEC-019)
  await expect(pickerCurrent(page)).toHaveText("Heading 1");

  await selectWholeLine(page);
  await page.keyboard.press(`${MOD}+\\`);
  await waitForSaved(page);

  await expect(pickerCurrent(page)).toHaveText("Normal");
  const model = await readLatestModel(page, id);
  const first = block(model, 0);
  expect(first.runs.map((run) => run.text).join("")).toBe("Styled list item");
  expect(first.runs.every((run) => Object.keys(run.formatting).length === 0)).toBe(true);
  expect(first.styleId).toBe("Normal");
  expect(first.list).toEqual({ type: "bullet", level: 1 });
});

test("the Clear formatting button performs the same action as Ctrl+\\ (button click)", async ({ page }) => {
  const id = await createDocument(page);
  await content(page).click();
  await page.keyboard.type("Bold text");
  await selectWholeLine(page);
  await page.keyboard.press(`${MOD}+b`);
  await selectWholeLine(page);
  await openFormatMenu(page);
  await clearFormattingButton(page).click();
  await waitForSaved(page);

  const model = await readLatestModel(page, id);
  expect(block(model, 0).runs[0].formatting).toEqual({});
});

test("Clear formatting is a single undo step (AC-18)", async ({ page }) => {
  const id = await createDocument(page);
  await content(page).click();
  await page.keyboard.type("Bold text");
  await selectWholeLine(page);
  await page.keyboard.press(`${MOD}+b`);
  await waitForSaved(page);
  await selectWholeLine(page);
  await page.keyboard.press(`${MOD}+\\`);
  await waitForSaved(page);
  expect(block(await readLatestModel(page, id), 0).runs[0].formatting).toEqual({});

  await page.keyboard.press(`${MOD}+z`);
  await waitForSaved(page);
  expect(block(await readLatestModel(page, id), 0).runs[0].formatting).toEqual({ bold: true });
});

test("a single click arms the painter (cursor + pressed state) and applies once to the next selection, then disarms (AC-9, AC-11, AC-12, AC-13, AC-19)", async ({
  page,
}) => {
  const id = await createDocument(page);
  await content(page).click();
  await page.keyboard.type("Source bold");
  await page.keyboard.press("Enter");
  await page.keyboard.type("Target plain");
  await page.keyboard.press(LINE_START);
  await page.keyboard.press("ArrowUp");
  await selectWholeLine(page);
  await page.keyboard.press(`${MOD}+b`);
  await waitForSaved(page);

  // Caret inside the bold source paragraph, then arm the painter.
  await page.keyboard.press(LINE_START);
  await expect(painterButton(page)).toHaveAttribute("aria-checked", "false");
  await armPainter(page);
  await expect(painterButton(page)).toHaveAttribute("aria-checked", "true");
  await expect(content(page)).toHaveCSS("cursor", "copy");

  // The next selection — the plain second paragraph — receives the copy.
  await page.keyboard.press("ArrowDown");
  await selectWholeLine(page);

  await polledModel(page, id, (model) => block(model, 1).runs[0].formatting).toEqual({ bold: true });
  await expect(painterButton(page)).toHaveAttribute("aria-checked", "false");
  await expect(content(page)).not.toHaveCSS("cursor", "copy");
});

test("double-clicking the painter keeps it armed for repeated application until Escape (AC-14, AC-15, AC-17)", async ({
  page,
}) => {
  const id = await createDocument(page);
  await content(page).click();
  await page.keyboard.type("Source bold");
  await page.keyboard.press("Enter");
  await page.keyboard.type("Target one");
  await page.keyboard.press("Enter");
  await page.keyboard.type("Target two");
  await page.keyboard.press(`${MOD}+Home`);
  await selectWholeLine(page);
  await page.keyboard.press(`${MOD}+b`);
  await waitForSaved(page);

  await page.keyboard.press(`${MOD}+Home`);
  await lockPainter(page);
  await expect(painterButton(page)).toHaveAttribute("aria-checked", "true");

  // Each application is awaited where it happens, not both at the end. In
  // repeat mode the painter stays armed, so neither waitForSaved nor the
  // aria-checked assertion is a barrier — without one the second selection can
  // be made before the FIRST application lands, and applyEdit's caret restore
  // then throws it away, leaving "Target two" plain.
  await page.keyboard.press("ArrowDown");
  await selectWholeLine(page);
  await polledModel(page, id, (model) => block(model, 1).runs[0].formatting).toEqual({ bold: true });
  await expect(painterButton(page)).toHaveAttribute("aria-checked", "true");

  await page.keyboard.press("ArrowDown");
  await page.keyboard.press(LINE_START);
  await selectWholeLine(page);
  await polledModel(page, id, (model) => block(model, 2).runs[0].formatting).toEqual({ bold: true });
  await expect(painterButton(page)).toHaveAttribute("aria-checked", "true");

  // Repeat mode left the first application in place while making the second.
  await polledModel(page, id, (model) => [
    block(model, 1).runs[0].formatting,
    block(model, 2).runs[0].formatting,
  ]).toEqual([{ bold: true }, { bold: true }]);

  await page.keyboard.press("Escape");
  await expect(painterButton(page)).toHaveAttribute("aria-checked", "false");
});

test("clicking the painter button again while armed for repeated use disarms it (AC-16)", async ({ page }) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("Text");
  await lockPainter(page);
  await expect(painterButton(page)).toHaveAttribute("aria-checked", "true");
  await painterButton(page).click();
  await expect(painterButton(page)).toHaveAttribute("aria-checked", "false");
});

test("copies paragraph formatting only when a whole paragraph is selected (AC-10)", async ({ page }) => {
  const id = await createDocument(page);
  await content(page).click();
  await page.keyboard.type("Centered source");
  await page.keyboard.press("Enter");
  await page.keyboard.type("Target line");
  await page.keyboard.press(`${MOD}+Home`);
  await page.getByTestId("align-center").click();
  await waitForSaved(page);

  // Whole-paragraph selection: paragraph formatting is copied too.
  await page.keyboard.press(`${MOD}+Home`);
  await selectWholeLine(page);
  await armPainter(page);
  await page.keyboard.press("ArrowDown");
  await selectWholeLine(page);

  await polledModel(page, id, (model) => block(model, 1).formatting).toEqual({ alignment: "center" });
});

test("format painter application is a single undo step (AC-19)", async ({ page }) => {
  const id = await createDocument(page);
  await content(page).click();
  await page.keyboard.type("Source bold");
  await page.keyboard.press("Enter");
  await page.keyboard.type("Target plain");
  await page.keyboard.press(`${MOD}+Home`);
  await selectWholeLine(page);
  await page.keyboard.press(`${MOD}+b`);
  await waitForSaved(page);

  await page.keyboard.press(`${MOD}+Home`);
  await armPainter(page);
  await page.keyboard.press("ArrowDown");
  await selectWholeLine(page);
  await polledModel(page, id, (model) => block(model, 1).runs[0].formatting).toEqual({ bold: true });

  await page.keyboard.press(`${MOD}+z`);
  await polledModel(page, id, (model) => block(model, 1).runs[0].formatting).toEqual({});
});
