import { test, expect, type Page } from "@playwright/test";
import { MOD, LINE_START } from "./keys";

// SPEC-012 — character formatting: bold, italic, underline, strikethrough.
// Builds on SPEC-011's editing engine, so selection/caret setup here follows
// the same patterns as e2e/SPEC-011.spec.ts.

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
const formatButton = (page: Page, name: string) => page.getByTestId(`format-${name}`);

interface ModelRun {
  text: string;
  formatting: Record<string, unknown>;
}

interface ModelShape {
  sections: Array<{ blocks: Array<{ runs: ModelRun[] }> }>;
}

/** The newest persisted model (draft beats committed when newer) — same
 * IndexedDB read as e2e/SPEC-011.spec.ts. */
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

function firstParagraphRuns(model: ModelShape): ModelRun[] {
  return model.sections[0].blocks[0].runs;
}

/** Selects a substring of the current (single-paragraph) line by character
 * offsets, via Home + arrow keys — the same technique SPEC-011's e2e specs
 * use for precise selections. */
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

test("the toolbar offers Bold, Italic, Underline and Strikethrough with name+shortcut tooltips (AC-1, AC-2)", async ({
  page,
}) => {
  await createDocument(page);
  await expect(formatButton(page, "bold")).toBeVisible();
  await expect(formatButton(page, "bold")).toHaveAttribute("title", /Bold/);
  await expect(formatButton(page, "italic")).toHaveAttribute("title", /Italic/);
  await expect(formatButton(page, "underline")).toHaveAttribute("title", /Underline/);
  await expect(formatButton(page, "strikethrough")).toHaveAttribute("title", /Strikethrough/);
});

test("Ctrl+B formats exactly the selection, splitting runs at the boundaries (AC-3, AC-7, AC-8)", async ({
  page,
}) => {
  const id = await createDocument(page);
  await content(page).click();
  await page.keyboard.type("Hello world");
  await selectRange(page, 6, 11);
  await page.keyboard.press(`${MOD}+b`);
  await waitForSaved(page);

  const runs = firstParagraphRuns(await readLatestModel(page, id));
  expect(runs).toEqual([
    { text: "Hello ", formatting: {} },
    { text: "world", formatting: { bold: true } },
  ]);
});

test("Ctrl+I, Ctrl+U and Ctrl+Shift+X apply italic, underline and strikethrough (AC-4, AC-5, AC-6)", async ({
  page,
}) => {
  const id = await createDocument(page);
  await content(page).click();
  await page.keyboard.type("Styled");
  await selectRange(page, 0, 6);
  await page.keyboard.press(`${MOD}+i`);
  await selectRange(page, 0, 6);
  await page.keyboard.press(`${MOD}+u`);
  await selectRange(page, 0, 6);
  await page.keyboard.press(`${MOD}+Shift+x`);
  await waitForSaved(page);

  const runs = firstParagraphRuns(await readLatestModel(page, id));
  expect(runs).toEqual([
    { text: "Styled", formatting: { italic: true, underline: true, strikethrough: true } },
  ]);
});

test("a mixed selection shows indeterminate; the same selection's first click bolds it all, the second clears it (AC-14, AC-15, AC-16, AC-17)", async ({
  page,
}) => {
  const id = await createDocument(page);
  await content(page).click();
  await page.keyboard.type("bold and plain");
  await selectRange(page, 0, 4);
  await page.keyboard.press(`${MOD}+b`);

  const bold = formatButton(page, "bold");
  await selectRange(page, 0, "bold and plain".length);
  await expect(bold).toHaveAttribute("aria-pressed", "mixed");

  // The click must not steal focus/collapse the selection (AC-13, AC-15).
  await bold.click();
  await expect(bold).toHaveAttribute("aria-pressed", "true");
  await waitForSaved(page);
  expect(firstParagraphRuns(await readLatestModel(page, id))).toEqual([
    { text: "bold and plain", formatting: { bold: true } },
  ]);

  // Same selection, still live: a second click clears it everywhere (AC-17).
  await bold.click();
  await expect(bold).toHaveAttribute("aria-pressed", "false");
  await waitForSaved(page);
  expect(firstParagraphRuns(await readLatestModel(page, id))).toEqual([
    { text: "bold and plain", formatting: {} },
  ]);
});

test("a collapsed caret sets a pending format the next typed characters carry, discarded by moving the caret (AC-9, AC-10, AC-11, AC-12)", async ({
  page,
}) => {
  const id = await createDocument(page);
  await content(page).click();
  await page.keyboard.type("hello world");
  await page.keyboard.press(LINE_START);
  for (let i = 0; i < 6; i++) {
    await page.keyboard.press("ArrowRight");
  }

  const bold = formatButton(page, "bold");
  await expect(bold).toHaveAttribute("aria-pressed", "false");
  await page.keyboard.press(`${MOD}+b`);
  await expect(bold).toHaveAttribute("aria-pressed", "true");
  await page.keyboard.type("AB");
  await expect(content(page)).toHaveText("hello ABworld");
  await waitForSaved(page);
  expect(firstParagraphRuns(await readLatestModel(page, id))).toEqual([
    { text: "hello ", formatting: {} },
    { text: "AB", formatting: { bold: true } },
    { text: "world", formatting: {} },
  ]);

  // Moving the caret discards the pending format (AC-11).
  await page.keyboard.press(LINE_START);
  await expect(bold).toHaveAttribute("aria-pressed", "false");
  await page.keyboard.type("Z");
  await waitForSaved(page);
  const runs = firstParagraphRuns(await readLatestModel(page, id));
  expect(runs[0]).toEqual({ text: "Zhello ", formatting: {} });
});

test("bold round-trips through the .docx and .odt writers (apply, save, re-open preserves it) (AC-18, AC-19)", async ({
  page,
}) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("Hello world");
  await selectRange(page, 6, 11);
  await page.keyboard.press(`${MOD}+b`);
  await waitForSaved(page);

  await page.getByTestId("menu-title-file").click();

  await page.getByTestId("file-menu-export").click();
  const docxDownload = await Promise.all([
    page.waitForEvent("download"),
    page.getByTestId("download-docx").click(),
  ]).then(([download]) => download);
  const docxPath = await docxDownload.path();
  expect(docxPath).toBeTruthy();

  await page.getByTestId("menu-title-file").click();

  await page.getByTestId("file-menu-export").click();
  const odtDownload = await Promise.all([
    page.waitForEvent("download"),
    page.getByTestId("download-odt").click(),
  ]).then(([download]) => download);
  const odtPath = await odtDownload.path();
  expect(odtPath).toBeTruthy();

  for (const path of [docxPath, odtPath]) {
    await page.goto("/");
    await page.getByTestId("open-document-input").setInputFiles(path!);
    await expect(page).toHaveURL(/\/doc\/.+/);
    const spans = content(page).locator("span");
    await expect(spans.first()).toHaveText("Hello ");
    await expect(spans.nth(1)).toHaveText("world");
    await expect(spans.nth(1)).toHaveCSS("font-weight", "700");
    const plainWeight = await spans.first().evaluate((el) => getComputedStyle(el).fontWeight);
    expect(plainWeight).not.toBe("700");
  }
});
