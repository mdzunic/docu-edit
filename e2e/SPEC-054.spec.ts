import { test, expect, type Page } from "@playwright/test";

// SPEC-054 — tab stops and the tab ruler: the Tab-key precedence table, the
// literal-tab routes, tab advances rendering as sized boxes, the bar rule and
// the dialog's new controls. Model reads follow e2e/SPEC-014's IndexedDB read.

async function createDocument(page: Page): Promise<string> {
  await page.goto("/");
  await page.getByTestId("new-document").click();
  await expect(page).toHaveURL(/\/doc\/.+/);
  const url = page.url();
  return url.slice(url.lastIndexOf("/") + 1);
}

async function waitForSaved(page: Page): Promise<void> {
  await expect(page.getByTestId("save-state")).toHaveText("All changes saved");
}

const content = (page: Page) => page.getByTestId("document-content");

interface ModelBlock {
  runs: Array<{ text: string }>;
  list?: { level: number };
}
interface ModelShape {
  sections: Array<{ blocks: ModelBlock[] }>;
}

async function readModel(page: Page, id: string): Promise<ModelShape> {
  await expect
    .poll(async () =>
      page.evaluate(async (docId) => {
        const db = await new Promise<IDBDatabase>((resolve, reject) => {
          const request = indexedDB.open("docu-edit");
          request.onsuccess = () => resolve(request.result);
          request.onerror = () => reject(request.error);
        });
        const get = (key: string) =>
          new Promise<{ lastModified?: number; savedAt?: number; model?: unknown } | undefined>((resolve) => {
            const rq = db.transaction("documents", "readonly").objectStore("documents").get(key);
            rq.onsuccess = () => resolve(rq.result);
            rq.onerror = () => resolve(undefined);
          });
        const committed = await get(docId);
        const draft = await get("draft:" + docId);
        db.close();
        const model =
          draft && (!committed || (draft.savedAt ?? 0) > (committed.lastModified ?? 0))
            ? draft.model
            : committed?.model;
        return model ? "ok" : null;
      }, id),
    )
    .toBe("ok");
  return page.evaluate(async (docId) => {
    const db = await new Promise<IDBDatabase>((resolve, reject) => {
      const request = indexedDB.open("docu-edit");
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error);
    });
    const get = (key: string) =>
      new Promise<{ lastModified?: number; savedAt?: number; model?: unknown } | undefined>((resolve) => {
        const rq = db.transaction("documents", "readonly").objectStore("documents").get(key);
        rq.onsuccess = () => resolve(rq.result);
        rq.onerror = () => resolve(undefined);
      });
    const committed = await get(docId);
    const draft = await get("draft:" + docId);
    db.close();
    return (
      draft && (!committed || (draft.savedAt ?? 0) > (committed.lastModified ?? 0))
        ? draft.model
        : committed?.model
    ) as ModelShape;
  }, id);
}

const modelText = (model: ModelShape) =>
  model.sections[0].blocks[0].runs.map((run) => run.text).join("");

test("Tab mid-word in a list item inserts a literal tab, not a demote (AC-29, AC-30)", async ({ page }) => {
  const id = await createDocument(page);
  await content(page).click();
  await page.getByTestId("list-bulleted").click();
  await page.keyboard.type("hello world");
  // Caret collapsed mid-word (offset 5, between "hello" and " world").
  await page.keyboard.press("Home");
  for (let i = 0; i < 5; i++) {
    await page.keyboard.press("ArrowRight");
  }
  await page.keyboard.press("Tab");
  await waitForSaved(page);
  const model = await readModel(page, id);
  // Rule 4 applies (not rule 2): a literal tab, and the list level is unchanged.
  expect(modelText(model)).toBe("hello\t world");
  expect(model.sections[0].blocks[0].list?.level).toBe(1);
});

test("a literal tab renders as a sized .doc-tab box, not white-space:pre-wrap (AC-1, AC-2, AC-16)", async ({
  page,
}) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("a");
  await page.keyboard.press("ArrowRight"); // ensure mid-paragraph (offset 1)
  await page.keyboard.press("Tab");
  await page.keyboard.type("b");
  const tab = content(page).locator(".doc-tab").first();
  await expect(tab).toHaveCount(1);
  const box = await tab.boundingBox();
  expect(box).not.toBeNull();
  // The box carries a real advance — the renderer sized it, the browser's
  // default tab grid did not.
  expect(box!.width).toBeGreaterThan(4);
});

test("a left tab stop aligns following text at the stop position (AC-1, AC-17)", async ({ page }) => {
  await createDocument(page);
  await content(page).click();
  // A left stop at 5 cm on this paragraph, via the dialog.
  await page.getByTestId("menu-title-format").click();
  await page.getByTestId("format-menu-tabs").click();
  await page.getByTestId("tab-position-input").fill("5");
  await page.getByTestId("tab-alignment-select").selectOption("left");
  await page.getByTestId("tab-add").click();
  await page.getByTestId("tabs-apply").click();
  await content(page).click();
  await page.keyboard.type("x");
  await page.keyboard.press("End");
  await page.keyboard.press("Tab"); // mid-paragraph → a literal tab
  await page.keyboard.type("y");
  const tab = content(page).locator(".doc-tab").first();
  await expect(tab).toHaveCount(1);
  // The tab box's right edge (where "y" begins) sits ~5 cm (≈189 px at 96 dpi)
  // from the paragraph's content start — the stop, not the browser's grid.
  const geom = await page.evaluate(() => {
    const para = document.querySelector('.doc-paragraph[data-block="0"]') as HTMLElement;
    const box = document.querySelector(".doc-tab") as HTMLElement;
    return { paraLeft: para.getBoundingClientRect().left, tabRight: box.getBoundingClientRect().right };
  });
  expect(geom.tabRight - geom.paraLeft).toBeGreaterThan(184);
  expect(geom.tabRight - geom.paraLeft).toBeLessThan(194);
});

test("Insert -> Tab character inserts a literal tab (AC-31)", async ({ page }) => {
  const id = await createDocument(page);
  await content(page).click();
  await page.keyboard.type("ab");
  await page.getByTestId("menu-title-insert").click();
  await page.getByTestId("insert-tab-character-button").click();
  await waitForSaved(page);
  expect(modelText(await readModel(page, id))).toBe("ab\t");
});

test("the Tabs dialog offers a bar alignment and a default-interval field (AC-9, AC-10, AC-13)", async ({
  page,
}) => {
  await createDocument(page);
  await content(page).click();
  await page.getByTestId("menu-title-format").click();
  await page.getByTestId("format-menu-tabs").click();
  await expect(page.getByTestId("tabs-dialog")).toBeVisible();
  await expect(page.getByTestId("tab-alignment-select").locator("option[value='bar']")).toHaveCount(1);
  // A fresh document's default interval reads 1.27 cm (AC-13).
  await expect(page.getByTestId("tab-default-interval-input")).toHaveValue("1.27");
});

test("a bar tab stop draws a vertical rule through the paragraph (AC-22, AC-23)", async ({ page }) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("barred");
  await page.getByTestId("menu-title-format").click();
  await page.getByTestId("format-menu-tabs").click();
  await page.getByTestId("tab-position-input").fill("3");
  await page.getByTestId("tab-alignment-select").selectOption("bar");
  await page.getByTestId("tab-add").click();
  await page.getByTestId("tabs-apply").click();
  // The bar rule renders as an out-of-flow vertical rule.
  await expect(content(page).locator(".tab-bar-rule")).toHaveCount(1);
});
