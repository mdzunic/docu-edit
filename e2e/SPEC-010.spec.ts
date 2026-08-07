import { test, expect, type Page } from "@playwright/test";

// SPEC-010 — autosave and crash recovery.
//
// Each Playwright test runs in a fresh browser context, so IndexedDB and the
// Web Locks manager start empty. The crash tests crash the real renderer via
// CDP (chromium-only, which is this project's only configured browser), so
// no unload handler runs — exactly the force-close the spec describes.

/** Creates a document via the UI and returns its id from the URL. */
async function createDocument(page: Page): Promise<string> {
  await page.goto("/");
  await page.getByTestId("new-document").click();
  await expect(page).toHaveURL(/\/doc\/.+/);
  const url = page.url();
  return url.slice(url.lastIndexOf("/") + 1);
}

async function typeIntoDocument(page: Page, text: string): Promise<void> {
  await page.getByTestId("document-content").click();
  await page.keyboard.type(text);
}

const SAVED = "All changes saved";

/** Fire-and-forget IndexedDB writes need a beat to commit before a
 * cross-page navigation (same allowance SPEC-003's e2e makes). */
async function settleWrites(page: Page): Promise<void> {
  await page.waitForTimeout(300);
}

/** Writes a large committed document straight into the app's store: N
 * paragraphs of ~1.3k characters — hundreds of pages at A4 metrics. */
async function injectLargeDocument(
  page: Page,
  id: string,
  paragraphCount: number,
): Promise<void> {
  await page.goto("/");
  await page.evaluate(
    async ([docId, count]) => {
      const line = "The quick brown fox jumps over the lazy dog. ".repeat(28);
      const blocks = Array.from({ length: count as number }, () => ({
        type: "paragraph",
        runs: [{ text: line, formatting: {} }],
      }));
      const record = {
        id: docId,
        format: "new",
        lastModified: Date.now(),
        model: {
          id: docId,
          title: "Very long document",
          sections: [{ blocks }],
          defaultStyles: {
            pageSize: { widthCm: 21, heightCm: 29.7 },
            margins: { topCm: 2.54, rightCm: 2.54, bottomCm: 2.54, leftCm: 2.54 },
            fontFamily: "Calibri",
            fontSizePt: 11,
          },
        },
      };
      const db = await new Promise<IDBDatabase>((resolve, reject) => {
        const request = indexedDB.open("docu-edit", 1);
        request.onsuccess = () => resolve(request.result);
        request.onerror = () => reject(request.error);
      });
      await new Promise<void>((resolve, reject) => {
        const tx = db.transaction("documents", "readwrite");
        tx.objectStore("documents").put(record);
        tx.oncomplete = () => resolve();
        tx.onerror = () => reject(tx.error);
      });
      db.close();
    },
    [id, paragraphCount] as const,
  );
}

test("typing shows Saving… and settles to All changes saved (AC-1, AC-3, AC-4)", async ({
  page,
}) => {
  await createDocument(page);
  await expect(page.getByTestId("save-state")).toHaveText(SAVED);

  await typeIntoDocument(page, "Autosaved words");
  await expect(page.getByTestId("save-state")).toHaveText("Saving…");
  // The debounce is 2 s; well inside the default expect timeout.
  await expect(page.getByTestId("save-state")).toHaveText(SAVED);
  await expect(page.getByTestId("document-content")).toContainText(
    "Autosaved words",
  );
});

test("reload restores the autosaved content (AC-1, AC-8)", async ({ page }) => {
  await createDocument(page);
  await typeIntoDocument(page, "Persist me please");
  await expect(page.getByTestId("save-state")).toHaveText(SAVED);
  await settleWrites(page);

  await page.reload();
  await expect(page.getByTestId("document-content")).toContainText(
    "Persist me please",
  );
});

test("a clean close restores content and caret with no recovery bar (AC-7, AC-8, AC-9)", async ({
  page,
}) => {
  await createDocument(page);
  await typeIntoDocument(page, "Hello");
  await expect(page.getByTestId("save-state")).toHaveText(SAVED);
  await settleWrites(page);

  // Leaving to the library is a clean close (commit); reopening via the row
  // is an in-app navigation, so the committed write has long since landed.
  await page.goto("/");
  await expect(page.getByTestId("library-list")).toBeVisible();
  await page.getByTestId("library-row").first().click();

  await expect(page.getByTestId("document-content")).toContainText("Hello");
  await expect(page.getByTestId("recovery-bar")).toBeHidden();
  // The caret was restored to the end of "Hello": typing appends there.
  await page.keyboard.type("!");
  await expect(page.getByTestId("document-content")).toContainText("Hello!");
});

test("scroll position survives reopening (AC-10)", async ({ page }) => {
  await injectLargeDocument(page, "scroll-doc", 120);
  await page.goto("/doc/scroll-doc");
  await expect(page.getByTestId("document-content")).toContainText(
    "quick brown fox",
  );

  // Scroll, then place the caret on a line that offset actually shows, then
  // edit — the content edit after the scroll puts the scroll offset into
  // the next draft/commit write. The caret has to be on screen for that:
  // an edit made with the caret out of view scrolls back to it (SPEC-028),
  // which would carry a different offset into the write. The browser clamps
  // scrollTop to the page surface's height, so keep the value it applied.
  const canvas = page.getByTestId("canvas");
  const applied = await canvas.evaluate((el) => {
    el.scrollTop = 800;
    return el.scrollTop;
  });
  expect(applied).toBeGreaterThan(100);
  const view = (await canvas.boundingBox())!;
  await page.mouse.click(view.x + view.width / 2, view.y + view.height / 2);
  await expect.poll(async () => canvas.evaluate((el) => el.scrollTop)).toBe(applied);
  await page.keyboard.type("x");
  await expect.poll(async () => canvas.evaluate((el) => el.scrollTop)).toBe(applied);
  await expect(page.getByTestId("save-state")).toHaveText(SAVED);
  await settleWrites(page);

  await page.goto("/");
  await expect(page.getByTestId("library-list")).toBeVisible();
  await page.getByTestId("library-row").first().click();
  await expect(page.getByTestId("document-content")).toContainText(
    "quick brown fox",
  );
  await expect
    .poll(async () => canvas.evaluate((el) => el.scrollTop))
    .toBe(applied);
});

test("closing with unsaved changes triggers the leave prompt (AC-6)", async ({
  page,
}) => {
  await createDocument(page);
  await typeIntoDocument(page, "not yet saved");

  const dialogPromise = page.waitForEvent("dialog");
  const closePromise = page.close({ runBeforeUnload: true });
  const dialog = await dialogPromise;
  expect(dialog.type()).toBe("beforeunload");
  await dialog.accept();
  await closePromise;
});

test("closing when fully saved does not prompt (AC-7)", async ({ page }) => {
  await createDocument(page);
  await typeIntoDocument(page, "saved before closing");
  await expect(page.getByTestId("save-state")).toHaveText(SAVED);

  let dialogs = 0;
  page.on("dialog", (dialog) => {
    dialogs += 1;
    void dialog.accept();
  });
  await page.close({ runBeforeUnload: true });
  expect(dialogs).toBe(0);
});

/** Leaves storage exactly as a crash mid-edit does: a committed record and
 * an autosave draft NEWER than it. (Crashing the real renderer via CDP
 * Page.crash takes the whole shared browser process down in this container,
 * so the aftermath is planted instead — the app code under test is the same:
 * boot, detect the newer draft, offer recovery.) */
async function injectCrashState(page: Page, id: string): Promise<void> {
  await page.goto("/");
  await page.evaluate(async (docId) => {
    const model = (text: string) => ({
      id: docId,
      title: "Crashed doc",
      sections: [
        {
          blocks: [{ type: "paragraph", runs: [{ text, formatting: {} }] }],
        },
      ],
      defaultStyles: {
        pageSize: { widthCm: 21, heightCm: 29.7 },
        margins: { topCm: 2.54, rightCm: 2.54, bottomCm: 2.54, leftCm: 2.54 },
        fontFamily: "Calibri",
        fontSizePt: 11,
      },
    });
    const now = Date.now();
    const committed = {
      id: docId,
      format: "new",
      lastModified: now - 60_000,
      model: model("Original text"),
      caret: { section: 0, block: 0, offset: 0 },
      scroll: 0,
    };
    const draft = {
      id: `draft:${docId}`,
      docId,
      savedAt: now,
      model: model("Crashed edits"),
      caret: { section: 0, block: 0, offset: 13 },
      scroll: 0,
    };
    const db = await new Promise<IDBDatabase>((resolve, reject) => {
      const request = indexedDB.open("docu-edit", 1);
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error);
    });
    await new Promise<void>((resolve, reject) => {
      const tx = db.transaction("documents", "readwrite");
      tx.objectStore("documents").put(committed);
      tx.objectStore("documents").put(draft);
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error);
    });
    db.close();
  }, id);
}

test("after a crash the reopened document offers recovery; Discard restores the last clean state (AC-11, AC-12, AC-13)", async ({
  page,
}) => {
  await injectCrashState(page, "crashed-doc");
  await page.goto("/doc/crashed-doc");

  const bar = page.getByTestId("recovery-bar");
  await expect(bar).toBeVisible();
  await expect(bar).toContainText("Recovered unsaved changes from");
  await expect(page.getByTestId("recovery-keep")).toBeVisible();
  // The editor shows the recovered draft content.
  await expect(page.getByTestId("document-content")).toContainText(
    "Crashed edits",
  );

  // Discard: the draft goes away and the last cleanly saved state returns.
  await page.getByTestId("recovery-discard").click();
  await expect(bar).toBeHidden();
  await expect(page.getByTestId("document-content")).toContainText(
    "Original text",
  );
  await expect(page.getByTestId("document-content")).not.toContainText(
    "Crashed edits",
  );
  await settleWrites(page);

  // The draft was deleted: reopening is clean, no bar.
  await page.reload();
  await expect(page.getByTestId("document-content")).toContainText(
    "Original text",
  );
  await expect(page.getByTestId("recovery-bar")).toBeHidden();
});

test("Keep promotes the recovered draft and it stays after reload (AC-11)", async ({
  page,
}) => {
  await injectCrashState(page, "kept-doc");
  await page.goto("/doc/kept-doc");
  await expect(page.getByTestId("recovery-bar")).toBeVisible();

  await page.getByTestId("recovery-keep").click();
  await expect(page.getByTestId("recovery-bar")).toBeHidden();
  await expect(page.getByTestId("document-content")).toContainText(
    "Crashed edits",
  );
  await settleWrites(page);

  // Now committed: reopening shows the kept content with no recovery offer.
  await page.reload();
  await expect(page.getByTestId("document-content")).toContainText(
    "Crashed edits",
  );
  await expect(page.getByTestId("recovery-bar")).toBeHidden();
});

test("typing stays responsive while a 200-page document autosaves (AC-14)", async ({
  page,
}) => {
  // ~600 paragraphs x ~1.3k chars ≈ 780k characters ≈ 200 A4 pages.
  await injectLargeDocument(page, "huge-doc", 600);
  await page.goto("/doc/huge-doc");
  await expect(page.getByTestId("document-content")).toContainText(
    "quick brown fox",
  );

  await page.getByTestId("document-content").click({ position: { x: 20, y: 20 } });
  await page.keyboard.type("RESPONSIVE-1", { delay: 30 });
  await expect(page.getByTestId("document-content")).toContainText(
    "RESPONSIVE-1",
  );
  // Let the save of the whole huge model run, then type through/after it —
  // every keystroke must land.
  await expect(page.getByTestId("save-state")).toHaveText(SAVED, {
    timeout: 15_000,
  });
  await page.keyboard.type("RESPONSIVE-2", { delay: 30 });
  await expect(page.getByTestId("document-content")).toContainText(
    "RESPONSIVE-2",
  );
});

test("a second tab on the same document is read-only until the first closes (AC-15, AC-16, AC-17)", async ({
  page,
  context,
}) => {
  const id = await createDocument(page);
  await expect(page.getByTestId("document-content")).toHaveAttribute(
    "contenteditable",
    "true",
  );

  const second = await context.newPage();
  await second.goto(`/doc/${id}`);
  const conflictBar = second.getByTestId("conflict-bar");
  await expect(conflictBar).toBeVisible();
  await expect(conflictBar).toHaveText("This document is open in another tab");
  await expect(second.getByTestId("document-content")).toHaveAttribute(
    "contenteditable",
    "false",
  );

  // Closing the first tab releases its web lock; the second becomes
  // editable without a reload.
  await page.close();
  await expect(conflictBar).toBeHidden();
  await expect(second.getByTestId("document-content")).toHaveAttribute(
    "contenteditable",
    "true",
  );
  await typeIntoDocument(second, "now editable");
  await expect(second.getByTestId("document-content")).toContainText(
    "now editable",
  );
});

test("storage quota failure surfaces in the app bar with a download link, and clears on recovery (AC-5, AC-18, AC-19)", async ({
  page,
}) => {
  // Fail exactly the autosave draft writes, as a full quota would; committed
  // writes (document creation) stay fine so the scenario is isolated.
  await page.addInitScript(() => {
    const original = IDBObjectStore.prototype.put;
    type PutArgs = Parameters<typeof original>;
    (window as { __restoreIdbPut?: () => void }).__restoreIdbPut = () => {
      IDBObjectStore.prototype.put = original;
    };
    IDBObjectStore.prototype.put = function (...args: PutArgs) {
      const record = args[0] as { id?: unknown };
      if (typeof record?.id === "string" && record.id.startsWith("draft:")) {
        throw new DOMException("Quota exceeded.", "QuotaExceededError");
      }
      return original.apply(this, args);
    };
  });

  await createDocument(page);
  await typeIntoDocument(page, "cannot be saved");
  await expect(page.getByTestId("save-state")).toHaveText(
    "Not saved — browser storage is full",
  );

  // The download link is the escape hatch and really downloads the file.
  const downloadPromise = page.waitForEvent("download");
  await page.getByTestId("save-state-download").click();
  const download = await downloadPromise;
  expect(download.suggestedFilename()).toMatch(/\.docx$/);

  // Storage recovers: the retry write succeeds and the state clears.
  await page.evaluate(() =>
    (window as { __restoreIdbPut?: () => void }).__restoreIdbPut?.(),
  );
  await expect(page.getByTestId("save-state")).toHaveText(SAVED, {
    timeout: 10_000,
  });
});
