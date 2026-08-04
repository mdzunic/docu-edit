import { test, expect, type Page } from "@playwright/test";

// SPEC-011 — the rich text editing core: caret placement, keyboard and
// mouse selection, mutation through the command API, IME composition and
// the assistive-technology surface. The caret-status live region doubles
// as the observable caret position: it renders exactly what a screen
// reader is told, in model coordinates.

/** Creates a document via the UI and returns its id from the URL. */
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
const status = (page: Page) => page.getByTestId("caret-status");
const paragraphs = (page: Page) =>
  page.locator('[data-testid="document-content"] [data-section]');

async function selectionText(page: Page): Promise<string> {
  return page.evaluate(() => window.getSelection()?.toString() ?? "");
}

/** The 1-based paragraph ordinal the caret status currently announces. */
async function caretParagraph(page: Page): Promise<number> {
  const text = (await status(page).textContent()) ?? "";
  const match = /[Pp]aragraph (\d+)/.exec(text);
  return match ? Number(match[1]) : NaN;
}

/** The newest persisted model (draft beats committed when newer), so tests
 * can compare what is on screen with what the model actually contains. */
async function readLatestModel(page: Page, id: string) {
  return page.evaluate(async (docId) => {
    const db = await new Promise<IDBDatabase>((resolve, reject) => {
      const request = indexedDB.open("docu-edit", 1);
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
  }, id);
}

interface ModelShape {
  sections: Array<{
    blocks: Array<{
      formatting?: { alignment?: string };
      runs: Array<{ text: string }>;
    }>;
  }>;
}

function modelParagraphs(model: unknown): string[] {
  return (model as ModelShape).sections.flatMap((section) =>
    section.blocks.map((block) => block.runs.map((run) => run.text).join("")),
  );
}

/** Writes a committed document straight into the app's store. */
async function injectDocument(
  page: Page,
  id: string,
  blocks: Array<{ text: string; alignment?: string }>,
): Promise<void> {
  await page.goto("/");
  await page.evaluate(
    async ([docId, blockSpecs]) => {
      const specs = blockSpecs as Array<{ text: string; alignment?: string }>;
      const record = {
        id: docId,
        format: "new",
        lastModified: Date.now(),
        model: {
          id: docId,
          title: "Injected document",
          sections: [
            {
              blocks: specs.map((spec) => ({
                type: "paragraph",
                runs: [{ text: spec.text, formatting: {} }],
                ...(spec.alignment
                  ? { formatting: { alignment: spec.alignment } }
                  : {}),
              })),
            },
          ],
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
    [id, blocks] as const,
  );
}

/** Types two paragraphs: "alpha beta" and "gamma delta". */
async function typeTwoParagraphs(page: Page): Promise<void> {
  await content(page).click();
  await page.keyboard.type("alpha beta");
  await page.keyboard.press("Enter");
  await page.keyboard.type("gamma delta");
}

test("clicking the page canvas (margins included) places a visible caret at the nearest text position (AC-1, AC-2)", async ({
  page,
}) => {
  await createDocument(page);
  await typeTwoParagraphs(page);
  const pageBox = (await page.getByTestId("document-page").boundingBox())!;
  const firstBox = (await paragraphs(page).first().boundingBox())!;
  const secondBox = (await paragraphs(page).nth(1).boundingBox())!;

  // Left page margin beside the second paragraph: the caret goes to its
  // nearest position — the start of that paragraph's line.
  await page.mouse.click(pageBox.x + 20, secondBox.y + secondBox.height / 2);
  await expect(status(page)).toHaveText("Paragraph 2 of 2, character 1");

  // Right page margin beside the first paragraph: its line's end.
  await page.mouse.click(
    pageBox.x + pageBox.width - 20,
    firstBox.y + firstBox.height / 2,
  );
  await expect(status(page)).toHaveText("Paragraph 1 of 2, character 11");

  // Empty page area below all text: the last text position. The A4 page is
  // taller than the viewport, so bring its bottom edge on screen first.
  await page
    .getByTestId("document-page")
    .evaluate((el) => el.scrollIntoView({ block: "end" }));
  const scrolledBox = (await page.getByTestId("document-page").boundingBox())!;
  await page.mouse.click(
    scrolledBox.x + scrolledBox.width / 2,
    scrolledBox.y + scrolledBox.height - 30,
  );
  await expect(status(page)).toHaveText("Paragraph 2 of 2, character 12");

  // The caret is a real, visible caret: a collapsed DOM selection focused
  // inside the editable canvas.
  const caretState = await page.evaluate(() => {
    const selection = window.getSelection();
    const root = document.querySelector('[data-testid="document-content"]');
    return {
      collapsed: selection?.isCollapsed ?? false,
      inside: !!selection?.focusNode && !!root?.contains(selection.focusNode),
      focused: document.activeElement === root,
    };
  });
  expect(caretState).toEqual({ collapsed: true, inside: true, focused: true });
});

test("typing inserts at the caret and the model matches the screen (AC-3, AC-4, AC-26)", async ({
  page,
}) => {
  const id = await createDocument(page);
  await content(page).click();
  await page.keyboard.type("Helo world");
  await page.keyboard.press("Home");
  for (let i = 0; i < 2; i += 1) {
    await page.keyboard.press("ArrowRight");
  }
  await page.keyboard.type("l");
  await expect(content(page)).toHaveText("Hello world");
  await waitForSaved(page);
  await expect
    .poll(async () => modelParagraphs(await readLatestModel(page, id)))
    .toEqual(["Hello world"]);
});

test("Backspace and Delete remove characters and merge paragraphs at boundaries (AC-5, AC-6, AC-7)", async ({
  page,
}) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("ab");
  await page.keyboard.press("Enter");
  await page.keyboard.type("cd");
  await expect(paragraphs(page)).toHaveCount(2);

  // Backspace at the start of "cd" merges the paragraphs (AC-7).
  await page.keyboard.press("Home");
  await page.keyboard.press("Backspace");
  await expect(paragraphs(page)).toHaveCount(1);
  await expect(content(page)).toHaveText("abcd");

  // The caret sits between b and c: Delete removes "c" (AC-6)…
  await page.keyboard.press("Delete");
  await expect(content(page)).toHaveText("abd");
  // …and Backspace removes "b" (AC-5).
  await page.keyboard.press("Backspace");
  await expect(content(page)).toHaveText("ad");

  // Delete at a paragraph end also merges (AC-7 forward).
  await page.keyboard.press("End");
  await page.keyboard.press("Enter");
  await page.keyboard.type("tail");
  await expect(paragraphs(page)).toHaveCount(2);
  await page.keyboard.press("Home");
  await page.keyboard.press("ArrowUp");
  await page.keyboard.press("End");
  await page.keyboard.press("Delete");
  await expect(paragraphs(page)).toHaveCount(1);
  await expect(content(page)).toHaveText("adtail");
});

test("Enter splits carrying paragraph formatting; Shift+Enter breaks the line in place (AC-8, AC-9, AC-10)", async ({
  page,
}) => {
  const id = "split-formatting";
  await injectDocument(page, id, [
    { text: "Centered paragraph", alignment: "center" },
  ]);
  await page.goto(`/doc/${id}`);
  await content(page).click();
  await page.keyboard.press("Home");
  for (let i = 0; i < "Centered".length; i += 1) {
    await page.keyboard.press("ArrowRight");
  }
  await page.keyboard.press("Enter");
  await expect(paragraphs(page)).toHaveCount(2);

  // The new paragraph renders with the same alignment (AC-9, on screen)…
  const alignments = await paragraphs(page).evaluateAll((elements) =>
    elements.map((el) => getComputedStyle(el).textAlign),
  );
  expect(alignments).toEqual(["center", "center"]);

  // …and carries it in the model.
  await waitForSaved(page);
  await expect
    .poll(async () => {
      const model = (await readLatestModel(page, id)) as ModelShape;
      return model.sections[0].blocks.map((b) => b.formatting?.alignment);
    })
    .toEqual(["center", "center"]);

  // Shift+Enter adds a line inside the second paragraph, not a new one.
  await page.keyboard.press("End");
  await page.keyboard.press("Shift+Enter");
  await page.keyboard.type("wrapped");
  await expect(paragraphs(page)).toHaveCount(2);
  await waitForSaved(page);
  await expect
    .poll(async () => modelParagraphs(await readLatestModel(page, id)))
    .toEqual(["Centered", " paragraph\nwrapped"]);
});

test("arrow keys, Home/End and word jumps move the caret; Shift extends (AC-11, AC-12, AC-15, AC-16)", async ({
  page,
}) => {
  await createDocument(page);
  await typeTwoParagraphs(page);
  await page.keyboard.press("Control+Home");
  await expect(status(page)).toHaveText("Paragraph 1 of 2, character 1");

  await page.keyboard.press("ArrowRight");
  await expect(status(page)).toHaveText("Paragraph 1 of 2, character 2");

  await page.keyboard.press("ArrowDown");
  await expect(status(page)).toContainText("Paragraph 2 of 2");

  await page.keyboard.press("End");
  await expect(status(page)).toHaveText("Paragraph 2 of 2, character 12");
  await page.keyboard.press("Home");
  await expect(status(page)).toHaveText("Paragraph 2 of 2, character 1");

  // Ctrl+Right jumps by word: past "gamma" (to its end or the next word's
  // start, depending on platform convention — both are word jumps).
  await page.keyboard.press("Control+ArrowRight");
  await expect(status(page)).toHaveText(/Paragraph 2 of 2, character (6|7)/);

  // Shift extends any movement (AC-15).
  await page.keyboard.press("Shift+End");
  await expect(status(page)).toContainText("Selected from paragraph 2");
  expect(await selectionText(page)).toMatch(/delta$/);
});

test("Ctrl+Home/End and PageUp/PageDown traverse a long document (AC-13, AC-14, AC-15)", async ({
  page,
}) => {
  const id = "long-doc";
  const line = "The quick brown fox jumps over the lazy dog. ".repeat(20);
  await injectDocument(
    page,
    id,
    Array.from({ length: 60 }, () => ({ text: line })),
  );
  await page.goto(`/doc/${id}`);
  await content(page).first().waitFor();
  await content(page).click();

  await page.keyboard.press("Control+End");
  await expect(status(page)).toContainText("Paragraph 60 of 60");

  await page.keyboard.press("Control+Home");
  await expect(status(page)).toHaveText("Paragraph 1 of 60, character 1");

  // PageDown moves the caret (and the view) about a viewport forward.
  await page.keyboard.press("PageDown");
  await expect.poll(() => caretParagraph(page)).toBeGreaterThan(1);
  const afterPageDown = await caretParagraph(page);
  const scrolled = await page
    .locator("main")
    .evaluate((main) => main.scrollTop);
  expect(scrolled).toBeGreaterThan(0);

  await page.keyboard.press("PageUp");
  await expect.poll(() => caretParagraph(page)).toBeLessThan(afterPageDown);

  // Shift+PageDown extends a selection while paging (AC-15).
  await page.keyboard.press("Control+Home");
  await page.keyboard.press("Shift+PageDown");
  await expect(status(page)).toContainText("Selected from paragraph 1");
});

test("double-click selects a word, triple-click a paragraph, Ctrl+A the document (AC-17, AC-18, AC-19)", async ({
  page,
}) => {
  await createDocument(page);
  await typeTwoParagraphs(page);
  const firstBox = (await paragraphs(page).first().boundingBox())!;

  await page.mouse.dblclick(firstBox.x + 8, firstBox.y + firstBox.height / 2);
  expect((await selectionText(page)).trim()).toBe("alpha");

  await page.mouse.click(firstBox.x + 8, firstBox.y + firstBox.height / 2, {
    clickCount: 3,
  });
  expect((await selectionText(page)).trim()).toBe("alpha beta");

  await page.keyboard.press("Control+a");
  await expect(status(page)).toHaveText(
    "Selected from paragraph 1, character 1 to paragraph 2, character 12",
  );
  const all = await selectionText(page);
  expect(all).toContain("alpha beta");
  expect(all).toContain("gamma delta");
});

test("mouse drag creates a visibly highlighted selection across paragraphs (AC-20, AC-21, AC-22)", async ({
  page,
}) => {
  await createDocument(page);
  await typeTwoParagraphs(page);
  const firstBox = (await paragraphs(page).first().boundingBox())!;
  const secondBox = (await paragraphs(page).nth(1).boundingBox())!;

  // The drag is paced like a real one: headless Chromium only reliably
  // extends a drag selection across block boundaries when the renderer gets
  // to run between the synthetic mouse moves, so each step yields to the
  // page before the next.
  const from = { x: firstBox.x + 5, y: firstBox.y + firstBox.height / 2 };
  const to = {
    x: secondBox.x + secondBox.width / 3,
    y: secondBox.y + secondBox.height / 2,
  };
  await page.mouse.move(from.x, from.y);
  await page.mouse.down();
  for (let i = 1; i <= 4; i += 1) {
    await page.mouse.move(
      from.x + ((to.x - from.x) * i) / 4,
      from.y + ((to.y - from.y) * i) / 4,
    );
    await page.evaluate(() => window.getSelection()?.toString());
  }
  await page.mouse.up();

  // The drag started inside the first word's glyphs, so the selection spans
  // from within "alpha" across the boundary into "gamma delta".
  const selected = await selectionText(page);
  expect(selected).toContain("lpha beta");
  expect(selected).toContain("gamma");
  await expect(status(page)).toContainText("Selected from paragraph 1");
  await expect(status(page)).toContainText("to paragraph 2");
  // The highlight is the DOM selection the browser paints; it must span
  // real rendered rects, not be collapsed.
  const rectCount = await page.evaluate(
    () => window.getSelection()?.getRangeAt(0).getClientRects().length ?? 0,
  );
  expect(rectCount).toBeGreaterThan(1);
});

test("typing with a non-empty selection replaces it (AC-23)", async ({ page }) => {
  const id = await createDocument(page);
  await content(page).click();
  await page.keyboard.type("alpha beta");
  const box = (await paragraphs(page).first().boundingBox())!;
  await page.mouse.dblclick(box.x + 8, box.y + box.height / 2);
  expect((await selectionText(page)).trim()).toBe("alpha");

  await page.keyboard.type("omega");
  await expect(content(page)).toHaveText("omega beta");
  await waitForSaved(page);
  await expect
    .poll(async () => modelParagraphs(await readLatestModel(page, id)))
    .toEqual(["omega beta"]);
});

test("IME composition events commit once without corrupting the model (AC-27, AC-28, AC-29)", async ({
  page,
}) => {
  const id = await createDocument(page);
  await content(page).click();
  await page.keyboard.type("hello ");
  await page.evaluate(() => {
    const root = document.querySelector('[data-testid="document-content"]')!;
    const fire = (type: string, data: string) =>
      root.dispatchEvent(new CompositionEvent(type, { bubbles: true, data }));
    fire("compositionstart", "");
    fire("compositionupdate", "にほんご");
    fire("compositionend", "日本語");
  });
  await expect(content(page)).toHaveText("hello 日本語");
  await expect(status(page)).toHaveText("Paragraph 1 of 1, character 10");

  // The model holds exactly the committed text — composition neither
  // duplicated nor dropped anything.
  await waitForSaved(page);
  await expect
    .poll(async () => modelParagraphs(await readLatestModel(page, id)))
    .toEqual(["hello 日本語"]);

  // And typing afterwards continues from the committed caret.
  await page.keyboard.type("!");
  await expect(content(page)).toHaveText("hello 日本語!");
});

test("dead-key composition produces the single accented character (AC-30)", async ({
  page,
}) => {
  const id = await createDocument(page);
  await content(page).click();
  await page.keyboard.type("caf");
  await page.evaluate(() => {
    const root = document.querySelector('[data-testid="document-content"]')!;
    const fire = (type: string, data: string) =>
      root.dispatchEvent(new CompositionEvent(type, { bubbles: true, data }));
    fire("compositionstart", "");
    fire("compositionupdate", "´");
    fire("compositionend", "é");
  });
  await expect(content(page)).toHaveText("café");
  await waitForSaved(page);
  await expect
    .poll(async () => modelParagraphs(await readLatestModel(page, id)))
    .toEqual(["café"]);
});

test("the editable region is an accessible textbox that announces caret movement (AC-31, AC-32, AC-33, AC-34)", async ({
  page,
}) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("Readable content");

  // AC-33/34: a multiline textbox whose accessible content is the text.
  const textbox = page.getByRole("textbox", { name: "Document content" });
  await expect(textbox).toBeVisible();
  await expect(textbox).toHaveText("Readable content");

  // AC-32: caret movement is announced through a polite live region.
  const live = status(page);
  await expect(live).toHaveAttribute("role", "status");
  await expect(live).toHaveAttribute("aria-live", "polite");
  await page.keyboard.press("Home");
  await expect(live).toHaveText("Paragraph 1 of 1, character 1");
  await page.keyboard.press("ArrowRight");
  await expect(live).toHaveText("Paragraph 1 of 1, character 2");
});
