import { test, expect, type Page } from "@playwright/test";
import { MOD, LINE_START } from "./keys";

// SPEC-016 — cut, copy and paste: keyboard shortcuts and context menu
// entries, the three clipboard flavours, exact internal round-trip paste,
// sanitised external-HTML import (Word Online / Google Docs / web page),
// plain-text paste, and the paste options chip's single-undo-step
// re-application. Model-level assertions read the persisted model from
// IndexedDB like e2e/SPEC-012.spec.ts; external clipboards are simulated by
// dispatching ClipboardEvents with a DataTransfer, which is exactly what the
// browser's own paste path does. Pasting an actual image is SPEC-024's
// (e2e/SPEC-024.spec.ts) — this file keeps only the malformed-image case,
// since an accepted image type with unreadable bytes is still possible
// (a truncated/corrupt file) regardless of that spec.

const DOC_MODEL_MIME = "application/x-docu-edit-docmodel+json";

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

interface ModelRun {
  text: string;
  formatting: Record<string, unknown>;
}

interface ModelBlock {
  runs: ModelRun[];
  formatting?: Record<string, unknown>;
  headingLevel?: number;
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

function blocks(model: ModelShape): ModelBlock[] {
  return model.sections[0].blocks;
}

function blockText(block: ModelBlock): string {
  return block.runs.map((run) => run.text).join("");
}

/** Selects a substring of the current line by character offsets — the same
 * Home + arrow-key technique as e2e/SPEC-012.spec.ts. */
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

/** Dispatches a paste ClipboardEvent carrying the given flavours — the same
 * event/DataTransfer shape the browser produces for a real paste, so the
 * whole handler pipeline runs. */
async function dispatchPaste(page: Page, flavours: Record<string, string>): Promise<void> {
  await page.evaluate((data) => {
    const transfer = new DataTransfer();
    for (const [type, value] of Object.entries(data)) {
      transfer.setData(type, value);
    }
    const target = document.querySelector('[data-testid="document-content"]');
    target?.dispatchEvent(
      new ClipboardEvent("paste", {
        clipboardData: transfer,
        bubbles: true,
        cancelable: true,
      }),
    );
  }, flavours);
}

test("Ctrl+X cuts in one undo step and Ctrl+V pastes the cut content back (AC-1, AC-3, AC-36..38)", async ({
  page,
}) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("hello world");
  await selectRange(page, 0, 6);
  await page.keyboard.press(`${MOD}+x`);
  await expect(content(page)).toHaveText("world");
  await expect(page.getByTestId("undo-button")).toHaveAttribute("title", /Undo cut/);

  // The removed content is on the clipboard (AC-37): paste it at the end.
  await page.keyboard.press(`${MOD}+End`);
  await page.keyboard.press(`${MOD}+v`);
  await expect(content(page)).toHaveText("worldhello ");

  // The cut was a single step (AC-38): two undos (paste, then cut) restore.
  await page.keyboard.press(`${MOD}+z`);
  await expect(content(page)).toHaveText("world");
  await page.keyboard.press(`${MOD}+z`);
  await expect(content(page)).toHaveText("hello world");
});

test("copying puts plain text, HTML and the internal payload on the clipboard (AC-2, AC-7..9)", async ({
  page,
}) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("hello world");
  await selectRange(page, 6, 11);
  await page.keyboard.press(`${MOD}+b`);
  await page.keyboard.press(`${MOD}+a`);
  await page.evaluate((mime) => {
    document.addEventListener("copy", (event) => {
      const data = (event as ClipboardEvent).clipboardData;
      (window as unknown as Record<string, unknown>).__flavours = data
        ? {
            text: data.getData("text/plain"),
            html: data.getData("text/html"),
            internal: data.getData(mime),
          }
        : null;
    });
  }, DOC_MODEL_MIME);
  await page.keyboard.press(`${MOD}+c`);
  const flavours = await page.evaluate(
    () => (window as unknown as Record<string, unknown>).__flavours,
  ) as { text: string; html: string; internal: string };
  expect(flavours.text).toBe("hello world");
  expect(flavours.html).toBe("<p>hello <b>world</b></p>");
  const payload = JSON.parse(flavours.internal);
  expect(payload.mimeType).toBe(DOC_MODEL_MIME);
  expect(payload.version).toBe(1);
  expect(payload.blocks).toEqual([
    {
      type: "paragraph",
      runs: [
        { text: "hello ", formatting: {} },
        { text: "world", formatting: { bold: true } },
      ],
    },
  ]);
});

test("pasting the editor's own copy restores formatting and structure exactly (AC-10..13, AC-29)", async ({
  page,
}) => {
  const id = await createDocument(page);
  await content(page).click();
  await page.keyboard.type("Alpha");
  await page.getByTestId("align-center").click();
  await page.keyboard.press(`${MOD}+End`);
  await page.keyboard.press("Enter");
  await page.keyboard.type("bold line");
  await selectRange(page, 0, 9);
  await page.keyboard.press(`${MOD}+b`);

  await page.keyboard.press(`${MOD}+a`);
  await page.keyboard.press(`${MOD}+c`);
  await page.keyboard.press(`${MOD}+End`);
  await page.keyboard.press("Enter");
  await page.keyboard.press(`${MOD}+v`);
  await expect(content(page)).toHaveText("Alphabold lineAlphabold line");
  // The chip appears after the paste (AC-29).
  await expect(page.getByTestId("paste-options-chip")).toBeVisible();
  await waitForSaved(page);

  const model = await readLatestModel(page, id);
  const all = blocks(model);
  expect(all).toHaveLength(4);
  expect(all[2].formatting).toEqual({ alignment: "center" });
  expect(blockText(all[2])).toBe("Alpha");
  expect(all[3].runs).toEqual([{ text: "bold line", formatting: { bold: true } }]);
});

test("Word Online HTML maps headings and bold to the model (AC-14, AC-17)", async ({ page }) => {
  const id = await createDocument(page);
  await content(page).click();
  await dispatchPaste(page, {
    "text/html":
      '<div class="OutlineElement Ltr SCXW100"><h1 class="Heading1 SCXW100">Word heading</h1></div>' +
      '<div class="OutlineElement Ltr SCXW100"><p class="Paragraph SCXW100" style="margin:0">' +
      '<span class="TextRun SCXW100" style="font-weight:bold"><span class="NormalTextRun">bold text</span></span></p></div>',
    "text/plain": "Word heading\nbold text",
  });
  await expect(content(page).locator("h1")).toHaveText("Word heading");
  await waitForSaved(page);
  const all = blocks(await readLatestModel(page, id));
  expect(all[0].headingLevel).toBe(1);
  expect(all[1].runs[0]).toEqual({ text: "bold text", formatting: { bold: true } });
  expect(JSON.stringify(all)).not.toContain("SCXW100");
});

test("Google Docs HTML maps headings and CSS bold/italic without bolding everything (AC-15, AC-17, AC-18)", async ({
  page,
}) => {
  const id = await createDocument(page);
  await content(page).click();
  await dispatchPaste(page, {
    "text/html":
      '<meta charset="utf-8"><b style="font-weight:normal" id="docs-internal-guid-abc">' +
      '<h2 dir="ltr"><span style="font-size:16pt;font-weight:400">Doc heading</span></h2>' +
      '<p dir="ltr"><span style="font-weight:700">bold</span><span style="font-style:italic">italic</span></p></b>',
    "text/plain": "Doc heading\nbolditalic",
  });
  await expect(content(page).locator("h2")).toHaveText("Doc heading");
  await waitForSaved(page);
  const all = blocks(await readLatestModel(page, id));
  expect(all[0].headingLevel).toBe(2);
  expect(all[0].runs[0].formatting.bold).toBeUndefined();
  expect(all[1].runs).toEqual([
    { text: "bold", formatting: { bold: true } },
    { text: "italic", formatting: { italic: true } },
  ]);
});

test("web page HTML maps headings, u/lists/links/tables per the deferral table (AC-16, AC-19..22)", async ({
  page,
}) => {
  const id = await createDocument(page);
  await content(page).click();
  await dispatchPaste(page, {
    "text/html":
      "<h3>Web heading</h3>" +
      "<p><u>underlined</u> and <a href=\"https://example.com/x\">a link</a></p>" +
      "<ul><li>first item</li><li>second item</li></ul>" +
      "<table><tr><td>a1</td><td>b1</td></tr><tr><td>a2</td><td>b2</td></tr></table>",
  });
  await waitForSaved(page);
  const all = blocks(await readLatestModel(page, id));
  expect(all[0].headingLevel).toBe(3);
  expect(blockText(all[0])).toBe("Web heading");
  expect(all[1].runs[0]).toEqual({ text: "underlined", formatting: { underline: true } });
  // The link keeps its text, loses its URL (SPEC-025 remaps later).
  expect(blockText(all[1])).toBe("underlined and a link");
  // List items import as plain paragraphs (SPEC-018 remaps later).
  expect(blockText(all[2])).toBe("first item");
  expect(blockText(all[3])).toBe("second item");
  expect(all[2].headingLevel).toBeUndefined();
  // Table rows flatten to tab-separated paragraphs (SPEC-022 remaps later).
  expect(blockText(all[4])).toBe("a1\tb1");
  expect(blockText(all[5])).toBe("a2\tb2");
  expect(JSON.stringify(all)).not.toContain("example.com");
});

test("hostile HTML is sanitised: scripts, styles, handlers, iframes, remote images, javascript: URLs (AC-23..26, AC-39..44)", async ({
  page,
}) => {
  const id = await createDocument(page);
  await content(page).click();
  await dispatchPaste(page, {
    "text/html":
      "<script>window.__xss = 1</script>" +
      '<style onclick="window.__xss = 2">p { color: red }</style>' +
      '<iframe src="https://evil.example/frame"></iframe>' +
      '<img src="https://evil.example/tracker.png">' +
      '<p class="Mso" style="position:absolute;top:0" onmouseover="window.__xss = 3">safe ' +
      '<a href="javascript:window.__xss = 4">link text</a></p>',
  });
  await expect(content(page)).toHaveText("safe link text");
  // Nothing active or remote reached the canvas DOM.
  expect(await content(page).locator("script, style, iframe, img, a").count()).toBe(0);
  const xss = await page.evaluate(
    () => (window as unknown as Record<string, unknown>).__xss,
  );
  expect(xss).toBeUndefined();
  await waitForSaved(page);
  const serialised = JSON.stringify(blocks(await readLatestModel(page, id)));
  expect(serialised).not.toContain("__xss");
  expect(serialised).not.toContain("evil.example");
  expect(serialised).not.toContain("javascript:");
  expect(serialised).not.toContain("Mso");
  expect(serialised).not.toContain("absolute");
  // The dropped <img> announces the image gate (AC-46).
  await expect(page.getByTestId("toast")).toHaveText("Images aren't supported yet");
});

test("Ctrl+Shift+V pastes as plain text with the caret paragraph's formatting (AC-27, AC-28)", async ({
  page,
  context,
}) => {
  // Where the browser turns this combination into a paste event (Windows,
  // Linux) the event path runs and the permission is never consulted; where
  // it does not (macOS, where the combination belongs to the application
  // menu) the editor reads the clipboard itself, which is the grant a real
  // user gives at the prompt.
  await context.grantPermissions(["clipboard-read", "clipboard-write"]);
  const id = await createDocument(page);
  await content(page).click();
  await page.keyboard.type("styled");
  await selectRange(page, 0, 6);
  await page.keyboard.press(`${MOD}+b`);
  await selectRange(page, 0, 6);
  await page.keyboard.press(`${MOD}+c`);
  await page.keyboard.press(`${MOD}+End`);
  await page.keyboard.press("Enter");
  // The new paragraph starts unformatted (fresh Enter at the end of a bold
  // run keeps the bold run behind, but typing position formatting comes
  // from the empty paragraph) — paste plain and type context around it.
  await page.keyboard.press(`${MOD}+Shift+v`);
  await expect(content(page)).toHaveText("styledstyled");
  await waitForSaved(page);
  const all = blocks(await readLatestModel(page, id));
  expect(all[0].runs).toEqual([{ text: "styled", formatting: { bold: true } }]);
  // The plain-text paste carries the destination formatting, not the bold
  // the clipboard's rich flavours describe.
  expect(all[1].runs[0].text).toBe("styled");
  expect(all[1].runs[0].formatting.bold).toBeUndefined();
});

test("pasting into a selection replaces it (AC-34)", async ({ page }) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("one two three");
  await selectRange(page, 4, 7);
  await dispatchPaste(page, { "text/plain": "TWO" });
  await expect(content(page)).toHaveText("one TWO three");
});

test("pasting multiple paragraphs mid-paragraph splits it correctly (AC-35)", async ({ page }) => {
  const id = await createDocument(page);
  await content(page).click();
  await page.keyboard.type("startend");
  await selectRange(page, 5, 5);
  await dispatchPaste(page, { "text/plain": "one\ntwo\nthree" });
  await waitForSaved(page);
  const all = blocks(await readLatestModel(page, id));
  expect(all.map(blockText)).toEqual(["startone", "two", "threeend"]);
});

test("the paste chip offers the three options and re-applies as one undo step (AC-29..33)", async ({
  page,
}) => {
  const id = await createDocument(page);
  await content(page).click();
  await page.keyboard.type("seed");
  await dispatchPaste(page, {
    "text/html": "<h1><b>Fancy</b></h1><p>plain</p>",
    "text/plain": "Fancy\nplain",
  });
  await expect(content(page)).toHaveText("seedFancyplain");
  const chip = page.getByTestId("paste-options-chip");
  await expect(chip).toBeVisible();
  await expect(page.getByTestId("paste-option-source")).toHaveText("Keep source formatting");
  await expect(page.getByTestId("paste-option-match")).toHaveText("Match destination formatting");
  await expect(page.getByTestId("paste-option-text")).toHaveText("Keep text only");

  await page.getByTestId("paste-option-text").click();
  await expect(chip).toBeVisible();
  await waitForSaved(page);
  let all = blocks(await readLatestModel(page, id));
  expect(all.map(blockText)).toEqual(["seedFancy", "plain"]);
  // Keep text only stripped the heading and the bold (AC-32).
  expect(all[0].headingLevel).toBeUndefined();
  expect(all[0].runs.every((run) => !run.formatting.bold)).toBe(true);

  // Switching back re-applies the import's own fidelity (AC-30). The first
  // pasted block merges into the "seed" paragraph (whose paragraph identity
  // wins — see pasteBlocks), so the fidelity that returns is the run-level
  // bold.
  await page.getByTestId("paste-option-source").click();
  await waitForSaved(page);
  all = blocks(await readLatestModel(page, id));
  expect(all[0].runs).toEqual([
    { text: "seed", formatting: {} },
    { text: "Fancy", formatting: { bold: true } },
  ]);

  // The paste and its re-applications are ONE undo step (AC-33): a single
  // undo returns to the pre-paste document.
  await page.keyboard.press(`${MOD}+z`);
  await expect(content(page)).toHaveText("seed");
});

test("'Match destination formatting' re-resolves the paste to the destination (AC-31)", async ({
  page,
}) => {
  const id = await createDocument(page);
  await content(page).click();
  await page.keyboard.type("context");
  await dispatchPaste(page, {
    "text/html": '<h2><span style="color:#ff0000">Loud</span></h2>',
    "text/plain": "Loud",
  });
  await expect(content(page)).toHaveText("contextLoud");
  await page.getByTestId("paste-option-match").click();
  await waitForSaved(page);
  const all = blocks(await readLatestModel(page, id));
  expect(all).toHaveLength(1);
  expect(all[0].headingLevel).toBeUndefined();
  expect(all[0].runs).toEqual([{ text: "contextLoud", formatting: {} }]);
});

test("pasting a malformed image from the clipboard is rejected, not inserted (SPEC-024)", async ({
  page,
}) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("keep");
  // Only the PNG signature bytes, no IHDR/IDAT — an accepted MIME type
  // (image/png) that still fails to decode, the one case SPEC-024's paste
  // path can reject regardless of type (a truncated/corrupt file).
  await page.evaluate(() => {
    const transfer = new DataTransfer();
    transfer.items.add(
      new File([new Uint8Array([137, 80, 78, 71])], "shot.png", { type: "image/png" }),
    );
    const target = document.querySelector('[data-testid="document-content"]');
    target?.dispatchEvent(
      new ClipboardEvent("paste", {
        clipboardData: transfer,
        bubbles: true,
        cancelable: true,
      }),
    );
  });
  await expect(page.getByTestId("toast")).toHaveText("That image could not be read");
  await expect(content(page)).toHaveText("keep");
});

test.describe("context menu", () => {
  test.use({ permissions: ["clipboard-read", "clipboard-write"] });

  test("right-click opens the menu; Cut removes and fills the clipboard (AC-4)", async ({
    page,
  }) => {
    await createDocument(page);
    await content(page).click();
    await page.keyboard.type("hello world");
    await selectRange(page, 0, 6);
    await content(page).click({ button: "right", position: { x: 30, y: 10 } });
    await expect(page.getByTestId("editor-context-menu")).toBeVisible();
    await page.getByTestId("context-menu-cut").click();
    await expect(page.getByTestId("editor-context-menu")).toBeHidden();
    await expect(content(page)).toHaveText("world");
    await page.keyboard.press(`${MOD}+End`);
    await page.keyboard.press(`${MOD}+v`);
    await expect(content(page)).toHaveText("worldhello ");
  });

  test("Copy then the menu's Paste round-trips with exact formatting (AC-5, AC-6)", async ({
    page,
  }) => {
    const id = await createDocument(page);
    await content(page).click();
    await page.keyboard.type("styled");
    await selectRange(page, 0, 6);
    await page.keyboard.press(`${MOD}+b`);
    await selectRange(page, 0, 6);
    // Right-click the paragraphs themselves: since SPEC-028 the editable
    // root wraps whole pages, so content-relative coordinates no longer
    // land in the text area.
    await page.locator('[data-section="0"][data-block="0"]').click({ button: "right" });
    await page.getByTestId("context-menu-copy").click();
    await expect(content(page)).toHaveText("styled");

    await page.keyboard.press(`${MOD}+End`);
    await page.keyboard.press("Enter");
    await page.locator('[data-section="0"][data-block="1"]').click({ button: "right" });
    await page.getByTestId("context-menu-paste").click();
    await expect(content(page)).toHaveText("styledstyled");
    await waitForSaved(page);
    const all = blocks(await readLatestModel(page, id));
    // The menu paste recognised the editor's own copy and restored the
    // bold run exactly, despite the async clipboard hiding the private
    // MIME type.
    expect(all[1].runs).toEqual([{ text: "styled", formatting: { bold: true } }]);
  });
});
