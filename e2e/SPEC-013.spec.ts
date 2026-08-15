import { test, expect, type Page } from "@playwright/test";
import { MOD, LINE_START } from "./keys";
import { readFileSync } from "node:fs";
import { unzip } from "../src/lib/docx/zip";

// SPEC-013 — font family, size, text colour, highlight, sub/superscript.
// Selection setup and model reads follow e2e/SPEC-012.spec.ts.

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

interface ModelShape {
  sections: Array<{ blocks: Array<{ runs: ModelRun[] }> }>;
}

/** The newest persisted model (draft beats committed when newer) — same
 * IndexedDB read as e2e/SPEC-012.spec.ts, or null while the document has
 * yet to reach the store. */
async function readPersistedModel(page: Page, id: string): Promise<ModelShape | null> {
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
  }, id) as Promise<ModelShape | null>;
}

/** The same read, waited on: opening a document routes as soon as the
 * import parses, and the model reaches IndexedDB a beat later — a single
 * read that raced that write came back null under load. The save-state
 * indicator is no help here, since an import never flips it away from
 * "All changes saved" in the first place. */
async function readLatestModel(page: Page, id: string): Promise<ModelShape> {
  await expect.poll(async () => (await readPersistedModel(page, id)) !== null).toBe(true);
  const model = await readPersistedModel(page, id);
  if (!model) {
    throw new Error(`no persisted model for ${id}`);
  }
  return model;
}

function firstParagraphRuns(model: ModelShape): ModelRun[] {
  return model.sections[0].blocks[0].runs;
}

/** Selects a substring of the current (single-paragraph) line by character
 * offsets, via Home + arrow keys — same technique as SPEC-011/012. */
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

/** Downloads the document in the given format and returns the file path. */
async function download(page: Page, format: "docx" | "odt"): Promise<string> {
  await page.getByTestId("download-menu").click();
  const file = await Promise.all([
    page.waitForEvent("download"),
    page.getByTestId(`download-${format}`).click(),
  ]).then(([d]) => d);
  const path = await file.path();
  expect(path).toBeTruthy();
  return path!;
}

async function reopen(page: Page, path: string): Promise<string> {
  await page.goto("/");
  await page.getByTestId("open-document-input").setInputFiles(path);
  await expect(page).toHaveURL(/\/doc\/.+/);
  const url = page.url();
  return url.slice(url.lastIndexOf("/") + 1);
}

test("the toolbar offers the font family, size, colour, highlight and sub/superscript controls (AC-1..4, AC-24)", async ({
  page,
}) => {
  await createDocument(page);
  await expect(page.getByTestId("font-family-select")).toBeVisible();
  await expect(page.getByTestId("font-size-input")).toBeVisible();
  await expect(page.getByTestId("font-size-decrease")).toBeVisible();
  await expect(page.getByTestId("font-size-increase")).toBeVisible();
  await expect(page.getByTestId("text-color-button")).toBeVisible();
  await expect(page.getByTestId("highlight-button")).toBeVisible();
  await expect(page.getByTestId("format-subscript")).toBeVisible();
  await expect(page.getByTestId("format-superscript")).toBeVisible();
});

test("the font list groups document fonts first and renders every item in its own typeface (AC-5..8)", async ({
  page,
}) => {
  const id = await createDocument(page);
  await content(page).click();
  await page.keyboard.type("styled text");
  await selectRange(page, 0, 6);
  // Carlito is in the standard list but (in this container) not installed —
  // once applied, the document group must still list it (AC-8).
  await page.getByTestId("font-family-select").selectOption("Carlito");
  await waitForSaved(page);
  expect(firstParagraphRuns(await readLatestModel(page, id))[0]).toEqual({
    text: "styled",
    formatting: { fontFamily: "Carlito" },
  });

  const docGroup = page.getByTestId("font-family-doc-group");
  await expect(docGroup).toHaveAttribute("label", "Fonts in this document");
  const docFonts = await docGroup.locator("option").allTextContents();
  expect(docFonts.map((f) => f.trim())).toEqual(["Calibri", "Carlito"]);

  const allOptions = page.getByTestId("font-family-all-group").locator("option");
  const families = (await allOptions.allTextContents()).map((f) => f.trim());
  for (const family of [
    "Calibri", "Cambria", "Times New Roman", "Arial", "Georgia",
    "Verdana", "Courier New", "Liberation Serif", "Liberation Sans", "Carlito",
  ]) {
    expect(families).toContain(family);
  }
  // Each item is rendered in its own typeface (AC-6).
  const count = await allOptions.count();
  for (let i = 0; i < count; i++) {
    const option = allOptions.nth(i);
    const family = (await option.textContent())!.trim();
    const fontFamily = await option.evaluate((el) => getComputedStyle(el).fontFamily);
    expect(fontFamily).toContain(family);
  }
});

test("the size control accepts 0.5 pt steps, has +/- buttons, and clamps out-of-range input with a tooltip (AC-9..12)", async ({
  page,
}) => {
  const id = await createDocument(page);
  await content(page).click();
  await page.keyboard.type("sized");
  await selectRange(page, 0, 5);

  const input = page.getByTestId("font-size-input");
  await input.fill("11.5");
  await input.press("Enter");
  await waitForSaved(page);
  expect(firstParagraphRuns(await readLatestModel(page, id))).toEqual([
    { text: "sized", formatting: { fontSizePt: 11.5 } },
  ]);

  // +/- step by half a point (AC-10, AC-11).
  await page.getByTestId("font-size-increase").click();
  await expect(input).toHaveValue("12");
  await page.getByTestId("font-size-decrease").click();
  await page.getByTestId("font-size-decrease").click();
  await expect(input).toHaveValue("11");
  await waitForSaved(page);
  expect(firstParagraphRuns(await readLatestModel(page, id))).toEqual([
    { text: "sized", formatting: { fontSizePt: 11 } },
  ]);

  // Out-of-range input clamps with a tooltip (AC-9, AC-12).
  await input.fill("500");
  await input.press("Enter");
  await expect(page.getByTestId("font-size-tooltip")).toBeVisible();
  await expect(page.getByTestId("font-size-tooltip")).toContainText("between 1 and 409");
  // …and genuinely on screen. toBeVisible() ignores clipping by an
  // ancestor's overflow, which is exactly how this hint used to be lost:
  // it hung below a toolbar strip that clips on overflow-y, so the user's
  // input was silently corrected with no explanation. Hit-test it.
  const reached = await page.evaluate(() => {
    const tip = document.querySelector('[data-testid="font-size-tooltip"]')!.getBoundingClientRect();
    const at = document.elementFromPoint(
      Math.round(tip.left + tip.width / 2),
      Math.round(tip.top + tip.height / 2),
    );
    return !!at?.closest('[data-testid="font-size-tooltip"]');
  });
  expect(reached).toBe(true);
  await expect(input).toHaveValue("409");
  await waitForSaved(page);
  expect(firstParagraphRuns(await readLatestModel(page, id))).toEqual([
    { text: "sized", formatting: { fontSizePt: 409 } },
  ]);
});

test("the text colour picker offers a palette, recent colours and a custom input (AC-16..18)", async ({
  page,
}) => {
  const id = await createDocument(page);
  await content(page).click();
  await page.keyboard.type("coloured");
  await selectRange(page, 0, 8);

  await page.getByTestId("text-color-button").click();
  await expect(page.getByTestId("text-color-menu")).toBeVisible();
  await page.locator('[data-testid="text-color-swatch"][data-color="#ff0000"]').click();
  await waitForSaved(page);
  expect(firstParagraphRuns(await readLatestModel(page, id))).toEqual([
    { text: "coloured", formatting: { color: "#ff0000" } },
  ]);
  await expect(content(page).locator("span").first()).toHaveCSS("color", "rgb(255, 0, 0)");

  // The custom hex/RGB input takes any colour (AC-18).
  await selectRange(page, 0, 8);
  await page.getByTestId("text-color-button").click();
  await page.getByTestId("text-color-custom-input").fill("#ff8800");
  await page.getByTestId("text-color-custom-apply").click();
  await waitForSaved(page);
  expect(firstParagraphRuns(await readLatestModel(page, id))).toEqual([
    { text: "coloured", formatting: { color: "#ff8800" } },
  ]);

  // Both applied colours are now in the recent section, newest first (AC-17).
  await page.getByTestId("text-color-button").click();
  const recent = page.getByTestId("text-color-recent").locator("button");
  await expect(recent.nth(0)).toHaveAttribute("data-color", "#ff8800");
  await expect(recent.nth(1)).toHaveAttribute("data-color", "#ff0000");
});

test("the highlight picker offers the 15 Word colours plus 'No colour' (AC-19)", async ({
  page,
}) => {
  const id = await createDocument(page);
  await content(page).click();
  await page.keyboard.type("marked");
  await selectRange(page, 0, 6);

  await page.getByTestId("highlight-button").click();
  await expect(page.locator('[data-testid="highlight-swatch"]')).toHaveCount(15);
  await expect(page.getByTestId("highlight-none")).toBeVisible();
  await page.locator('[data-testid="highlight-swatch"][data-color="yellow"]').click();
  await waitForSaved(page);
  expect(firstParagraphRuns(await readLatestModel(page, id))).toEqual([
    { text: "marked", formatting: { highlight: "yellow" } },
  ]);
  await expect(content(page).locator("span").first()).toHaveCSS(
    "background-color",
    "rgb(255, 255, 0)",
  );

  // "No colour" clears it as an explicit off (SPEC-013 decision).
  await selectRange(page, 0, 6);
  await page.getByTestId("highlight-button").click();
  await page.getByTestId("highlight-none").click();
  await waitForSaved(page);
  expect(firstParagraphRuns(await readLatestModel(page, id))).toEqual([
    { text: "marked", formatting: { highlight: false } },
  ]);
});

test("all four controls reflect the caret and blank out on a mixed selection (AC-22, AC-23)", async ({
  page,
}) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("one two");

  // Style "one": Georgia 18.
  await selectRange(page, 0, 3);
  await page.getByTestId("font-family-select").selectOption("Georgia");
  await selectRange(page, 0, 3);
  const input = page.getByTestId("font-size-input");
  await input.fill("18");
  await input.press("Enter");

  // Caret inside "one" reflects its formatting (AC-22).
  await selectRange(page, 1, 1);
  await expect(page.getByTestId("font-family-select")).toHaveValue("Georgia");
  await expect(input).toHaveValue("18");
  await expect(page.getByTestId("text-color-current")).toHaveAttribute("data-color", "#000000");
  await expect(page.getByTestId("highlight-current")).toHaveAttribute("data-color", "none");

  // Caret in "two" falls back to the document defaults.
  await selectRange(page, 5, 5);
  await expect(page.getByTestId("font-family-select")).toHaveValue("Calibri");
  await expect(input).toHaveValue("11");

  // A selection spanning both blanks the family and size (AC-23).
  await selectRange(page, 0, 7);
  await expect(page.getByTestId("font-family-select")).toHaveValue("");
  await expect(input).toHaveValue("");
});

test("subscript and superscript toggle, are mutually exclusive, and use Ctrl+, / Ctrl+. (AC-24..27)", async ({
  page,
}) => {
  const id = await createDocument(page);
  await content(page).click();
  await page.keyboard.type("H2O");
  await selectRange(page, 1, 2);

  await page.keyboard.press(`${MOD}+Comma`);
  await waitForSaved(page);
  expect(firstParagraphRuns(await readLatestModel(page, id))).toEqual([
    { text: "H", formatting: {} },
    { text: "2", formatting: { vertAlign: "subscript" } },
    { text: "O", formatting: {} },
  ]);
  await expect(page.getByTestId("format-subscript")).toHaveAttribute("aria-pressed", "true");

  // Superscript replaces the subscript — mutually exclusive (AC-25).
  await page.keyboard.press(`${MOD}+Period`);
  await waitForSaved(page);
  expect(firstParagraphRuns(await readLatestModel(page, id))[1]).toEqual({
    text: "2",
    formatting: { vertAlign: "superscript" },
  });
  await expect(page.getByTestId("format-subscript")).toHaveAttribute("aria-pressed", "false");
  await expect(page.getByTestId("format-superscript")).toHaveAttribute("aria-pressed", "true");

  // The toolbar button toggles it back off.
  await page.getByTestId("format-superscript").click();
  await waitForSaved(page);
  expect(firstParagraphRuns(await readLatestModel(page, id))).toEqual([
    { text: "H2O", formatting: {} },
  ]);
});

test("font, size, colour, highlight and superscript round-trip through .docx and .odt (AC-13..15, AC-20, AC-28, AC-29)", async ({
  page,
}) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("styled text");
  await selectRange(page, 0, 11);
  await page.getByTestId("font-family-select").selectOption("Carlito");
  await selectRange(page, 0, 11);
  const input = page.getByTestId("font-size-input");
  await input.fill("11.5");
  await input.press("Enter");
  await selectRange(page, 0, 11);
  await page.getByTestId("text-color-button").click();
  await page.getByTestId("text-color-custom-input").fill("#ff8800");
  await page.getByTestId("text-color-custom-apply").click();
  await selectRange(page, 0, 11);
  await page.getByTestId("highlight-button").click();
  await page.locator('[data-testid="highlight-swatch"][data-color="yellow"]').click();
  await selectRange(page, 7, 11);
  await page.keyboard.press(`${MOD}+Period`);
  await waitForSaved(page);

  const docxPath = await download(page, "docx");
  const odtPath = await download(page, "odt");

  for (const path of [docxPath, odtPath]) {
    const id = await reopen(page, path);
    const runs = firstParagraphRuns(await readLatestModel(page, id));
    expect(runs[0].text).toBe("styled ");
    // Half-points in .docx, points in .odt — either way the exact size
    // comes back with no drift (AC-13..15).
    expect(runs[0].formatting.fontSizePt).toBe(11.5);
    expect(runs[0].formatting.fontFamily).toBe("Carlito");
    expect(runs[0].formatting.color).toBe("#ff8800");
    // The palette colour maps through w:highlight / its ODF hex (AC-20).
    expect(runs[0].formatting.highlight).toBe("yellow");
    expect(runs[0].formatting.vertAlign).toBeUndefined();
    expect(runs[1].text).toBe("text");
    expect(runs[1].formatting.vertAlign).toBe("superscript");

    // The canvas shows it too.
    const first = content(page).locator("span").first();
    await expect(first).toHaveCSS("background-color", "rgb(255, 255, 0)");
    await expect(first).toHaveCSS("color", "rgb(255, 136, 0)");
  }
});

test("a highlight outside the Word palette falls back to shading on .docx save (AC-21)", async ({
  page,
}) => {
  // The toolbar's highlight picker only exposes the 15 Word colours + "No
  // colour" (AC-19), so an out-of-palette highlight only ever arrives via
  // import — e.g. an .odt run with fo:background-color outside that
  // palette. e2e/fixtures/custom-highlight.odt carries exactly that
  // (generated from RunFormatting { highlight: "#ff8800" } via writeOdt).
  const id = await reopen(page, "e2e/fixtures/custom-highlight.odt");
  const runs = firstParagraphRuns(await readLatestModel(page, id));
  expect(runs[0].formatting.highlight).toBe("#ff8800");
  await expect(page.getByTestId("highlight-current")).toHaveAttribute("data-color", "#ff8800");

  const docxPath = await download(page, "docx");
  const parts = await unzip(readFileSync(docxPath));
  const documentXml = new TextDecoder().decode(parts.get("word/document.xml"));
  expect(documentXml).toContain('<w:shd w:val="clear" w:color="auto" w:fill="FF8800"/>');
  expect(documentXml).not.toContain("<w:highlight");
});
