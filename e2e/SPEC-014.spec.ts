import { test, expect, type Page } from "@playwright/test";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import { MOD, LINE_START } from "./keys";
import { equalStructural } from "../src/lib/equality";

// SPEC-014 — paragraph alignment, indents, spacing, RTL and the ruler.
// Selection setup, model reads and download/reopen follow e2e/SPEC-013.spec.ts.

const odfFixturesDir = join(dirname(fileURLToPath(import.meta.url)), "..", "src", "lib", "odf", "fixtures");

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
const paragraphAt = (page: Page, index: number) =>
  content(page).locator(".doc-paragraph").nth(index);

interface ModelBlock {
  runs: Array<{ text: string; formatting: Record<string, unknown> }>;
  formatting?: Record<string, unknown>;
}

interface ModelShape {
  sections: Array<{ blocks: ModelBlock[] }>;
  defaultStyles: { margins: { leftCm: number; rightCm: number } };
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

function blockFormatting(model: ModelShape, index: number): Record<string, unknown> {
  return model.sections[0].blocks[index].formatting ?? {};
}

/** Extends the selection from the caret's paragraph start down `lines`
 * paragraphs — the multi-paragraph "touched by the selection" setup. */
async function selectDown(page: Page, lines: number): Promise<void> {
  await page.keyboard.press(LINE_START);
  await page.keyboard.down("Shift");
  for (let i = 0; i < lines; i++) {
    await page.keyboard.press("ArrowDown");
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

test("the toolbar has alignment, indent and line-spacing controls, and a ruler sits above the page (AC-1, AC-11, AC-22)", async ({
  page,
}) => {
  await createDocument(page);
  for (const id of [
    "align-left", "align-center", "align-right", "align-justify",
    "indent-increase", "indent-decrease", "line-spacing-button", "paragraph-direction",
  ]) {
    await expect(page.getByTestId(id)).toBeVisible();
  }
  await page.getByTestId("line-spacing-button").click();
  const menu = page.getByTestId("line-spacing-menu");
  await expect(menu.getByTestId("line-spacing-single")).toHaveText("Single");
  await expect(menu.getByTestId("line-spacing-115")).toHaveText("1.15");
  await expect(menu.getByTestId("line-spacing-15")).toHaveText("1.5");
  await expect(menu.getByTestId("line-spacing-double")).toHaveText("Double");
  await expect(menu.getByTestId("line-spacing-custom")).toHaveText("Custom spacing…");
  await expect(menu.getByTestId("space-before-toggle")).toHaveText("Add space before paragraph");
  await page.keyboard.press("Escape");

  await expect(page.getByTestId("document-ruler")).toBeVisible();
  for (const marker of ["start", "first-line", "end", "margin-left", "margin-right"]) {
    await expect(page.getByTestId(`ruler-marker-${marker}`)).toBeVisible();
  }
});

test("alignment applies to every paragraph touched by the selection (AC-2, AC-4, AC-5)", async ({
  page,
}) => {
  const id = await createDocument(page);
  await content(page).click();
  await page.keyboard.type("first");
  await page.keyboard.press("Enter");
  await page.keyboard.type("second");
  await page.keyboard.press("Enter");
  await page.keyboard.type("third");

  // Selection from "second" into "third": both touched, "first" is not.
  await page.keyboard.press("ArrowUp");
  await selectDown(page, 1);
  await page.getByTestId("align-center").click();
  await waitForSaved(page);

  const model = await readLatestModel(page, id);
  expect(blockFormatting(model, 0).alignment).toBeUndefined();
  expect(blockFormatting(model, 1).alignment).toBe("center");
  expect(blockFormatting(model, 2).alignment).toBe("center");
  await expect(paragraphAt(page, 1)).toHaveCSS("text-align", "center");
  await expect(page.getByTestId("align-center")).toHaveAttribute("aria-pressed", "true");
});

test("Ctrl+L, Ctrl+E, Ctrl+R and Ctrl+J set the alignment (AC-3)", async ({ page }) => {
  const id = await createDocument(page);
  await content(page).click();
  await page.keyboard.type("aligned by shortcut");

  for (const [combo, alignment] of [
    [`${MOD}+KeyE`, "center"],
    [`${MOD}+KeyR`, "right"],
    [`${MOD}+KeyJ`, "justify"],
    [`${MOD}+KeyL`, "left"],
  ] as const) {
    await page.keyboard.press(combo);
    await waitForSaved(page);
    expect(blockFormatting(await readLatestModel(page, id), 0).alignment).toBe(alignment);
  }
});

test("increase indent steps by 1.27 cm and decrease stops at 0 (AC-6, AC-7)", async ({
  page,
}) => {
  const id = await createDocument(page);
  await content(page).click();
  await page.keyboard.type("indented");

  await page.getByTestId("indent-increase").click();
  await page.getByTestId("indent-increase").click();
  await waitForSaved(page);
  expect(blockFormatting(await readLatestModel(page, id), 0).indentStartCm as number).toBeCloseTo(2.54, 5);
  // 2.54 cm at CSS's 96 dpi is exactly 96px.
  await expect(paragraphAt(page, 0)).toHaveCSS("margin-inline-start", "96px");

  await page.getByTestId("indent-decrease").click();
  await page.getByTestId("indent-decrease").click();
  await page.getByTestId("indent-decrease").click();
  await waitForSaved(page);
  expect(blockFormatting(await readLatestModel(page, id), 0).indentStartCm ?? 0).toBe(0);
});

test("Tab at a paragraph's offset 0 indents; Shift+Tab un-indents; Tab elsewhere types a tab (AC-8, AC-9, AC-10)", async ({
  page,
}) => {
  const id = await createDocument(page);
  await content(page).click();
  await page.keyboard.type("tabbed");

  // Collapsed caret at offset 0 — the start of the first line.
  await page.keyboard.press(LINE_START);
  await page.keyboard.press("Tab");
  await waitForSaved(page);
  let model = await readLatestModel(page, id);
  expect(blockFormatting(model, 0).indentStartCm as number).toBeCloseTo(1.27, 5);
  expect(model.sections[0].blocks[0].runs[0].text).toBe("tabbed");

  await page.keyboard.press("Shift+Tab");
  await waitForSaved(page);
  model = await readLatestModel(page, id);
  expect(blockFormatting(model, 0).indentStartCm ?? 0).toBe(0);

  // Mid-paragraph: a literal tab character, no indent change.
  await page.keyboard.press(LINE_START);
  for (let i = 0; i < 3; i++) {
    await page.keyboard.press("ArrowRight");
  }
  await page.keyboard.press("Tab");
  await waitForSaved(page);
  model = await readLatestModel(page, id);
  expect(model.sections[0].blocks[0].runs.map((run) => run.text).join("")).toBe("tab\tbed");
  expect(blockFormatting(model, 0).indentStartCm ?? 0).toBe(0);

  // A non-collapsed selection also gets the literal tab (replacing it).
  await page.keyboard.press(LINE_START);
  await page.keyboard.down("Shift");
  for (let i = 0; i < 4; i++) {
    await page.keyboard.press("ArrowRight");
  }
  await page.keyboard.up("Shift");
  await page.keyboard.press("Tab");
  await waitForSaved(page);
  model = await readLatestModel(page, id);
  expect(model.sections[0].blocks[0].runs.map((run) => run.text).join("")).toBe("\tbed");
  expect(blockFormatting(model, 0).indentStartCm ?? 0).toBe(0);
});

test("the custom spacing dialog applies to every touched paragraph; Cancel discards (AC-12, AC-13, AC-14, AC-5)", async ({
  page,
}) => {
  const id = await createDocument(page);
  await content(page).click();
  await page.keyboard.type("one");
  await page.keyboard.press("Enter");
  await page.keyboard.type("two");

  // A preset first: 1.5 lines on the second paragraph only.
  await page.getByTestId("line-spacing-button").click();
  await page.getByTestId("line-spacing-15").click();
  await waitForSaved(page);
  expect(blockFormatting(await readLatestModel(page, id), 1).lineSpacing).toBe(1.5);

  // Custom spacing over both paragraphs: exactly 14 pt, 12 pt before, 6 after.
  await page.keyboard.press("ArrowUp");
  await selectDown(page, 1);
  await page.getByTestId("line-spacing-button").click();
  await page.getByTestId("line-spacing-custom").click();
  const dialog = page.getByTestId("custom-spacing-dialog");
  await expect(dialog).toBeVisible();
  await dialog.getByTestId("spacing-mode").selectOption("exact");
  await dialog.getByTestId("spacing-value").fill("14");
  await dialog.getByTestId("spacing-before").fill("12");
  await dialog.getByTestId("spacing-after").fill("6");
  await dialog.getByTestId("spacing-apply").click();
  await waitForSaved(page);
  let model = await readLatestModel(page, id);
  for (const index of [0, 1]) {
    expect(blockFormatting(model, index)).toMatchObject({
      lineSpacingPt: 14,
      lineSpacingRule: "exact",
      spaceBeforePt: 12,
      spaceAfterPt: 6,
    });
    expect(blockFormatting(model, index).lineSpacing).toBeUndefined();
  }

  // Cancel discards edited values (AC-14).
  await page.getByTestId("line-spacing-button").click();
  await page.getByTestId("line-spacing-custom").click();
  await dialog.getByTestId("spacing-value").fill("99");
  await dialog.getByTestId("spacing-cancel").click();
  await expect(dialog).not.toBeVisible();
  await waitForSaved(page);
  model = await readLatestModel(page, id);
  expect(blockFormatting(model, 0).lineSpacingPt).toBe(14);
});

test("the space-before toggle follows the 0/absent -> 12 pt table, statelessly, per paragraph (AC-15..20)", async ({
  page,
}) => {
  const id = await createDocument(page);
  await content(page).click();
  await page.keyboard.type("first");
  await page.keyboard.press("Enter");
  await page.keyboard.type("second");

  // Give "first" a custom 6 pt space-before via the dialog.
  await page.keyboard.press("ArrowUp");
  await page.getByTestId("line-spacing-button").click();
  await page.getByTestId("line-spacing-custom").click();
  await page.getByTestId("spacing-before").fill("6");
  await page.getByTestId("spacing-apply").click();
  await waitForSaved(page);
  expect(blockFormatting(await readLatestModel(page, id), 0).spaceBeforePt).toBe(6);

  // Select both paragraphs. The label reflects only the FIRST touched
  // paragraph (6 pt -> "Remove", AC-16, AC-17)...
  await selectDown(page, 1);
  await page.getByTestId("line-spacing-button").click();
  const toggle = page.getByTestId("space-before-toggle");
  await expect(toggle).toHaveText("Remove space before paragraph");
  await toggle.click();
  await waitForSaved(page);
  // ...but the effect is per paragraph, strictly by the table (AC-18):
  // 6 pt -> explicit 0; absent -> explicit 12.
  let model = await readLatestModel(page, id);
  expect(blockFormatting(model, 0).spaceBeforePt).toBe(0);
  expect(blockFormatting(model, 1).spaceBeforePt).toBe(12);

  // Stateless (AC-19, AC-20): adding back to "first" yields 12, never 6.
  await paragraphAt(page, 0).click();
  await page.getByTestId("line-spacing-button").click();
  await expect(toggle).toHaveText("Add space before paragraph");
  await toggle.click();
  await waitForSaved(page);
  model = await readLatestModel(page, id);
  expect(blockFormatting(model, 0).spaceBeforePt).toBe(12);
});

test("an explicitly overridden style space-before does not reappear after save and reload (AC-21)", async ({
  page,
}) => {
  // The LibreOffice fixture's centred paragraph carries a 12 pt style-driven
  // space-before; removing it writes an explicit 0 that must survive an
  // .odt save + reopen.
  const id = await reopen(page, join(odfFixturesDir, "libreoffice-writer.odt"));
  // The fixture stores the 12 pt as an ODF cm length, so it comes back a
  // hair off exact (11.9906 pt) — well inside the 0.01 cm tolerance.
  expect(blockFormatting(await readLatestModel(page, id), 6).spaceBeforePt as number).toBeCloseTo(12, 1);

  await paragraphAt(page, 6).click();
  await page.getByTestId("line-spacing-button").click();
  const toggle = page.getByTestId("space-before-toggle");
  await expect(toggle).toHaveText("Remove space before paragraph");
  await toggle.click();
  await waitForSaved(page);
  expect(blockFormatting(await readLatestModel(page, id), 6).spaceBeforePt).toBe(0);

  const reopenedId = await reopen(page, await download(page, "odt"));
  expect(blockFormatting(await readLatestModel(page, reopenedId), 6).spaceBeforePt).toBe(0);
});

test("dragging a ruler marker changes the value live (AC-22, AC-23)", async ({ page }) => {
  const id = await createDocument(page);
  await content(page).click();
  await page.keyboard.type("ruled");

  const ruler = page.getByTestId("document-ruler");
  // Typing scrolls the caret into view, which can push the ruler above the
  // scroll container's fold — bring it back before measuring coordinates.
  await ruler.scrollIntoViewIfNeeded();
  const rulerBox = (await ruler.boundingBox())!;
  const pxPerCm = rulerBox.width / 21;

  const start = page.getByTestId("ruler-marker-start");
  const startBox = (await start.boundingBox())!;
  const grabX = startBox.x + startBox.width / 2;
  const grabY = startBox.y + startBox.height / 2;

  await page.mouse.move(grabX, grabY);
  await page.mouse.down();
  await page.mouse.move(grabX + 2 * pxPerCm, grabY, { steps: 5 });
  // Still mid-drag: the paragraph indent and the marker have already moved.
  const liveMargin = await paragraphAt(page, 0).evaluate(
    (el) => parseFloat(getComputedStyle(el).marginInlineStart),
  );
  expect(liveMargin).toBeGreaterThan(pxPerCm);
  await page.mouse.up();
  await waitForSaved(page);

  const model = await readLatestModel(page, id);
  expect(model.sections[0].blocks[0].formatting?.indentStartCm as number).toBeCloseTo(2, 1);

  // The first-line marker drags independently (AC-22's fourth marker is
  // exercised through the model value it controls).
  const firstLine = page.getByTestId("ruler-marker-first-line");
  const firstBox = (await firstLine.boundingBox())!;
  await page.mouse.move(firstBox.x + firstBox.width / 2, firstBox.y + firstBox.height / 2);
  await page.mouse.down();
  await page.mouse.move(firstBox.x + firstBox.width / 2 + pxPerCm, firstBox.y, { steps: 5 });
  await page.mouse.up();
  await waitForSaved(page);
  const updated = await readLatestModel(page, id);
  expect(updated.sections[0].blocks[0].formatting?.indentFirstLineCm as number).toBeCloseTo(1, 1);
});

test("right-to-left direction is settable and paints the logical start indent on the right edge (AC-24, AC-25)", async ({
  page,
}) => {
  const id = await createDocument(page);
  await content(page).click();
  await page.keyboard.type("نص تجريبي");
  await page.getByTestId("indent-increase").click();
  await page.getByTestId("paragraph-direction").click();
  await waitForSaved(page);

  const model = await readLatestModel(page, id);
  expect(blockFormatting(model, 0).direction).toBe("rtl");
  expect(blockFormatting(model, 0).indentStartCm as number).toBeCloseTo(1.27, 5);

  const paragraph = paragraphAt(page, 0);
  await expect(paragraph).toHaveCSS("direction", "rtl");
  // The logical start indent (1.27 cm = 48px) paints on the physical RIGHT.
  await expect(paragraph).toHaveCSS("margin-right", "48px");
  await expect(paragraph).toHaveCSS("margin-left", "0px");
  await expect(page.getByTestId("paragraph-direction")).toHaveAttribute("aria-pressed", "true");

  // The ruler mirrors: the start marker now sits right of the end marker.
  const startBox = (await page.getByTestId("ruler-marker-start").boundingBox())!;
  const endBox = (await page.getByTestId("ruler-marker-end").boundingBox())!;
  expect(startBox.x).toBeGreaterThan(endBox.x);

  // Toggling back clears the direction rather than storing "ltr".
  await page.getByTestId("paragraph-direction").click();
  await waitForSaved(page);
  expect(blockFormatting(await readLatestModel(page, id), 0).direction).toBeUndefined();
});

test.describe("same-format round trips (AC-26)", () => {
  for (const format of ["docx", "odt"] as const) {
    test(`every SPEC-014 property survives a .${format} save + reopen with an equalStructural model`, async ({
      page,
    }) => {
      const id = await createDocument(page);
      await content(page).click();
      await page.keyboard.type("centred and indented");
      await page.keyboard.press("Enter");
      await page.keyboard.type("نص من اليمين");

      // First paragraph: centre, one indent step, custom exact spacing.
      await paragraphAt(page, 0).click();
      await page.getByTestId("align-center").click();
      await page.getByTestId("indent-increase").click();
      await page.getByTestId("line-spacing-button").click();
      await page.getByTestId("line-spacing-custom").click();
      await page.getByTestId("spacing-mode").selectOption("exact");
      await page.getByTestId("spacing-value").fill("14");
      await page.getByTestId("spacing-before").fill("12");
      await page.getByTestId("spacing-after").fill("6");
      await page.getByTestId("spacing-apply").click();

      // Second paragraph: RTL with a start indent and 1.15 line spacing.
      await paragraphAt(page, 1).click();
      await page.getByTestId("paragraph-direction").click();
      await page.getByTestId("indent-increase").click();
      await page.getByTestId("line-spacing-button").click();
      await page.getByTestId("line-spacing-115").click();
      await waitForSaved(page);
      const edited = await readLatestModel(page, id);

      // Save, reopen (the "opened" model), save that, reopen again (the
      // "reopened" model): the two must be structurally equal.
      const openedId = await reopen(page, await download(page, format));
      const opened = await readLatestModel(page, openedId);
      const reopenedId = await reopen(page, await download(page, format));
      const reopened = await readLatestModel(page, reopenedId);
      expect(equalStructural(opened.sections, reopened.sections)).toBe(true);

      // No unit drift beyond 0.01 cm between the edit and what came back.
      const CM_PER_PT = 2.54 / 72;
      for (const [block, key] of [
        [0, "indentStartCm"], [1, "indentStartCm"],
      ] as const) {
        const before = blockFormatting(edited, block)[key] as number;
        const after = blockFormatting(opened, block)[key] as number;
        expect(Math.abs(after - before)).toBeLessThan(0.01);
      }
      for (const key of ["lineSpacingPt", "spaceBeforePt", "spaceAfterPt"] as const) {
        const before = blockFormatting(edited, 0)[key] as number;
        const after = blockFormatting(opened, 0)[key] as number;
        expect(Math.abs(after - before) * CM_PER_PT).toBeLessThan(0.01);
      }
      expect(blockFormatting(opened, 0).alignment).toBe("center");
      expect(blockFormatting(opened, 0).lineSpacingRule).toBe("exact");
      expect(blockFormatting(opened, 1).direction).toBe("rtl");
      expect(blockFormatting(opened, 1).lineSpacing).toBe(1.15);
    });
  }
});
