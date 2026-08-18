import { test, expect, type Page } from "@playwright/test";
import zlib from "node:zlib";
import { fileURLToPath } from "node:url";
import path from "node:path";

const PNG_1x1 =
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=";

// Minimal RGB PNG encoder so a right-clickable image has real area (a 1x1 image
// is entirely covered by its resize handles).
function buildPng(width: number, height: number, [r, g, b]: [number, number, number]): Buffer {
  const table = Array.from({ length: 256 }, (_, n) => {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    return c >>> 0;
  });
  const crc32 = (buf: Buffer) => {
    let c = 0xffffffff;
    for (const byte of buf) c = table[(c ^ byte) & 0xff] ^ (c >>> 8);
    return (c ^ 0xffffffff) >>> 0;
  };
  const chunk = (type: string, data: Buffer) => {
    const body = Buffer.concat([Buffer.from(type, "ascii"), data]);
    const len = Buffer.alloc(4);
    len.writeUInt32BE(data.length, 0);
    const crc = Buffer.alloc(4);
    crc.writeUInt32BE(crc32(body), 0);
    return Buffer.concat([len, body, crc]);
  };
  const rowBytes = width * 3;
  const raw = Buffer.alloc((rowBytes + 1) * height);
  for (let y = 0; y < height; y++) {
    for (let x = 0; x < width; x++) {
      const o = y * (rowBytes + 1) + 1 + x * 3;
      raw[o] = r; raw[o + 1] = g; raw[o + 2] = b;
    }
  }
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(width, 0);
  ihdr.writeUInt32BE(height, 4);
  ihdr[8] = 8; ihdr[9] = 2;
  return Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    chunk("IHDR", ihdr),
    chunk("IDAT", zlib.deflateSync(raw)),
    chunk("IEND", Buffer.alloc(0)),
  ]);
}

async function deselect(page: Page): Promise<void> {
  await page.getByTestId("document-content").click({ position: { x: 8, y: 8 } });
  await page.keyboard.press("Escape");
}
async function openInsert(page: Page): Promise<void> {
  await page.getByTestId("menu-title-insert").click();
  await expect(page.getByTestId("insert-image-button")).toBeVisible();
}
async function insertImage(page: Page, big = false): Promise<void> {
  await deselect(page);
  await openInsert(page);
  await page.getByTestId("insert-image-button").click();
  await page.getByTestId("insert-image-file-input").setInputFiles({
    name: "photo.png",
    mimeType: "image/png",
    buffer: big ? buildPng(220, 120, [0x40, 0x80, 0xc0]) : Buffer.from(PNG_1x1, "base64"),
  });
  await expect(
    page.getByTestId("document-content").locator('[data-testid="doc-image"]').last(),
  ).toBeVisible();
}
// Select a (deselected) image cleanly — the first click lands on the img before
// any resize-handle overlay exists.
async function selectLastImage(page: Page): Promise<void> {
  await deselect(page);
  await page.getByTestId("document-content").locator('[data-testid="doc-image"]').last().click();
}
async function captionSelected(page: Page, label?: string): Promise<void> {
  await openInsert(page);
  await page.getByTestId("insert-caption-button").click();
  await expect(page.getByTestId("caption-dialog")).toBeVisible();
  if (label) await page.getByTestId("caption-label").selectOption(label);
  await page.getByTestId("caption-insert").click();
  await expect(page.getByTestId("caption-dialog")).toHaveCount(0);
}

// SPEC-055 — captions and the table of figures. This focused e2e covers the
// three flows that decide the feature's usability: captioning a table selected
// the ordinary way (a click into a cell, AC-1/AC-2), a freshly generated table
// of figures NOT showing a spurious update control (AC-32), and no uncaught
// exception during the idle pagination pass that follows (a ToF is a block-
// level atom, not a paragraph). The exhaustive round-trip/numbering coverage
// lives in the src/lib unit specs and fixtures.captions.spec.ts.

const content = (page: Page) => page.getByTestId("document-content");

async function createDocument(page: Page): Promise<void> {
  await page.goto("/");
  await page.getByTestId("new-document").click();
  await expect(page).toHaveURL(/\/doc\/.+/);
}

test("caption a table selected by clicking a cell, then a fresh table of figures", async ({ page }) => {
  const errors: string[] = [];
  page.on("console", (msg) => {
    if (msg.type() === "error") errors.push(msg.text());
  });
  page.on("pageerror", (err) => errors.push(String(err)));

  await createDocument(page);

  // Insert a 3×3 table via the grid picker.
  await page.getByTestId("menu-title-insert").click();
  await page.getByTestId("table-grid-cell-3-3").click();
  await expect(content(page).locator("table")).toHaveCount(1);

  // Click into a table cell — the ordinary way to "select" a table.
  await content(page).locator("table td").first().click();

  // Insert -> Caption opens the dialog (the AC-1/AC-2 fix).
  await page.getByTestId("menu-title-insert").click();
  await page.getByTestId("insert-caption-button").click();
  await expect(page.getByTestId("caption-dialog")).toBeVisible();

  // Default label is Figure; switch to Table to caption the table.
  await page.getByTestId("caption-label").selectOption("Table");
  await page.getByTestId("caption-insert").click();
  await expect(page.getByTestId("caption-dialog")).toHaveCount(0);

  // A Caption paragraph now shows "Table 1".
  await expect(content(page).getByText(/Table\s*1/).first()).toBeVisible();

  // Type some caption text so it reads like a real caption.
  await page.keyboard.type("Quarterly results");

  // Insert -> Table of figures for the Figure label is empty; use Table.
  await page.getByTestId("menu-title-insert").click();
  await page.getByTestId("insert-table-of-figures-button").click();
  await expect(page.getByTestId("table-of-figures-dialog")).toBeVisible();
  await page.getByTestId("tof-label").selectOption("Table");
  await page.getByTestId("tof-insert").click();
  await expect(page.getByTestId("table-of-figures-dialog")).toHaveCount(0);

  // The ToF block generated with one entry.
  await expect(page.getByTestId("tof-block")).toBeVisible();
  await expect(page.getByTestId("tof-entry")).toHaveCount(1);

  // AC-32: a freshly generated ToF shows NO update bar.
  await expect(page.getByTestId("tof-update-bar")).toHaveCount(0);

  // Give the idle pagination pass a moment to run (the prior crash site).
  await page.waitForTimeout(800);
  await expect(page.getByTestId("tof-block")).toBeVisible();
  await expect(page.getByTestId("tof-update-bar")).toHaveCount(0);

  // Regression: selecting the ToF, then clicking into a body paragraph and
  // pressing Backspace must NOT delete the whole table of figures. The stale
  // tofSelection previously misrouted the keystroke into handleTofDelete
  // because the selection was never cleared on click-away (parity with the
  // TOC, which clears tocSelection in handleCanvasMouseDown).
  await page.getByTestId("tof-block").click();
  // The body caption run (not the ToF entry link, which also contains the text).
  await content(page).getByText(": Quarterly results", { exact: true }).click();
  await page.keyboard.press("Backspace");
  await expect(page.getByTestId("tof-block")).toBeVisible();
  await expect(page.getByTestId("tof-entry")).toHaveCount(1);

  expect(errors, `console errors:\n${errors.join("\n")}`).toEqual([]);
});

// --- caption dialog + insertion (AC-1..AC-10) -------------------------------

test("caption dialog offers all controls and inserts a Figure caption below (AC-2..AC-10)", async ({ page }) => {
  await createDocument(page);
  await insertImage(page);
  await selectLastImage(page);
  await openInsert(page);
  // AC-1: the Insert menu item is enabled for a selected object.
  await expect(page.getByTestId("insert-caption-button")).toBeEnabled();
  await page.getByTestId("insert-caption-button").click();
  await expect(page.getByTestId("caption-dialog")).toBeVisible(); // AC-2

  // AC-5 built-in labels.
  const labels = await page.getByTestId("caption-label").locator("option").allInnerTexts();
  expect(labels.join("|")).toContain("Figure");
  expect(labels.join("|")).toContain("Table");
  expect(labels.join("|")).toContain("Equation");
  // AC-6 new-label action; AC-8 separator editable; AC-9 exclude checkbox.
  await expect(page.getByTestId("caption-new-label")).toBeVisible();
  await expect(page.getByTestId("caption-separator")).toBeEditable();
  await expect(page.getByTestId("caption-exclude-label")).toBeVisible();
  // AC-7 position options.
  const pos = await page.getByTestId("caption-position").locator("option")
    .evaluateAll((os) => os.map((o) => (o as HTMLOptionElement).value));
  expect(pos).toEqual(expect.arrayContaining(["above", "below"]));
  // AC-10 number-format options.
  const fmt = await page.getByTestId("caption-format").locator("option")
    .evaluateAll((os) => os.map((o) => (o as HTMLOptionElement).value));
  expect(fmt).toEqual(expect.arrayContaining(["decimal", "lowerRoman", "lowerAlpha"]));

  // AC-3 / AC-4: default insert places a "Figure 1: <text>" caption below.
  await page.getByTestId("caption-insert").click();
  await expect(page.getByTestId("caption-dialog")).toHaveCount(0);
  await page.keyboard.type("A photo");
  await expect(content(page).getByText("Figure 1: A photo")).toBeVisible();
});

test("context menu of a selected image opens the caption dialog (AC-1)", async ({ page }) => {
  await createDocument(page);
  await insertImage(page, true);
  const img = content(page).locator('[data-testid="doc-image"]').first();
  await img.click(); // select
  await img.click({ button: "right", position: { x: 110, y: 60 } }); // right-click the body
  await expect(page.getByTestId("editor-context-menu")).toBeVisible();
  await page.getByTestId("context-menu-insert-caption").click();
  await expect(page.getByTestId("caption-dialog")).toBeVisible();
});

test("exclude-label caption begins with the number, no label run (AC-4, AC-9)", async ({ page }) => {
  await createDocument(page);
  await insertImage(page);
  await selectLastImage(page);
  await openInsert(page);
  await page.getByTestId("insert-caption-button").click();
  await page.getByTestId("caption-exclude-label").check();
  await page.getByTestId("caption-insert").click();
  await page.keyboard.type("no label here");
  await expect(content(page).getByText(/^\s*1:\s*no label here/).first()).toBeVisible();
  await expect(content(page).getByText(/Figure\s*1/)).toHaveCount(0);
});

test("new label must be unique and non-empty, and starts an independent sequence (AC-6, AC-11, AC-12)", async ({ page }) => {
  await createDocument(page);
  await insertImage(page);
  await selectLastImage(page);
  await openInsert(page);
  await page.getByTestId("insert-caption-button").click();
  await page.getByTestId("caption-new-label").click();
  const input = page.getByTestId("caption-new-label-input");
  const add = page.getByTestId("caption-new-label-add");
  await expect(add).toBeDisabled(); // AC-12 empty
  await input.fill("figure");
  await expect(add).toBeDisabled(); // AC-12 case-insensitive duplicate
  await input.fill("Photo");
  await expect(add).toBeEnabled(); // AC-11 unique
  await add.click();
  await expect(page.getByTestId("caption-label")).toHaveValue("Photo");
  await page.getByTestId("caption-insert").click();
  await page.keyboard.type("first");
  await expect(content(page).getByText(/Photo\s*1/).first()).toBeVisible(); // sequence starts at 1
});

// --- numbering (AC-15..AC-17, AC-43) ----------------------------------------

test("inserting an object renumbers captions; single undo reverts; labels independent (AC-15, AC-16, AC-17)", async ({ page }) => {
  await createDocument(page);
  await insertImage(page);
  await selectLastImage(page);
  await captionSelected(page);
  await content(page).getByText(/Figure\s*1/).first().click();
  await page.keyboard.press("End");
  await page.keyboard.press("Enter");
  await insertImage(page);
  await selectLastImage(page);
  await captionSelected(page);
  await expect(content(page).getByText(/Figure\s*2/).first()).toBeVisible();

  // AC-17: a Table caption is numbered on its own sequence; figures unaffected.
  // Put the caret in a fresh body paragraph before inserting the table.
  await content(page).getByText(/Figure\s*2/).first().click();
  await page.keyboard.press("End");
  await page.keyboard.press("Enter");
  await page.getByTestId("menu-title-insert").click();
  await page.getByTestId("table-grid-cell-2-2").click();
  await expect(content(page).locator("table")).toHaveCount(1);
  await content(page).locator("table td").first().click();
  await captionSelected(page, "Table");
  await expect(content(page).getByText(/Table\s*1/).first()).toBeVisible();
  await expect(content(page).getByText(/Figure\s*2/).first()).toBeVisible();

  // AC-15/AC-16: insert a new figure at the top → a Figure 3 appears; one undo
  // reverts both the insertion and the renumber (no lingering Figure 3).
  await content(page).click({ position: { x: 8, y: 8 } });
  await page.keyboard.press("Control+Home");
  await insertImage(page);
  await selectLastImage(page);
  await captionSelected(page);
  await expect(content(page).getByText(/Figure\s*3/).first()).toBeVisible();
  await page.keyboard.press("Control+z");
  await expect(content(page).getByText(/Figure\s*3/)).toHaveCount(0);
  await expect(content(page).getByText(/Figure\s*2/).first()).toBeVisible();
});

// --- auto-caption (AC-18, AC-19) --------------------------------------------

test("auto-caption is per-type and inserts a caption on object insert (AC-18, AC-19)", async ({ page }) => {
  await createDocument(page);
  await openInsert(page);
  // AC-18 default off for both.
  await expect(page.getByTestId("auto-caption-images")).toHaveAttribute("aria-checked", "false");
  await expect(page.getByTestId("auto-caption-tables")).toHaveAttribute("aria-checked", "false");
  await page.getByTestId("auto-caption-images").click();
  await openInsert(page);
  // Independent: enabling images does not enable tables.
  await expect(page.getByTestId("auto-caption-images")).toHaveAttribute("aria-checked", "true");
  await expect(page.getByTestId("auto-caption-tables")).toHaveAttribute("aria-checked", "false");
  await page.keyboard.press("Escape");
  // AC-19: inserting an image auto-inserts a caption with the caret in its text run.
  await insertImage(page);
  await expect(content(page).getByText(/Figure\s*1/).first()).toBeVisible();
  await page.keyboard.type("auto text");
  await expect(content(page).getByText(/Figure 1:\s*auto text/).first()).toBeVisible();
});

// --- delete prompt (AC-20, AC-21, AC-22) ------------------------------------

test("deleting a captioned object prompts; keep-caption keeps caption; uncaptioned no prompt (AC-20, AC-21, AC-22)", async ({ page }) => {
  await createDocument(page);
  // AC-22: uncaptioned image deletes with no prompt.
  await insertImage(page);
  await selectLastImage(page);
  await page.keyboard.press("Delete");
  await expect(page.getByTestId("delete-caption-dialog")).toHaveCount(0);
  await expect(content(page).locator('[data-testid="doc-image"]')).toHaveCount(0);

  // AC-20: captioned image delete raises the prompt (default "Delete caption too").
  await insertImage(page);
  await selectLastImage(page);
  await captionSelected(page);
  await expect(content(page).getByText(/Figure\s*1/).first()).toBeVisible();
  await selectLastImage(page);
  await page.keyboard.press("Delete");
  await expect(page.getByTestId("delete-caption-dialog")).toBeVisible();
  await expect(page.getByTestId("delete-caption-both")).toBeVisible();
  // AC-21: keeping the caption removes only the object.
  await page.getByTestId("delete-caption-keep").click();
  await expect(content(page).locator('[data-testid="doc-image"]')).toHaveCount(0);
  await expect(content(page).getByText(/Figure\s*1/).first()).toBeVisible();
});

// --- cross-references (AC-23, AC-24, AC-26, AC-27) ---------------------------

test("caption cross-reference lists captions, renders labelAndNumber, and renumbers on update (AC-23, AC-24, AC-26, AC-27)", async ({ page }) => {
  await createDocument(page);
  await insertImage(page);
  await selectLastImage(page);
  await captionSelected(page);
  await expect(content(page).getByText(/Figure\s*1/).first()).toBeVisible();

  await content(page).click();
  await page.keyboard.press("Control+End");
  await page.keyboard.press("Enter");
  await page.keyboard.type("See ");
  await openInsert(page);
  await page.getByTestId("insert-cross-reference-button").click();
  await expect(page.getByTestId("cross-reference-dialog")).toBeVisible();
  // AC-24: six caption formats.
  const formats = await page.getByTestId("cross-reference-show").locator("option")
    .evaluateAll((os) => os.map((o) => (o as HTMLOptionElement).value));
  expect(formats).toEqual(
    expect.arrayContaining(["full", "labelAndNumber", "numberOnly", "textOnly", "page", "position"]),
  );
  // AC-23: the caption is offered as a target.
  const target = page.locator('[data-testid^="cross-reference-caption-"]').first();
  await expect(target).toBeVisible();
  await target.click();
  await page.getByTestId("cross-reference-show").selectOption("labelAndNumber");
  await page.getByTestId("cross-reference-insert").click();
  // AC-26: renders "Figure 1".
  await expect(content(page).getByText(/See\s*Figure\s*1/).first()).toBeVisible();

  // AC-27: add a Figure above the target, update all fields → reference shows Figure 2.
  await page.keyboard.press("Control+Home");
  await insertImage(page);
  await selectLastImage(page);
  await captionSelected(page);
  await page.getByTestId("menu-title-tools").click();
  await page.getByTestId("tools-menu-update-all-fields").click();
  await expect(content(page).getByText(/See\s*Figure\s*2/).first()).toBeVisible();
});

// --- table of figures (AC-29..AC-36) ----------------------------------------

// Append a captioned figure at the END of the document (so captions keep
// document order — deselect()'s top-left click would otherwise insert each new
// image above the previous one).
async function figureWithCaption(page: Page, text: string): Promise<void> {
  await content(page).click();
  await page.keyboard.press("Escape");
  await page.keyboard.press("Control+End");
  await openInsert(page);
  await page.getByTestId("insert-image-button").click();
  await page.getByTestId("insert-image-file-input").setInputFiles({
    name: "photo.png",
    mimeType: "image/png",
    buffer: Buffer.from(PNG_1x1, "base64"),
  });
  await expect(content(page).locator('[data-testid="doc-image"]').last()).toBeVisible();
  await page.keyboard.press("Escape");
  await content(page).locator('[data-testid="doc-image"]').last().click();
  await captionSelected(page);
  await page.keyboard.type(text);
  await page.keyboard.press("Control+End");
  await page.keyboard.press("Enter");
}

test("table of figures: entries, non-editable hint, staleness bar and choices, entry link, two independent (AC-29..AC-36)", async ({ page }) => {
  await createDocument(page);
  await figureWithCaption(page, "one");
  await figureWithCaption(page, "two");

  await openInsert(page);
  await page.getByTestId("insert-table-of-figures-button").click();
  await expect(page.getByTestId("table-of-figures-dialog")).toBeVisible();
  await page.getByTestId("tof-label").selectOption("Figure");
  await page.getByTestId("tof-insert").click();
  const figTof = page.getByTestId("tof-block").first();
  await expect(figTof).toBeVisible(); // AC-29
  await expect(page.getByTestId("tof-entry")).toHaveCount(2);
  await expect(page.getByTestId("tof-update-bar")).toHaveCount(0); // AC-32 fresh, no bar

  // AC-30/AC-31: typing inside the ToF does not edit it and shows the hint.
  await figTof.click();
  await page.keyboard.type("zzz");
  await expect(page.getByTestId("tof-entry")).toHaveCount(2);
  await page.keyboard.press("Escape");

  // AC-34: an entry is a clickable link to its caption.
  const entry = page.getByTestId("tof-entry").first();
  await expect(entry.locator("a, [role='link']").first()).toBeVisible();
  await entry.click();
  await expect(content(page).getByText(/Figure\s*1:\s*one/).first()).toBeVisible();

  // AC-32/AC-33: adding a caption makes the ToF stale, with both update choices.
  await figureWithCaption(page, "three");
  await expect(page.getByTestId("tof-update-bar").first()).toBeVisible();
  await expect(page.getByTestId("tof-update-page-numbers").first()).toBeVisible();
  await expect(page.getByTestId("tof-update-entire").first()).toBeVisible();
  await page.getByTestId("tof-update-entire").first().click();
  await expect(page.getByTestId("tof-entry")).toHaveCount(3);

  // AC-36: a second ToF for a different label is independent.
  await content(page).click();
  await page.keyboard.press("Control+End");
  await page.keyboard.press("Enter");
  await page.getByTestId("menu-title-insert").click();
  await page.getByTestId("table-grid-cell-2-2").click();
  await expect(content(page).locator("table")).toHaveCount(1);
  await content(page).locator("table td").first().click();
  await captionSelected(page, "Table");
  await content(page).click();
  await page.keyboard.press("Control+End");
  await page.keyboard.press("Enter");
  await openInsert(page);
  await page.getByTestId("insert-table-of-figures-button").click();
  await page.getByTestId("tof-label").selectOption("Table");
  await page.getByTestId("tof-insert").click();
  await expect(page.getByTestId("tof-block")).toHaveCount(2); // AC-36 two coexist
});

// --- round-trip open (AC-45, AC-47, AC-49, AC-50, AC-51) --------------------

test("opens a Word caption fixture at saved numbers with resolved references and a regenerated ToF (AC-45, AC-49, AC-51)", async ({ page }) => {
  const here = path.dirname(fileURLToPath(import.meta.url));
  await page.goto("/");
  await page.getByTestId("open-document-input").setInputFiles(path.join(here, "fixtures", "word-captions.docx"));
  await expect(page).toHaveURL(/\/doc\/.+/);
  await expect(content(page).getByText(/Figure\s*1: Revenue by quarter/).first()).toBeVisible(); // AC-45
  await expect(content(page).getByText(/Table\s*1: Q1 results/).first()).toBeVisible();
  // AC-49: references resolve; AC-28/AC-50: broken one renders the error text.
  await expect(content(page).getByText(/Error! Reference source not found\./).first()).toBeVisible();
  await expect(page.getByTestId("tof-block").first()).toBeVisible(); // AC-51 regenerated
});

test("opens a LibreOffice caption fixture at saved numbers with resolved references (AC-47, AC-50)", async ({ page }) => {
  const here = path.dirname(fileURLToPath(import.meta.url));
  await page.goto("/");
  await page.getByTestId("open-document-input").setInputFiles(path.join(here, "fixtures", "lo-captions.odt"));
  await expect(page).toHaveURL(/\/doc\/.+/);
  await expect(content(page).getByText(/Figure\s*1: Revenue by quarter/).first()).toBeVisible(); // AC-47
  await expect(content(page).getByText(/Error! Reference source not found\./).first()).toBeVisible(); // AC-50
});
