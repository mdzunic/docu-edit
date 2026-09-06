import { test, expect, type Page, type Download } from "@playwright/test";
import { readFileSync, writeFileSync } from "node:fs";
import { crc32, inflateRawSync } from "node:zlib";
import { MOD } from "./keys";

// SPEC-041 — document properties and metadata: File → Properties opens a
// four-tab dialog (General/Description/Statistics/Custom Properties);
// Description-tab edits and typed custom properties round-trip through
// .docx docProps/core.xml + docProps/custom.xml and .odt meta.xml; a
// "Remove personal information on save" toggle strips author/last-modified-
// by and anonymizes comment/tracked-change authors; the document identity
// used by comments and track changes is settable here and persists; and
// every field renders as untrusted plain text.

async function createDocument(page: Page): Promise<void> {
  await page.goto("/");
  await page.getByTestId("new-document").click();
  await expect(page).toHaveURL(/\/doc\/.+/);
}

const content = (page: Page) => page.getByTestId("document-content");
const dialog = (page: Page) => page.getByTestId("properties-dialog");

async function openProperties(page: Page): Promise<void> {
  await page.getByTestId("menu-title-file").click();
  await page.getByTestId("file-menu-properties").click();
  await expect(dialog(page)).toBeVisible();
}

async function switchTab(page: Page, tab: "general" | "description" | "statistics" | "custom"): Promise<void> {
  await page.getByTestId(`properties-tab-${tab}`).click();
  await expect(page.getByTestId(`properties-panel-${tab}`)).toBeVisible();
}

/** Downloads the document in the given format and returns the file path
 * (SPEC-013's own helper, reused here for the save/reopen round trips). */
async function download(page: Page, format: "docx" | "odt"): Promise<string> {
  await page.getByTestId("menu-title-file").click();
  await page.getByTestId("file-menu-export").click();
  const file = await Promise.all([
    page.waitForEvent("download"),
    page.getByTestId(`download-${format}`).click(),
  ]).then(([d]: [Download, unknown]) => d);
  const path = await file.path();
  expect(path).toBeTruthy();
  return path!;
}

async function reopen(page: Page, path: string): Promise<void> {
  await page.goto("/");
  await page.getByTestId("open-document-input").setInputFiles(path);
  await expect(page).toHaveURL(/\/doc\/.+/);
  // The URL changes the moment the document is stored — before the parse, and
  // well before layout: without a barrier here a test reads state the app has
  // not built yet. Same pair SPEC-030/051 wait on.
  await expect(page.getByTestId("document-opening")).toBeHidden();
  await expect(page.getByTestId("document-pages")).toHaveAttribute("data-pagination-state", "clean");
}

/** Minimal central-directory walk over a downloaded archive (same shape as
 * e2e/SPEC-039.spec.ts) — enough to patch a single part of a downloaded
 * .docx from Node before reopening it. */
function unzipParts(buffer: Buffer): Map<string, Buffer> {
  let eocd = buffer.length - 22;
  while (eocd >= 0 && buffer.readUInt32LE(eocd) !== 0x06054b50) {
    eocd--;
  }
  const totalEntries = buffer.readUInt16LE(eocd + 10);
  let offset = buffer.readUInt32LE(eocd + 16);
  const parts = new Map<string, Buffer>();
  for (let i = 0; i < totalEntries; i++) {
    const method = buffer.readUInt16LE(offset + 10);
    const compressedSize = buffer.readUInt32LE(offset + 20);
    const nameLength = buffer.readUInt16LE(offset + 28);
    const extraLength = buffer.readUInt16LE(offset + 30);
    const commentLength = buffer.readUInt16LE(offset + 32);
    const localOffset = buffer.readUInt32LE(offset + 42);
    const name = buffer.toString("utf8", offset + 46, offset + 46 + nameLength);
    const localNameLength = buffer.readUInt16LE(localOffset + 26);
    const localExtraLength = buffer.readUInt16LE(localOffset + 28);
    const dataStart = localOffset + 30 + localNameLength + localExtraLength;
    const data = buffer.subarray(dataStart, dataStart + compressedSize);
    parts.set(name, method === 8 ? inflateRawSync(data) : Buffer.from(data));
    offset += 46 + nameLength + extraLength + commentLength;
  }
  return parts;
}

// Minimal store-method (uncompressed) zip writer, same shape as
// e2e/SPEC-039.spec.ts's own copy.
function buildZip(parts: Map<string, Buffer>): Buffer {
  const localChunks: Buffer[] = [];
  const centralChunks: Buffer[] = [];
  let offset = 0;
  for (const [name, data] of parts) {
    const nameBuf = Buffer.from(name, "utf8");
    const crc = crc32(data) >>> 0;
    const local = Buffer.alloc(30);
    local.writeUInt32LE(0x04034b50, 0);
    local.writeUInt16LE(20, 4);
    local.writeUInt16LE(0, 6);
    local.writeUInt16LE(0, 8);
    local.writeUInt16LE(0, 10);
    local.writeUInt16LE(0, 12);
    local.writeUInt32LE(crc, 14);
    local.writeUInt32LE(data.length, 18);
    local.writeUInt32LE(data.length, 22);
    local.writeUInt16LE(nameBuf.length, 26);
    local.writeUInt16LE(0, 28);
    localChunks.push(local, nameBuf, data);

    const central = Buffer.alloc(46);
    central.writeUInt32LE(0x02014b50, 0);
    central.writeUInt16LE(20, 4);
    central.writeUInt16LE(20, 6);
    central.writeUInt16LE(0, 8);
    central.writeUInt16LE(0, 10);
    central.writeUInt16LE(0, 12);
    central.writeUInt16LE(0, 14);
    central.writeUInt32LE(crc, 16);
    central.writeUInt32LE(data.length, 20);
    central.writeUInt32LE(data.length, 24);
    central.writeUInt16LE(nameBuf.length, 28);
    central.writeUInt16LE(0, 30);
    central.writeUInt16LE(0, 32);
    central.writeUInt16LE(0, 34);
    central.writeUInt16LE(0, 36);
    central.writeUInt32LE(0, 38);
    central.writeUInt32LE(offset, 42);
    centralChunks.push(central, nameBuf);
    offset += local.length + nameBuf.length + data.length;
  }
  const centralStart = offset;
  const centralBuf = Buffer.concat(centralChunks);
  const eocd = Buffer.alloc(22);
  eocd.writeUInt32LE(0x06054b50, 0);
  eocd.writeUInt16LE(parts.size, 8);
  eocd.writeUInt16LE(parts.size, 10);
  eocd.writeUInt32LE(centralBuf.length, 12);
  eocd.writeUInt32LE(centralStart, 16);
  return Buffer.concat([...localChunks, centralBuf, eocd]);
}

test("File → Properties opens a modal dialog with General/Description/Statistics/Custom Properties tabs (AC-1)", async ({
  page,
}) => {
  await createDocument(page);
  await openProperties(page);
  for (const tab of ["general", "description", "statistics", "custom"]) {
    await expect(page.getByTestId(`properties-tab-${tab}`)).toBeVisible();
  }
  await expect(page.getByTestId("properties-panel-general")).toBeVisible();
});

test("the dialog closes on the Close button and on Escape", async ({ page }) => {
  await createDocument(page);
  await openProperties(page);
  await page.getByTestId("properties-close").click();
  await expect(dialog(page)).toHaveCount(0);

  await openProperties(page);
  await page.keyboard.press("Escape");
  await expect(dialog(page)).toHaveCount(0);
});

test("for a new, unsaved document the General/Description/Statistics tabs are blank or zeroed (AC-18a, AC-18b, AC-18c)", async ({
  page,
}) => {
  await createDocument(page);
  await openProperties(page);
  await expect(page.getByTestId("properties-general-filename")).toHaveText("");
  await expect(page.getByTestId("properties-general-format")).toHaveText("");
  await expect(page.getByTestId("properties-general-size")).toHaveText("");
  await expect(page.getByTestId("properties-general-created")).toHaveText("");
  await expect(page.getByTestId("properties-general-modified")).toHaveText("");

  await switchTab(page, "description");
  await expect(page.getByTestId("properties-title-input")).toHaveValue("");
  await expect(page.getByTestId("properties-author-input")).toHaveValue("");

  await switchTab(page, "statistics");
  await expect(page.getByTestId("properties-statistics-words")).toHaveText("0");
  await expect(page.getByTestId("properties-statistics-characters-with-spaces")).toHaveText("0");
});

test("General tab shows file name, format and a computed size for an opened/saved document (AC-2)", async ({
  page,
}) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("Some body text");
  const path = await download(page, "docx");
  await reopen(page, path);

  await openProperties(page);
  await expect(page.getByTestId("properties-general-format")).toHaveText("Word (.docx)");
  await expect(page.getByTestId("properties-general-filename")).toContainText(".docx");
  await expect(page.getByTestId("properties-general-size")).not.toHaveText("");
  await expect(page.getByTestId("properties-general-size")).not.toHaveText("Calculating…");
});

test("Description-tab edits are written to .docx metadata and read back on reopen (AC-3, AC-5)", async ({ page }) => {
  await createDocument(page);
  await openProperties(page);
  await switchTab(page, "description");
  await page.getByTestId("properties-title-input").fill("Quarterly Report");
  await page.getByTestId("properties-subject-input").fill("Finance");
  await page.getByTestId("properties-keywords-input").fill("q3, finance");
  await page.getByTestId("properties-comments-input").fill("Draft for review");
  await page.getByTestId("properties-author-input").fill("Ada Lovelace");
  await page.getByTestId("properties-last-modified-by-input").fill("Grace Hopper");
  // Blur the last field to commit its @change.
  await page.keyboard.press("Tab");
  await page.getByTestId("properties-close").click();

  const path = await download(page, "docx");
  await reopen(page, path);
  await openProperties(page);
  await switchTab(page, "description");
  await expect(page.getByTestId("properties-title-input")).toHaveValue("Quarterly Report");
  await expect(page.getByTestId("properties-subject-input")).toHaveValue("Finance");
  await expect(page.getByTestId("properties-keywords-input")).toHaveValue("q3, finance");
  await expect(page.getByTestId("properties-comments-input")).toHaveValue("Draft for review");
  await expect(page.getByTestId("properties-author-input")).toHaveValue("Ada Lovelace");
  await expect(page.getByTestId("properties-last-modified-by-input")).toHaveValue("Grace Hopper");
});

test("an unmodelled core.xml field is shown with its actual value, not its field name (AC-6)", async ({ page }) => {
  await createDocument(page);
  const path = await download(page, "docx");

  // The editor never writes cp:category itself (AC-6: "including fields the
  // editor does not expose for editing"), so hand-patch a downloaded docx
  // to add one, the same way e2e/SPEC-039.spec.ts hand-builds fixtures for
  // states the app can't reach on its own.
  const parts = unzipParts(readFileSync(path));
  const core = parts.get("docProps/core.xml")!.toString("utf8");
  parts.set(
    "docProps/core.xml",
    Buffer.from(core.replace("</cp:coreProperties>", "<cp:category>Finance</cp:category></cp:coreProperties>"), "utf8"),
  );
  const fixturePath = test.info().outputPath("spec041-ac6-fixture.docx");
  writeFileSync(fixturePath, buildZip(parts));

  await reopen(page, fixturePath);
  await openProperties(page);
  await switchTab(page, "description");
  const extraFields = page.getByTestId("properties-extra-fields");
  await expect(extraFields).toContainText("Category");
  await expect(extraFields).toContainText("Finance");
  await expect(extraFields).not.toContainText("category");
});

test("Description-tab edits are written to .odt metadata and read back on reopen (AC-5)", async ({ page }) => {
  await createDocument(page);
  await openProperties(page);
  await switchTab(page, "description");
  await page.getByTestId("properties-title-input").fill("Field Report");
  await page.getByTestId("properties-author-input").fill("Ada Lovelace");
  await page.keyboard.press("Tab");
  await page.getByTestId("properties-close").click();

  const path = await download(page, "odt");
  await reopen(page, path);
  await openProperties(page);
  await switchTab(page, "description");
  await expect(page.getByTestId("properties-title-input")).toHaveValue("Field Report");
  await expect(page.getByTestId("properties-author-input")).toHaveValue("Ada Lovelace");
});

test("the Statistics tab shows the SPEC-036 statistics (AC-4)", async ({ page }) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("Hello world. This is a test.");
  await openProperties(page);
  await switchTab(page, "statistics");
  await expect(page.getByTestId("properties-statistics-words")).toHaveText("6");
  await expect(page.getByTestId("properties-statistics-sentences")).toHaveText("2");
  await expect(page.getByTestId("properties-statistics-pages")).toHaveText("1");
});

test("Custom Properties: text, number, date and boolean properties can be added and round-trip through save/reopen (AC-7, AC-8, AC-9)", async ({
  page,
}) => {
  await createDocument(page);
  await openProperties(page);
  await switchTab(page, "custom");

  async function addProperty(name: string, type: string, value: string): Promise<void> {
    await page.getByTestId("custom-property-new-name").fill(name);
    await page.getByTestId("custom-property-new-type").selectOption(type);
    if (type === "boolean") {
      await page.getByTestId("custom-property-new-value").selectOption(value);
    } else {
      await page.getByTestId("custom-property-new-value").fill(value);
    }
    await page.getByTestId("custom-property-add").click();
  }

  await addProperty("Client", "text", "Acme Corp");
  await addProperty("Budget", "number", "4200");
  await addProperty("Due", "date", "2026-12-31");
  await addProperty("Approved", "boolean", "true");

  const rows = page.getByTestId("custom-property-row");
  await expect(rows).toHaveCount(4);

  await page.getByTestId("properties-close").click();
  const path = await download(page, "docx");
  await reopen(page, path);
  await openProperties(page);
  await switchTab(page, "custom");
  await expect(page.getByTestId("custom-property-row")).toHaveCount(4);
  const nameInputs = await page.getByTestId("custom-property-name").all();
  const names = await Promise.all(nameInputs.map((input) => input.inputValue()));
  expect(names.sort()).toEqual(["Approved", "Budget", "Client", "Due"]);
});

test("Custom Properties: an invalid number value shows an inline error and is not applied (AC-11)", async ({
  page,
}) => {
  await createDocument(page);
  await openProperties(page);
  await switchTab(page, "custom");
  await page.getByTestId("custom-property-new-name").fill("Budget");
  await page.getByTestId("custom-property-new-type").selectOption("number");
  await page.getByTestId("custom-property-new-value").fill("not a number");
  await expect(page.getByTestId("custom-property-new-value-error")).toBeVisible();
  await expect(page.getByTestId("custom-property-add")).toBeDisabled();

  // Add a valid one, then type an invalid edit into its value — the error
  // shows and the committed model keeps the last valid value.
  await page.getByTestId("custom-property-new-value").fill("100");
  await page.getByTestId("custom-property-add").click();
  await page.getByTestId("custom-property-value").fill("still not a number");
  await expect(page.getByTestId("custom-property-value-error")).toBeVisible();

  await page.getByTestId("properties-close").click();
  const path = await download(page, "docx");
  await reopen(page, path);
  await openProperties(page);
  await switchTab(page, "custom");
  await expect(page.getByTestId("custom-property-value")).toHaveValue("100");
});

test("Custom Properties: removing a property drops it after save/reopen (AC-8)", async ({ page }) => {
  await createDocument(page);
  await openProperties(page);
  await switchTab(page, "custom");
  await page.getByTestId("custom-property-new-name").fill("Client");
  await page.getByTestId("custom-property-new-value").fill("Acme");
  await page.getByTestId("custom-property-add").click();
  await expect(page.getByTestId("custom-property-row")).toHaveCount(1);
  await page.getByTestId("custom-property-remove").click();
  await expect(page.getByTestId("custom-property-row")).toHaveCount(0);
  await page.getByTestId("properties-close").click();

  const path = await download(page, "docx");
  await reopen(page, path);
  await openProperties(page);
  await switchTab(page, "custom");
  await expect(page.getByTestId("custom-property-row")).toHaveCount(0);
});

test("the 'Remove personal information on save' toggle is present and states exactly what it removes (AC-12a, AC-12b)", async ({
  page,
}) => {
  await createDocument(page);
  await openProperties(page);
  const toggle = page.getByTestId("properties-remove-personal-info-toggle");
  await expect(toggle).toBeVisible();
  await expect(page.getByTestId("properties-remove-personal-info-note")).toHaveCount(0);
  await toggle.click();
  const note = page.getByTestId("properties-remove-personal-info-note");
  await expect(note).toBeVisible();
  await expect(note).toContainText("author");
  await expect(note).toContainText("last-modified-by");
  await expect(note).toContainText("Author 1");
  await expect(note).toContainText("revision timestamps");
});

test("with the toggle on, saving anonymizes comment authors and clears author/last-modified-by (AC-13, AC-14)", async ({
  page,
}) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("Hello world");
  await page.keyboard.press("Home");
  await page.keyboard.down("Shift");
  for (let i = 0; i < 5; i++) {
    await page.keyboard.press("ArrowRight");
  }
  await page.keyboard.up("Shift");
  await page.keyboard.press(`${MOD}+Alt+m`);
  await page.getByTestId("comments-pane").getByTestId("comment-composer-input").fill("A note");
  await page.getByTestId("comments-pane").getByTestId("comment-composer-submit").click();
  await expect(page.getByTestId("comments-pane").locator('[data-testid^="comment-thread-"]')).toHaveCount(1);

  await openProperties(page);
  await switchTab(page, "description");
  await page.getByTestId("properties-author-input").fill("Ada Lovelace");
  await page.keyboard.press("Tab");
  await page.getByTestId("properties-remove-personal-info-toggle").click();
  await page.getByTestId("properties-close").click();

  const path = await download(page, "docx");
  await reopen(page, path);

  await openProperties(page);
  await switchTab(page, "description");
  await expect(page.getByTestId("properties-author-input")).toHaveValue("");

  await page.getByTestId("properties-close").click();
  // The comments pane's open/closed state is a persisted per-browser
  // preference (independent of any one document), so it may already be
  // open from earlier in this test — the menu item toggles it, so only
  // click if it isn't already showing.
  if (!(await page.getByTestId("comments-pane").isVisible())) {
    await page.getByTestId("menu-title-view").click();
    await page.getByTestId("view-menu-comments-pane").click();
  }
  const card = page.getByTestId("comments-pane").locator('[data-testid^="comment-thread-"]').first();
  await expect(card).toContainText("Author 1");
});

test("the document identity persists across a reload (AC-15, AC-16)", async ({ page }) => {
  await createDocument(page);
  await openProperties(page);
  // Defaults to a placeholder identity, not a real OS user name (no browser
  // API exposes one) — see the implementation summary for the reasoning.
  await expect(page.getByTestId("properties-identity-input")).toHaveValue("You");
  await page.getByTestId("properties-identity-input").fill("Ada Lovelace");
  await page.keyboard.press("Tab");
  await page.getByTestId("properties-close").click();

  await page.reload();
  await openProperties(page);
  await expect(page.getByTestId("properties-identity-input")).toHaveValue("Ada Lovelace");
});

test("metadata fields render markup and formula payloads as literal text, never executed (AC-17)", async ({
  page,
}) => {
  await createDocument(page);
  const payload = '<img src=x onerror="window.__xss=true">=1+1';
  await openProperties(page);
  await switchTab(page, "description");
  await page.getByTestId("properties-title-input").fill(payload);
  await page.keyboard.press("Tab");

  // No script executed and no stray element was injected into the dialog.
  const xssRan = await page.evaluate(() => (window as unknown as { __xss?: boolean }).__xss);
  expect(xssRan).toBeUndefined();
  await expect(dialog(page).locator("img")).toHaveCount(0);
  await expect(page.getByTestId("properties-title-input")).toHaveValue(payload);

  await switchTab(page, "custom");
  await page.getByTestId("custom-property-new-name").fill(payload);
  await page.getByTestId("custom-property-new-value").fill(payload);
  await page.getByTestId("custom-property-add").click();
  await expect(dialog(page).locator("img")).toHaveCount(0);
  await expect(page.getByTestId("custom-property-name").first()).toHaveValue(payload);
});
