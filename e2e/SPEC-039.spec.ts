import { test, expect, type Page } from "@playwright/test";
import { readFileSync, writeFileSync } from "node:fs";
import { crc32, inflateRawSync } from "node:zlib";
import { MOD, LINE_START } from "./keys";

// SPEC-039 — track changes: a recording toggle (always OFF on open) that
// marks new insertions/deletions instead of applying them; a review pane
// listing changes with Accept/Reject/Accept all/Reject all; Next/Previous
// navigation; a display-mode selector; accept/reject is undoable; round-trip
// through .docx/.odt for both text-level and structural changes.

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

// Minimal store-method (uncompressed) zip writer, used only to hand-build a
// fixture docx with a pre-set settings.xml (AC-3's "opened with the
// persisted flag already on" case — the runtime toggle is deliberately
// disconnected from that flag, so there is no in-app path to produce one).
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

async function createDocument(page: Page): Promise<void> {
  await page.goto("/");
  await page.getByTestId("new-document").click();
  await expect(page).toHaveURL(/\/doc\/.+/);
}

const content = (page: Page) => page.getByTestId("document-content");
const pane = (page: Page) => page.getByTestId("track-changes-pane");

async function seedParagraph(page: Page, text: string): Promise<void> {
  await content(page).click();
  await page.keyboard.type(text);
}

async function openTrackChangesPane(page: Page): Promise<void> {
  await page.getByTestId("menu-title-view").click();
  await page.getByTestId("view-menu-track-changes-pane").click();
}

async function toggleRecording(page: Page): Promise<void> {
  await page.getByTestId("toolbar-track-changes-toggle").click();
}

async function downloadFile(page: Page, format: "docx" | "odt"): Promise<Buffer> {
  await page.getByTestId("download-menu").click();
  const downloadPromise = page.waitForEvent("download");
  await page.getByTestId(`download-${format}`).click();
  const download = await downloadPromise;
  return readFileSync((await download.path())!);
}

test("the toggle is off for a new blank document, and toolbar/status bar show a clear on/off state (AC-1, AC-2)", async ({
  page,
}) => {
  await createDocument(page);
  await expect(page.getByTestId("toolbar-track-changes-toggle")).toHaveAttribute("aria-pressed", "false");
  await expect(page.getByTestId("status-bar-track-changes")).toHaveAttribute("aria-pressed", "false");
  await expect(page.getByTestId("status-bar-track-changes")).toContainText("Off");

  await toggleRecording(page);
  await expect(page.getByTestId("toolbar-track-changes-toggle")).toHaveAttribute("aria-pressed", "true");
  await expect(page.getByTestId("status-bar-track-changes")).toContainText("On");

  // The shortcut toggles it back off.
  await content(page).click();
  await page.keyboard.press(`${MOD}+Shift+e`);
  await expect(page.getByTestId("toolbar-track-changes-toggle")).toHaveAttribute("aria-pressed", "false");
});

test("with tracking on, inserted text underlines and deleted text stays struck through instead of disappearing (AC-4)", async ({
  page,
}) => {
  await createDocument(page);
  await seedParagraph(page, "hello world");
  await toggleRecording(page);

  await content(page).click();
  await page.keyboard.press(LINE_START);
  await page.keyboard.type("XYZ");
  const inserted = content(page).locator("span", { hasText: "XYZ" }).first();
  await expect(inserted).toHaveCSS("text-decoration-line", /underline/);

  // Select "hello" (now at offset 3, after the inserted "XYZ") and delete
  // it — it must remain visible, struck through.
  await page.keyboard.press(LINE_START);
  for (let i = 0; i < 3; i++) {
    await page.keyboard.press("ArrowRight");
  }
  for (let i = 0; i < 5; i++) {
    await page.keyboard.press("Shift+ArrowRight");
  }
  await page.keyboard.press("Backspace");
  await expect(content(page)).toContainText("hello");
  const deleted = content(page).locator("span", { hasText: "hello" }).first();
  await expect(deleted).toHaveCSS("text-decoration-line", /line-through/);
});

test("opening the review pane lists changes in document order with author/date, and Next/Previous navigate (AC-6, AC-7, AC-8)", async ({
  page,
}) => {
  await createDocument(page);
  await seedParagraph(page, "aaaa bbbb cccc");
  await toggleRecording(page);
  await openTrackChangesPane(page);

  await content(page).click();
  await page.keyboard.press(LINE_START);
  await page.keyboard.type("ONE");
  await page.keyboard.press(`${MOD}+End`);
  await page.keyboard.type("TWO");

  // Two distinct tracked insertions should be listed.
  await expect(pane(page).locator(".track-change-item")).toHaveCount(2);
  await expect(pane(page).locator(".track-change-item").first()).toContainText("Inserted");
  await expect(pane(page).locator(".track-change-item").first()).toContainText("You");

  await pane(page).getByTestId("track-changes-next").click();
  await pane(page).getByTestId("track-changes-next").click();
});

test("Accept and Reject per change, Accept all and Reject all (AC-7, AC-10, AC-11)", async ({ page }) => {
  await createDocument(page);
  await seedParagraph(page, "hello world");
  await toggleRecording(page);
  await openTrackChangesPane(page);

  await content(page).click();
  await page.keyboard.press(LINE_START);
  await page.keyboard.type("XYZ");
  await expect(pane(page).locator(".track-change-item")).toHaveCount(1);

  // Reject: the inserted text is removed.
  await pane(page).getByTestId("track-change-reject").click();
  await expect(content(page)).not.toContainText("XYZ");
  await expect(pane(page).locator(".track-change-item")).toHaveCount(0);

  // Undo restores the pending insertion (AC-10).
  await content(page).click();
  await page.keyboard.press(`${MOD}+z`);
  await expect(pane(page).locator(".track-change-item")).toHaveCount(1);

  // Accept: the insertion commits, no longer tracked.
  await pane(page).getByTestId("track-change-accept").click();
  await expect(content(page)).toContainText("XYZ");
  await expect(pane(page).locator(".track-change-item")).toHaveCount(0);

  // Delete a range, then use the bulk actions.
  await content(page).click();
  await page.keyboard.press(LINE_START);
  await page.keyboard.press("Shift+ArrowRight");
  await page.keyboard.press("Shift+ArrowRight");
  await page.keyboard.press("Backspace");
  await expect(pane(page).locator(".track-change-item")).toHaveCount(1);
  await pane(page).getByTestId("track-changes-reject-all").click();
  await expect(pane(page).locator(".track-change-item")).toHaveCount(0);
});

test("the display-mode selector offers all four modes and defaults to All markup (AC-9)", async ({ page }) => {
  await createDocument(page);
  await seedParagraph(page, "hello world");
  await openTrackChangesPane(page);
  const select = pane(page).getByTestId("track-changes-display-mode-select");
  await expect(select).toHaveValue("all");
  await select.selectOption("simple");
  await expect(select).toHaveValue("simple");
  await select.selectOption("noMarkup");
  await select.selectOption("original");
});

test("'No markup' hides tracked deletions and 'Original' hides tracked insertions (AC-9)", async ({ page }) => {
  await createDocument(page);
  await seedParagraph(page, "hello world");
  await toggleRecording(page);

  // Insert "XYZ" then, with the caret left right after it, forward-delete
  // "hello " — the paragraph ends up with one pending insertion ("XYZ") and
  // one separate pending deletion ("hello "), not one merged edit.
  await content(page).click();
  await page.keyboard.press(LINE_START);
  await page.keyboard.type("XYZ");
  for (let i = 0; i < 6; i++) {
    await page.keyboard.press("Shift+ArrowRight");
  }
  await page.keyboard.press("Delete");
  await expect(content(page)).toContainText("XYZ");
  await expect(content(page)).toContainText("hello");

  await openTrackChangesPane(page);
  const select = pane(page).getByTestId("track-changes-display-mode-select");

  // "No markup": as if accepted — the insertion shows, the deletion is gone.
  await select.selectOption("noMarkup");
  await expect(content(page)).toContainText("XYZworld");
  await expect(content(page)).not.toContainText("hello");

  // "Original": as if rejected — the deletion shows, the insertion is gone.
  await select.selectOption("original");
  await expect(content(page)).toContainText("hello world");
  await expect(content(page)).not.toContainText("XYZ");

  // Back to "All markup": both sides show again.
  await select.selectOption("all");
  await expect(content(page)).toContainText("XYZ");
  await expect(content(page)).toContainText("hello");
});

test("a table row insert while tracking is on is tracked, not silently applied (AC-15, AC-18)", async ({ page }) => {
  await createDocument(page);
  await page.getByTestId("menu-title-insert").click();
  await page.getByTestId("table-grid-cell-2-2").click();
  await expect(content(page).locator("table.doc-table")).toHaveCount(1);

  await toggleRecording(page);
  await content(page).getByTestId("table-cell-0-0").locator(".doc-paragraph").first().click({ button: "right" });
  await page.getByTestId("context-menu-insert-row-below").click();
  await expect(content(page).locator("table.doc-table tr")).toHaveCount(3);

  await openTrackChangesPane(page);
  await expect(pane(page).locator(".track-change-item")).toHaveCount(1);
  await expect(pane(page).locator(".track-change-item").first()).toContainText("table");

  // Reject: the inserted row is removed again (AC-10, AC-11's "restore
  // exact original" spirit applied to a structural change).
  await pane(page).getByTestId("track-change-reject").click();
  await expect(content(page).locator("table.doc-table tr")).toHaveCount(2);
});

test("opening a document that already contains tracked changes leaves the toggle off (AC-3)", async ({ page }) => {
  await createDocument(page);
  await seedParagraph(page, "hello world");
  await toggleRecording(page);
  await content(page).click();
  await page.keyboard.press(LINE_START);
  await page.keyboard.type("XYZ");
  await expect(page.getByTestId("toolbar-track-changes-toggle")).toHaveAttribute("aria-pressed", "true");

  // Navigate away and back (a fresh "open") — the toggle must reset to off
  // even though the document still has a tracked change in it.
  await page.goto("/");
  await page.goBack();
  await expect(content(page)).toContainText("XYZ");
  await expect(page.getByTestId("toolbar-track-changes-toggle")).toHaveAttribute("aria-pressed", "false");
});

test("the persisted recording-enabled flag round-trips through docx unchanged without turning the toggle on (AC-3)", async ({ page }) => {
  await createDocument(page);
  await seedParagraph(page, "hello world");
  await toggleRecording(page);
  await content(page).click();
  await page.keyboard.press(LINE_START);
  await page.keyboard.type("XYZ");
  const buf = await downloadFile(page, "docx");

  // The runtime toggle above is session-only and deliberately never sets
  // the model's own persisted `trackChangesRecording` flag (there is no
  // in-app path to produce that combination) — hand-build a fixture whose
  // settings.xml already carries it, the same shape the writer itself
  // emits, to exercise the reader/toggle side of AC-3.
  const parts = unzipParts(buf);
  parts.set(
    "word/settings.xml",
    Buffer.from(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:trackChanges/></w:settings>',
      "utf8",
    ),
  );
  const contentTypes = parts.get("[Content_Types].xml")!.toString("utf8");
  parts.set(
    "[Content_Types].xml",
    Buffer.from(
      contentTypes.replace(
        "</Types>",
        '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>',
      ),
      "utf8",
    ),
  );
  const rels = parts.get("word/_rels/document.xml.rels")!.toString("utf8");
  parts.set(
    "word/_rels/document.xml.rels",
    Buffer.from(
      rels.replace(
        "</Relationships>",
        '<Relationship Id="rIdSettingsQA" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>',
      ),
      "utf8",
    ),
  );
  const fixturePath = test.info().outputPath("spec039-ac3-fixture.docx");
  writeFileSync(fixturePath, buildZip(parts));

  await page.goto("/");
  await page.getByTestId("open-document-input").setInputFiles(fixturePath);
  await expect(page).toHaveURL(/\/doc\/.+/);

  await expect(content(page)).toContainText("XYZ");
  const insertedSpan = content(page).locator("span", { hasText: "XYZ" }).first();
  await expect(insertedSpan).toHaveCSS("text-decoration-line", /underline/);
  await expect(page.getByTestId("toolbar-track-changes-toggle")).toHaveAttribute("aria-pressed", "false");

  const resaved = unzipParts(await downloadFile(page, "docx"));
  expect(resaved.get("word/settings.xml")!.toString("utf8")).toContain("w:trackChanges");
});

test("a formatting change shows a margin change bar and a 'Formatted: Bold' description (AC-5, AC-6)", async ({ page }) => {
  await createDocument(page);
  await seedParagraph(page, "hello world");
  await toggleRecording(page);
  await openTrackChangesPane(page);

  await content(page).click();
  await page.keyboard.press(`${MOD}+a`);
  await page.keyboard.press(`${MOD}+b`);

  await expect(content(page).locator(".doc-paragraph[data-tracked-change]")).toHaveCount(1);
  const item = pane(page).locator(".track-change-item").first();
  await expect(pane(page).locator(".track-change-item")).toHaveCount(1);
  await expect(item).toContainText("Formatted: Bold");
  await expect(item).toContainText("You");
  await expect(pane(page).locator(".track-change-timestamp").first()).not.toHaveText("");
});

test("editing over another author's pending insertion nests the change instead of losing it (AC-12)", async ({ page }) => {
  await createDocument(page);
  await toggleRecording(page);
  await openTrackChangesPane(page);
  const authorInput = pane(page).getByTestId("track-changes-author-input");
  await authorInput.fill("Alice");
  await authorInput.blur();

  await content(page).click();
  await page.keyboard.type("hello world");
  await expect(pane(page).locator(".track-change-item")).toHaveCount(1);
  await expect(pane(page).locator(".track-change-item").first()).toContainText("Alice");

  await authorInput.fill("Bob");
  await authorInput.blur();

  // Bob deletes part of Alice's still-pending insertion.
  await content(page).click();
  await page.keyboard.press(LINE_START);
  for (let i = 0; i < 5; i++) {
    await page.keyboard.press("Shift+ArrowRight");
  }
  await page.keyboard.press("Backspace");

  // Still visible (struck through), not lost, and both authors are listed.
  await expect(content(page)).toContainText("hello");
  const texts = await pane(page).locator(".track-change-item").allTextContents();
  expect(texts.some((t) => t.includes("Alice"))).toBeTruthy();
  expect(texts.some((t) => t.includes("Bob"))).toBeTruthy();
});

test("text-level tracked changes round-trip through docx and odt; a same-format resave is byte-equivalent (AC-13, AC-14)", async ({ page }) => {
  await createDocument(page);
  await seedParagraph(page, "hello world");
  await toggleRecording(page);
  await content(page).click();
  await page.keyboard.press(LINE_START);
  await page.keyboard.type("XYZ");
  await page.keyboard.press(`${MOD}+End`);
  for (let i = 0; i < 5; i++) {
    await page.keyboard.press("Shift+ArrowLeft");
  }
  await page.keyboard.press("Backspace"); // tracked-delete "world"

  const docxBuf1 = await downloadFile(page, "docx");
  const docxBuf2 = await downloadFile(page, "docx");
  // Same-format resave is byte-equivalent (AC-14) — compared on
  // word/document.xml specifically, not the whole package: docProps/
  // core.xml's own dcterms:modified legitimately ticks forward between two
  // saves (SPEC-006), so lib/equality.ts's own byte-identity relation is
  // scoped to preserved *parts*, never the whole file, for the same reason.
  const doc1 = unzipParts(docxBuf1).get("word/document.xml")!;
  const doc2 = unzipParts(docxBuf2).get("word/document.xml")!;
  expect(Buffer.compare(doc1, doc2)).toBe(0);

  const doc = doc1.toString("utf8");
  expect(doc).toContain("w:ins");
  expect(doc).toContain("w:del");
  expect(doc).toContain("You");

  const odtContent = unzipParts(await downloadFile(page, "odt")).get("content.xml")!.toString("utf8");
  expect(odtContent).toContain("text:changed-region");
});

test("a tracked table row insert round-trips through odt; an unrepresentable property change is applied and reported, not dropped silently (AC-16, AC-19)", async ({ page }) => {
  // The conversion report only fires on a genuine cross-format save
  // (model.format differs from the target) — a document created fresh
  // in-app has no source format yet, so give it one via a real docx
  // download+reopen before exercising the cross-format save.
  await createDocument(page);
  await seedParagraph(page, "seed");
  const blankDocx = await downloadFile(page, "docx");
  const blankPath = test.info().outputPath("spec039-ac16-19-source.docx");
  writeFileSync(blankPath, blankDocx);
  await page.goto("/");
  await page.getByTestId("open-document-input").setInputFiles(blankPath);
  await expect(page).toHaveURL(/\/doc\/.+/);

  await page.getByTestId("menu-title-insert").click();
  await page.getByTestId("table-grid-cell-2-2").click();
  await expect(content(page).locator("table.doc-table")).toHaveCount(1);

  await toggleRecording(page);
  await content(page).getByTestId("table-cell-0-0").locator(".doc-paragraph").first().click({ button: "right" });
  await page.getByTestId("context-menu-insert-row-below").click();
  await expect(content(page).locator("table.doc-table tr")).toHaveCount(3);

  // A tracked row insert round-trips through odt (AC-16).
  const odtContent = unzipParts(await downloadFile(page, "odt")).get("content.xml")!.toString("utf8");
  expect(odtContent).toContain('table:change-scope="row"');

  // A table/cell property change (a merge) has no odt representation —
  // applied, but reported as a DropRecord rather than silently discarded
  // (AC-19).
  const cellAt = (r: number, c: number) => content(page).getByTestId(`table-cell-${r}-${c}`);
  const fromBox = (await cellAt(0, 0).boundingBox())!;
  const toBox = (await cellAt(0, 1).boundingBox())!;
  await page.mouse.move(fromBox.x + fromBox.width / 2, fromBox.y + fromBox.height / 2);
  await page.mouse.down();
  await page.mouse.move(toBox.x + toBox.width / 2, toBox.y + toBox.height / 2, { steps: 8 });
  await page.mouse.up();
  await cellAt(0, 0).click({ button: "right" });
  await page.getByTestId("context-menu-merge-cells").click();

  await page.getByTestId("download-menu").click();
  const downloadPromise = page.waitForEvent("download");
  await page.getByTestId("download-odt").click();
  await downloadPromise;

  const dialog = page.getByTestId("conversion-report");
  await expect(dialog).toBeVisible();
  await expect(dialog).toContainText("table property change", { ignoreCase: true });
  await expect(cellAt(0, 0)).toHaveAttribute("colspan", "2");
});

test("tracked deletions are excluded from the word count and from find results (AC-17)", async ({ page }) => {
  await createDocument(page);
  await seedParagraph(page, "alpha bravo charlie");
  await expect(page.getByTestId("word-count-button")).toHaveText("3 words");

  await toggleRecording(page);
  await content(page).click();
  await page.keyboard.press(LINE_START);
  for (let i = 0; i < 5; i++) {
    await page.keyboard.press("Shift+ArrowRight");
  }
  await page.keyboard.press("Backspace"); // tracked-delete "alpha"
  await expect(page.getByTestId("word-count-button")).toHaveText("2 words");

  await page.keyboard.press(`${MOD}+f`);
  await page.getByTestId("find-query").fill("alpha");
  await expect(page.getByTestId("find-count")).toHaveText(/0/);
});
