import { test, expect, type Page } from "@playwright/test";
import { deflateSync } from "node:zlib";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import { tmpdir } from "node:os";
import { writeFileSync, readFileSync } from "node:fs";
import { unzip } from "../src/lib/core/zip";
import { createDocument as createDocumentModel } from "../src/lib/documentModel";
import { writeDocx } from "../src/lib/docx/writer";
import { LINE_START, MOD } from "./keys";

// SPEC-024 — images: insert (file picker, drag-and-drop, paste), the
// scaled-to-fit initial size, click-to-select with eight resize handles and
// a floating toolbar, corner/edge resize with Shift-constrained aspect
// ratio, the four wrap modes, the "Image options" panel (alt text, size,
// lock aspect ratio, reset size), delete with reference-counted media
// cleanup, the 10 MB size limit, and opening the images.docx/images.odt
// fixtures (one inline, one floating image each).

const fixturesDir = join(dirname(fileURLToPath(import.meta.url)), "fixtures");
const docxImagesFixture = join(fixturesDir, "images.docx");
const odtImagesFixture = join(fixturesDir, "images.odt");

async function createDocument(page: Page): Promise<void> {
  await page.goto("/");
  await page.getByTestId("new-document").click();
  await expect(page).toHaveURL(/\/doc\/.+/);
}

async function openFixture(page: Page, path: string): Promise<void> {
  await page.goto("/");
  await page.getByTestId("open-document-input").setInputFiles(path);
  await expect(page).toHaveURL(/\/doc\/.+/);
}

const content = (page: Page) => page.getByTestId("document-content");
const images = (page: Page) => content(page).locator('[data-testid="doc-image"]');

// --- minimal PNG encoder (24-bit RGB, no filtering) — mirrors
// scripts/build-spec024-fixtures.mjs's own encoder, restated here so this
// spec file needs no built fixture for a plain insert test. ---------------

const CRC_TABLE = (() => {
  const table = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) {
      c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    }
    table[n] = c >>> 0;
  }
  return table;
})();

function crc32(buf: Buffer): number {
  let c = 0xffffffff;
  for (let i = 0; i < buf.length; i++) {
    c = CRC_TABLE[(c ^ buf[i]) & 0xff] ^ (c >>> 8);
  }
  return (c ^ 0xffffffff) >>> 0;
}

function pngChunk(type: string, data: Buffer): Buffer {
  const typeBuf = Buffer.from(type, "ascii");
  const body = Buffer.concat([typeBuf, data]);
  const crc = Buffer.alloc(4);
  crc.writeUInt32BE(crc32(body), 0);
  const length = Buffer.alloc(4);
  length.writeUInt32BE(data.length, 0);
  return Buffer.concat([length, body, crc]);
}

function buildPng(width: number, height: number, [r, g, b]: [number, number, number]): Buffer {
  const rowBytes = width * 3;
  const raw = Buffer.alloc((rowBytes + 1) * height);
  for (let y = 0; y < height; y++) {
    const rowStart = y * (rowBytes + 1);
    raw[rowStart] = 0;
    for (let x = 0; x < width; x++) {
      const px = rowStart + 1 + x * 3;
      raw[px] = r;
      raw[px + 1] = g;
      raw[px + 2] = b;
    }
  }
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(width, 0);
  ihdr.writeUInt32BE(height, 4);
  ihdr[8] = 8;
  ihdr[9] = 2;
  const signature = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
  return Buffer.concat([
    signature,
    pngChunk("IHDR", ihdr),
    pngChunk("IDAT", deflateSync(raw)),
    pngChunk("IEND", Buffer.alloc(0)),
  ]);
}

const SMALL_PNG = buildPng(200, 100, [0x33, 0x77, 0xbb]);

async function insertViaPicker(page: Page, buffer: Buffer, name = "photo.png", mimeType = "image/png"): Promise<void> {
  // SPEC-049 moved "Image…" into the Insert menu; the hidden file input it
  // drives lives alongside the menu bar and takes files directly.
  await page.getByTestId("insert-image-file-input").setInputFiles({ name, mimeType, buffer });
}

// --- insertion (AC-1, AC-4..AC-6) -------------------------------------------

test("inserts a PNG at the caret via the file picker, at its natural size (AC-1, AC-4, AC-6)", async ({ page }) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("before after");
  await page.keyboard.press(LINE_START);
  for (let i = 0; i < 6; i++) {
    await page.keyboard.press("ArrowRight");
  }
  // Caret is now between "before" and " after".
  await insertViaPicker(page, SMALL_PNG);
  await expect(images(page)).toHaveCount(1);
  const img = images(page).first();
  await expect(img).toBeVisible();
  // 200x100px at 96 DPI = 5.2917cm x 2.6458cm, well under the text area
  // width, so no scale-down applies (AC-6) — CSS "cm" is itself defined as
  // 96/2.54 px, so a 200px-natural image displayed at its natural size
  // renders at ~200 CSS px directly, not 200 times that ratio.
  const width = await img.evaluate((el) => (el as HTMLElement).getBoundingClientRect().width);
  expect(width).toBeGreaterThan(200 * 0.9);
  expect(width).toBeLessThan(200 * 1.1);
  // The surrounding text survived the insert, split around the image.
  await expect(content(page)).toContainText("before");
  await expect(content(page)).toContainText("after");
});

test("scales a wide image down to fit the text area width, preserving aspect ratio (AC-5)", async ({ page }) => {
  await createDocument(page);
  await content(page).click();
  const wide = buildPng(4000, 1000, [0x22, 0x99, 0x44]);
  await insertViaPicker(page, wide);
  const img = images(page).first();
  await expect(img).toBeVisible();
  const box = (await img.boundingBox())!;
  // A4 minus 2.54cm margins each side, in px at 96 DPI, is well under the
  // image's native 4000px width.
  const cmPx = 96 / 2.54;
  const textAreaPx = (21 - 2.54 * 2) * cmPx;
  expect(box.width).toBeLessThan(textAreaPx + 2);
  // Aspect ratio preserved (4:1).
  expect(box.width / box.height).toBeGreaterThan(3.9);
  expect(box.width / box.height).toBeLessThan(4.1);
});

test("rejects a file over the 10 MB limit with a toast and inserts nothing (AC-33, AC-34)", async ({ page }) => {
  await createDocument(page);
  await content(page).click();
  const oversized = Buffer.concat([SMALL_PNG, Buffer.alloc(10 * 1024 * 1024)]);
  const consoleErrors: string[] = [];
  page.on("console", (msg) => {
    if (msg.type() === "error") {
      consoleErrors.push(msg.text());
    }
  });
  await insertViaPicker(page, oversized);
  const toast = page.getByTestId("toast");
  await expect(toast).toContainText("That image is too large (limit 10 MB)");
  await expect(images(page)).toHaveCount(0);
  expect(consoleErrors).toEqual([]);
});

test("rejects an unsupported file type", async ({ page }) => {
  await createDocument(page);
  await content(page).click();
  await insertViaPicker(page, Buffer.from("not an image"), "notes.txt", "text/plain");
  await expect(images(page)).toHaveCount(0);
});

// --- selection, resize handles, floating toolbar (AC-7..AC-12) -------------

test("clicking an image selects it, showing eight resize handles and a floating toolbar (AC-7..AC-9)", async ({ page }) => {
  await createDocument(page);
  await content(page).click();
  await insertViaPicker(page, SMALL_PNG);
  await images(page).first().click();
  await expect(page.getByTestId("image-handles")).toBeVisible();
  for (const handle of ["nw", "n", "ne", "e", "se", "s", "sw", "w"]) {
    await expect(page.getByTestId(`image-resize-${handle}`)).toBeVisible();
  }
  await expect(page.getByTestId("image-toolbar")).toBeVisible();
});

test("the floating toolbar clears the top resize handles instead of covering them (AC-9)", async ({ page }) => {
  await createDocument(page);
  await content(page).click();
  // Push the image down the page so there is room for the toolbar above it.
  for (let i = 0; i < 4; i++) {
    await page.keyboard.press("Enter");
  }
  await insertViaPicker(page, SMALL_PNG);
  await images(page).first().click();
  const toolbar = (await page.getByTestId("image-toolbar").boundingBox())!;
  const image = (await images(page).first().boundingBox())!;
  const handle = (await page.getByTestId("image-resize-n").boundingBox())!;
  // The bar sits above the image with clear air between it and the handles.
  expect(toolbar.y + toolbar.height).toBeLessThan(handle.y);
  expect(toolbar.y + toolbar.height).toBeLessThanOrEqual(image.y - 16);
});

test("dragging a corner handle resizes proportionally (AC-10)", async ({ page }) => {
  await createDocument(page);
  await content(page).click();
  await insertViaPicker(page, SMALL_PNG);
  await images(page).first().click();
  const before = (await images(page).first().boundingBox())!;
  const ratio = before.width / before.height;
  const handle = page.getByTestId("image-resize-se");
  const box = (await handle.boundingBox())!;
  await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2);
  await page.mouse.down();
  await page.mouse.move(box.x + 60, box.y + 30, { steps: 6 });
  await page.mouse.up();
  const after = (await images(page).first().boundingBox())!;
  expect(after.width).toBeGreaterThan(before.width);
  expect(after.width / after.height).toBeGreaterThan(ratio - 0.15);
  expect(after.width / after.height).toBeLessThan(ratio + 0.15);
});

test("dragging an edge handle resizes only that dimension; Shift constrains it (AC-11, AC-12)", async ({ page }) => {
  await createDocument(page);
  await content(page).click();
  await insertViaPicker(page, SMALL_PNG);
  await images(page).first().click();
  const before = (await images(page).first().boundingBox())!;
  const eHandle = page.getByTestId("image-resize-e");
  const eBox = (await eHandle.boundingBox())!;
  await page.mouse.move(eBox.x + eBox.width / 2, eBox.y + eBox.height / 2);
  await page.mouse.down();
  await page.mouse.move(eBox.x + 50, eBox.y, { steps: 6 });
  await page.mouse.up();
  const afterEdge = (await images(page).first().boundingBox())!;
  expect(afterEdge.width).toBeGreaterThan(before.width);
  expect(afterEdge.height).toBeCloseTo(before.height, 0);

  // The image stays selected after its own resize, so the same handle can
  // be dragged again immediately — this time with Shift held, so the
  // height follows too (AC-12).
  const eHandle2 = page.getByTestId("image-resize-e");
  const eBox2 = (await eHandle2.boundingBox())!;
  await page.mouse.move(eBox2.x + eBox2.width / 2, eBox2.y + eBox2.height / 2);
  await page.mouse.down();
  await page.keyboard.down("Shift");
  await page.mouse.move(eBox2.x + 40, eBox2.y, { steps: 6 });
  await page.mouse.up();
  await page.keyboard.up("Shift");
  const afterShift = (await images(page).first().boundingBox())!;
  expect(afterShift.height).not.toBeCloseTo(afterEdge.height, 0);
});

// --- wrap modes (AC-13, AC-14) ----------------------------------------------

test("the toolbar's wrap modes change the image's rendering (AC-13, AC-14)", async ({ page }) => {
  await createDocument(page);
  await content(page).click();
  await insertViaPicker(page, SMALL_PNG);
  await images(page).first().click();
  await expect(page.getByTestId("image-wrap-inline")).toHaveAttribute("aria-pressed", "true");

  // The image stays selected across each of its own edits (the toolbar
  // persists), so the wrap buttons can be clicked back to back.
  await page.getByTestId("image-wrap-square").click();
  await expect(page.getByTestId("image-wrap-square")).toHaveAttribute("aria-pressed", "true");
  const floatValue = await images(page).first().evaluate((el) => getComputedStyle(el).float);
  expect(floatValue).toBe("left");

  await page.getByTestId("image-wrap-behind").click();
  await expect(page.getByTestId("image-wrap-behind")).toHaveAttribute("aria-pressed", "true");
  const behindPosition = await images(page).first().evaluate((el) => getComputedStyle(el).position);
  expect(behindPosition).toBe("absolute");

  await page.getByTestId("image-wrap-infront").click();
  await expect(page.getByTestId("image-wrap-infront")).toHaveAttribute("aria-pressed", "true");
});

// --- Image options panel (AC-15..22) ----------------------------------------

test("the Image options panel edits alt text and size, with lock-aspect-ratio and reset (AC-15..22)", async ({ page }) => {
  await createDocument(page);
  await content(page).click();
  await insertViaPicker(page, SMALL_PNG);
  await images(page).first().click();
  await page.getByTestId("image-options-button").click();
  const dialog = page.getByTestId("image-options-dialog");
  await expect(dialog).toBeVisible();

  await page.getByTestId("image-options-alt-text").fill("A test photo");
  const widthField = page.getByTestId("image-options-width");
  const heightField = page.getByTestId("image-options-height");
  await expect(page.getByTestId("image-options-lock-aspect")).toBeChecked();
  await widthField.fill("10");
  const heightAfterLockedEdit = await heightField.inputValue();
  expect(Number(heightAfterLockedEdit)).toBeCloseTo(5, 0); // 200x100 -> 2:1

  // Unlock: width/height become independent (AC-20).
  await page.getByTestId("image-options-lock-aspect").uncheck();
  await widthField.fill("12");
  await heightField.fill("3");
  await page.getByTestId("image-options-apply").click();
  await expect(dialog).not.toBeVisible();

  const img = images(page).first();
  await expect(img).toHaveAttribute("alt", "A test photo");
  const box = (await img.boundingBox())!;
  const cmPx = 96 / 2.54;
  expect(box.width).toBeGreaterThan(12 * cmPx * 0.9);
  expect(box.height).toBeGreaterThan(3 * cmPx * 0.9);
  expect(box.height).toBeLessThan(3 * cmPx * 1.1);

  // Reset size restores the natural size and re-enables the lock (AC-21, AC-22).
  await img.click();
  await page.getByTestId("image-options-button").click();
  await page.getByTestId("image-options-reset-size").click();
  await expect(page.getByTestId("image-options-lock-aspect")).toBeChecked();
  const resetWidth = await widthField.inputValue();
  expect(Number(resetWidth)).toBeCloseTo(200 / cmPx, 1);
  await page.getByTestId("image-options-apply").click();
});

// --- paste (AC-1) ------------------------------------------------------------

test("pasting an image from the clipboard inserts it at the caret (AC-1)", async ({ page }) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("keep");
  const bytes = Array.from(SMALL_PNG);
  await page.evaluate((pngBytes) => {
    const transfer = new DataTransfer();
    transfer.items.add(new File([new Uint8Array(pngBytes)], "shot.png", { type: "image/png" }));
    const target = document.querySelector('[data-testid="document-content"]');
    target?.dispatchEvent(
      new ClipboardEvent("paste", { clipboardData: transfer, bubbles: true, cancelable: true }),
    );
  }, bytes);
  await expect(images(page)).toHaveCount(1);
  await expect(content(page)).toContainText("keep");
});

// --- delete + reference-counted media (AC-30..32) ---------------------------

test("deleting a selected image removes it (AC-30)", async ({ page }) => {
  await createDocument(page);
  await content(page).click();
  await insertViaPicker(page, SMALL_PNG);
  await expect(images(page)).toHaveCount(1);
  await images(page).first().click();
  await page.keyboard.press("Delete");
  await expect(images(page)).toHaveCount(0);
});

// --- opening real fixtures (AC-28, AC-29) -----------------------------------

test("opens a .docx with one inline and one floating image at the source's own size (AC-28)", async ({ page }) => {
  await openFixture(page, docxImagesFixture);
  await expect(images(page)).toHaveCount(2);
  const first = (await images(page).nth(0).boundingBox())!;
  const second = (await images(page).nth(1).boundingBox())!;
  // Inline picture: 300x200px; floating picture: 120x120px — CSS "cm" is
  // itself 96/2.54 px, so the source's own pixel sizes are the expected
  // rendered CSS pixel widths directly.
  expect(first.width).toBeGreaterThan(300 * 0.9);
  expect(first.width).toBeLessThan(300 * 1.1);
  expect(second.width).toBeGreaterThan(120 * 0.9);
  expect(second.width).toBeLessThan(120 * 1.1);
});

test("opens a .odt with one inline and one floating image at the source's own size (AC-29)", async ({ page }) => {
  await openFixture(page, odtImagesFixture);
  await expect(images(page)).toHaveCount(2);
  const first = (await images(page).nth(0).boundingBox())!;
  const second = (await images(page).nth(1).boundingBox())!;
  expect(first.width).toBeGreaterThan(300 * 0.9);
  expect(first.width).toBeLessThan(300 * 1.1);
  expect(second.width).toBeGreaterThan(120 * 0.9);
  expect(second.width).toBeLessThan(120 * 1.1);
});

// --- round trip (AC-23..27) -------------------------------------------------

test("alt text and size survive a .docx save and re-open (AC-23, AC-25)", async ({ page }) => {
  await createDocument(page);
  await content(page).click();
  await insertViaPicker(page, SMALL_PNG);
  await images(page).first().click();
  await page.getByTestId("image-options-button").click();
  await page.getByTestId("image-options-alt-text").fill("Round trip photo");
  await page.getByTestId("image-options-apply").click();

  await page.getByTestId("menu-title-file").click();

  await page.getByTestId("file-menu-export").click();
  const [download] = await Promise.all([
    page.waitForEvent("download"),
    page.getByTestId("download-docx").click(),
  ]);
  const path = await download.path();
  expect(path).toBeTruthy();

  await openFixture(page, path!);
  await expect(images(page)).toHaveCount(1);
  await expect(images(page).first()).toHaveAttribute("alt", "Round trip photo");
});

// --- drag-and-drop (AC-1) ----------------------------------------------------

test("dropping an image file onto the canvas inserts it (AC-1)", async ({ page }) => {
  await createDocument(page);
  await content(page).click();
  const bytes = Array.from(SMALL_PNG);
  await page.evaluate(async (pngBytes) => {
    const dt = new DataTransfer();
    dt.items.add(new File([new Uint8Array(pngBytes)], "dropped.png", { type: "image/png" }));
    const target = document.querySelector('[data-testid="document-content"]')!;
    const rect = target.getBoundingClientRect();
    const opts = { bubbles: true, cancelable: true, dataTransfer: dt, clientX: rect.x + 20, clientY: rect.y + 20 };
    target.dispatchEvent(new DragEvent("dragenter", opts));
    target.dispatchEvent(new DragEvent("dragover", opts));
    target.dispatchEvent(new DragEvent("drop", opts));
  }, bytes);
  await expect(images(page)).toHaveCount(1);
});

// --- SVG handling: PNG for .docx, SVG for .odt (AC-2, AC-3) -----------------

const SIMPLE_SVG = Buffer.from(
  '<svg xmlns="http://www.w3.org/2000/svg" width="150" height="80"><rect width="150" height="80" fill="#ff6600"/></svg>',
);

test("an SVG inserted into a .docx-bound document is converted to PNG, not retained as SVG (AC-2)", async ({ page }) => {
  await createDocument(page);
  await content(page).click();
  await insertViaPicker(page, SIMPLE_SVG, "vector.svg", "image/svg+xml");
  await expect(images(page)).toHaveCount(1);

  await page.getByTestId("menu-title-file").click();

  await page.getByTestId("file-menu-export").click();
  const [download] = await Promise.all([
    page.waitForEvent("download"),
    page.getByTestId("download-docx").click(),
  ]);
  const bytes = readFileSync((await download.path())!);
  const parts = await unzip(bytes);
  const media = [...parts.keys()].filter((k) => k.startsWith("word/media/"));
  expect(media.some((f) => f.endsWith(".png"))).toBe(true);
  expect(media.some((f) => f.endsWith(".svg"))).toBe(false);
});

test("an SVG inserted into an .odt-bound document is stored as an SVG media part (AC-3)", async ({ page }) => {
  // Opening an existing .odt fixture puts the document in odt-bound mode,
  // so the insert path below stores the SVG as-is instead of converting it.
  await openFixture(page, odtImagesFixture);
  await content(page).click();
  await page.keyboard.press(`${MOD}+End`);
  await insertViaPicker(page, SIMPLE_SVG, "vector.svg", "image/svg+xml");
  await expect(images(page)).toHaveCount(3);

  await page.getByTestId("menu-title-file").click();

  await page.getByTestId("file-menu-export").click();
  const [download] = await Promise.all([
    page.waitForEvent("download"),
    page.getByTestId("download-odt").click(),
  ]);
  const bytes = readFileSync((await download.path())!);
  const parts = await unzip(bytes);
  const media = [...parts.keys()].filter((k) => k.startsWith("Pictures/"));
  expect(media.some((f) => f.toLowerCase().endsWith(".svg"))).toBe(true);
});

// --- alt text for .odt (AC-24) ----------------------------------------------

test("alt text is written to .odt as svg:desc (AC-24)", async ({ page }) => {
  await createDocument(page);
  await content(page).click();
  await insertViaPicker(page, SMALL_PNG);
  await images(page).first().click();
  await page.getByTestId("image-options-button").click();
  await page.getByTestId("image-options-alt-text").fill("ODT alt text");
  await page.getByTestId("image-options-apply").click();

  await page.getByTestId("menu-title-file").click();

  await page.getByTestId("file-menu-export").click();
  const [download] = await Promise.all([
    page.waitForEvent("download"),
    page.getByTestId("download-odt").click(),
  ]);
  const bytes = readFileSync((await download.path())!);
  const parts = await unzip(bytes);
  const contentXml = new TextDecoder().decode(parts.get("content.xml"));
  expect(contentXml).toContain("<svg:desc");
  expect(contentXml).toContain("ODT alt text");
});

// --- media parts and content types (AC-26, AC-27) ---------------------------

test("images are stored as real media parts with content types and relationships, never as base64 in the XML (AC-26, AC-27)", async ({ page }) => {
  await createDocument(page);
  await content(page).click();
  await insertViaPicker(page, SMALL_PNG);
  await page.getByTestId("menu-title-file").click();
  await page.getByTestId("file-menu-export").click();
  const [download] = await Promise.all([
    page.waitForEvent("download"),
    page.getByTestId("download-docx").click(),
  ]);
  const bytes = readFileSync((await download.path())!);
  const parts = await unzip(bytes);

  const media = [...parts.keys()].filter((k) => k.startsWith("word/media/"));
  expect(media.length).toBe(1);

  const contentTypesXml = new TextDecoder().decode(parts.get("[Content_Types].xml"));
  expect(contentTypesXml).toContain('Extension="png"');

  const relsXml = new TextDecoder().decode(parts.get("word/_rels/document.xml.rels"));
  expect(relsXml).toMatch(/Type=".*\/image"/);
  expect(relsXml).toContain("Target=\"media/");

  const docXml = new TextDecoder().decode(parts.get("word/document.xml"));
  expect(docXml).not.toContain("base64");
  expect(docXml).not.toContain("data:image");
});

// --- reference-counted media deletion (AC-31, AC-32) ------------------------

test("deleting one of two uses of the same image keeps its media part; deleting both removes it (AC-31, AC-32)", async ({ page }) => {
  // Two Image blocks sharing one mediaId is the "image used twice" case
  // (see lib/documentModel.ts's Image.mediaId doc comment). There is no
  // in-app way to duplicate an image onto a second block (copy/paste only
  // carries text runs — lib/clipboard.ts's extractFragment works over
  // Paragraph blocks), so real documents reach this state by having the
  // same relationship/media referenced twice already, which is what this
  // fixture (built through the app's own model + writer, the same
  // dedup path AC-2's docx save always exercises) reproduces.
  const model = createDocumentModel("dup-fixture");
  model.title = "Duplicate Image Fixture";
  const TINY_PNG_BASE64 = "iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAIAAAD91JpzAAAADElEQVR4nGP8z8DwHwAFBQIAX8jx0gAAAABJRU5ErkJggg==";
  model.media = [{ mediaId: "shared-media", contentType: "image/png", data: TINY_PNG_BASE64 }];
  const imageBlock = {
    type: "image" as const,
    mediaId: "shared-media",
    naturalWidthPx: 40,
    naturalHeightPx: 40,
    widthCm: 40 / (96 / 2.54),
    heightCm: 40 / (96 / 2.54),
    wrap: "inline" as const,
  };
  model.sections = [
    {
      blocks: [
        { type: "paragraph", runs: [{ text: "first use:", formatting: {} }] },
        { ...imageBlock },
        { type: "paragraph", runs: [{ text: "second use:", formatting: {} }] },
        { ...imageBlock },
        { type: "paragraph", runs: [] },
      ],
    },
  ];
  const { bytes } = await writeDocx(model);
  const fixturePath = join(tmpdir(), `spec024-dup-image-${Date.now()}.docx`);
  writeFileSync(fixturePath, bytes);

  await openFixture(page, fixturePath);
  await expect(images(page)).toHaveCount(2);

  // Delete one of the two uses — the media part must survive (AC-32).
  await images(page).first().click();
  await page.keyboard.press("Delete");
  await expect(images(page)).toHaveCount(1);

  await page.getByTestId("menu-title-file").click();

  await page.getByTestId("file-menu-export").click();
  const [download1] = await Promise.all([
    page.waitForEvent("download"),
    page.getByTestId("download-docx").click(),
  ]);
  const bytes1 = readFileSync((await download1.path())!);
  const parts1 = await unzip(bytes1);
  expect([...parts1.keys()].filter((k) => k.startsWith("word/media/"))).toHaveLength(1);

  // Delete the remaining use — now the media part must be gone (AC-31).
  await images(page).first().click();
  await page.keyboard.press("Delete");
  await expect(images(page)).toHaveCount(0);

  await page.getByTestId("menu-title-file").click();

  await page.getByTestId("file-menu-export").click();
  const [download2] = await Promise.all([
    page.waitForEvent("download"),
    page.getByTestId("download-docx").click(),
  ]);
  const bytes2 = readFileSync((await download2.path())!);
  const parts2 = await unzip(bytes2);
  expect([...parts2.keys()].filter((k) => k.startsWith("word/media/"))).toHaveLength(0);
});
