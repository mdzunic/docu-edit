import { test, expect, type Download, type Page } from "@playwright/test";
import { readFileSync } from "node:fs";
import { inflateRawSync } from "node:zlib";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";

// Presentations preparation — opening `.pptx`/`.odp` into the library, the
// read-only /slides/:id preview, kind-aware routing, and the SPEC-008-style
// round-trip promise: a download re-emits every source part byte-identical.
// The fixtures are generated from src/lib/presentation/fixtures/build.ts
// (scripts/build-presentation-fixtures.ts), so these assertions and the
// unit suite's describe the same decks.
const here = dirname(fileURLToPath(import.meta.url));
const pptxFixture = join(here, "fixtures", "mini-deck.pptx");
const odpFixture = join(here, "fixtures", "mini-deck.odp");

/** Minimal central-directory walk over a downloaded archive (same shape as
 * e2e/SPEC-006.spec.ts) — enough to compare preserved parts byte-for-byte
 * from Node. */
function unzipParts(buffer: Buffer): Map<string, Buffer> {
  let eocd = buffer.length - 22;
  while (eocd >= 0 && buffer.readUInt32LE(eocd) !== 0x06054b50) {
    eocd--;
  }
  expect(eocd, "end-of-central-directory record should exist").toBeGreaterThanOrEqual(0);
  const totalEntries = buffer.readUInt16LE(eocd + 10);
  let offset = buffer.readUInt32LE(eocd + 16);
  const parts = new Map<string, Buffer>();
  for (let i = 0; i < totalEntries; i++) {
    expect(buffer.readUInt32LE(offset), "central directory entry signature").toBe(0x02014b50);
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

async function downloadedBuffer(download: Download): Promise<Buffer> {
  const path = await download.path();
  expect(path).toBeTruthy();
  return readFileSync(path);
}

async function openFixture(page: Page, fixture: string): Promise<void> {
  await page.goto("/");
  await page.getByTestId("open-document-input").setInputFiles(fixture);
  await expect(page).toHaveURL(/\/slides\/.+/);
}

function expectPartsEqual(written: Map<string, Buffer>, source: Map<string, Buffer>): void {
  expect([...written.keys()]).toEqual([...source.keys()]);
  for (const [name, data] of source) {
    expect(written.get(name)!.equals(data), `part ${name} must round-trip byte-identical`).toBe(true);
  }
}

test("opening a .pptx lands on the /slides stage with the deck's slides", async ({ page }) => {
  await openFixture(page, pptxFixture);

  await expect(page.getByTestId("presentation-view")).toBeVisible();
  await expect(page.getByTestId("document-title")).toHaveText("mini-deck");
  await expect(page.getByTestId("presentation-slide-count")).toHaveText("4 slides");
  // The filmstrip lists every slide; the stage renders the first (SPEC-079).
  await expect(page.getByTestId("filmstrip-slide")).toHaveCount(4);
  await expect(page.getByTestId("stage-workspace").getByText("Quarterly Review")).toBeVisible();
  // Grouped shapes' text renders when navigated to (the reader walks p:grpSp).
  await page.getByTestId("filmstrip-slide").nth(2).click();
  await expect(page.getByTestId("stage-workspace").getByText("Grouped note")).toBeVisible();
  // The read-only preview note is gone.
  await expect(page.getByTestId("presentation-note")).toHaveCount(0);
});

test("the library lists a presentation with its badge and routes it back to /slides", async ({ page }) => {
  await openFixture(page, pptxFixture);
  const slidesUrl = page.url();

  await page.getByTestId("presentation-close").click();
  await expect(page.getByTestId("library")).toBeVisible();
  const row = page.getByTestId("library-row").filter({ hasText: "mini-deck" });
  await expect(row.getByTestId("row-format")).toHaveText("PPTX");

  await row.getByTestId("row-title").click();
  await expect(page).toHaveURL(slidesUrl);
});

test("a presentation id navigated to /doc/:id redirects to /slides/:id (kind guard)", async ({ page }) => {
  await openFixture(page, pptxFixture);
  const id = page.url().split("/slides/")[1];

  await page.goto(`/doc/${id}`);
  await expect(page).toHaveURL(new RegExp(`/slides/${id}$`));
  await expect(page.getByTestId("presentation-view")).toBeVisible();
});

test("downloading a .pptx re-emits every source part byte-identical (SPEC-008's rule)", async ({ page }) => {
  await openFixture(page, pptxFixture);

  const downloadPromise = page.waitForEvent("download");
  await page.getByTestId("presentation-download").click();
  const download = await downloadPromise;
  expect(download.suggestedFilename()).toBe("mini-deck.pptx");

  const written = unzipParts(await downloadedBuffer(download));
  const source = unzipParts(readFileSync(pptxFixture));
  expectPartsEqual(written, source);
});

test("opening and downloading an .odp round-trips, mimetype first and STOREd", async ({ page }) => {
  await openFixture(page, odpFixture);
  await expect(page.getByTestId("presentation-slide-count")).toHaveText("4 slides");
  await expect(page.getByTestId("stage-workspace").getByText("Quarterly Review")).toBeVisible();

  const downloadPromise = page.waitForEvent("download");
  await page.getByTestId("presentation-download").click();
  const download = await downloadPromise;
  expect(download.suggestedFilename()).toBe("mini-deck.odp");

  const buffer = await downloadedBuffer(download);
  expectPartsEqual(unzipParts(buffer), unzipParts(readFileSync(odpFixture)));
  // ODF container rule: first local entry is `mimetype`, STOREd (method 0),
  // so sniffers can read it from the raw bytes.
  expect(buffer.readUInt16LE(8)).toBe(0);
  expect(buffer.toString("utf8", 30, 38)).toBe("mimetype");
});

test("a presentation shows in the grid with a slide count", async ({ page }) => {
  await openFixture(page, odpFixture);
  await page.getByTestId("presentation-close").click();

  await page.getByTestId("view-grid").click();
  const card = page.getByTestId("library-row").filter({ hasText: "mini-deck" });
  await expect(card.getByTestId("row-counts")).toHaveText(/4 slides/);
});
