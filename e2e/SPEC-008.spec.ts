import { test, expect, type Download, type Page } from "@playwright/test";
import { readFileSync } from "node:fs";
import { inflateRawSync } from "node:zlib";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";

// SPEC-008 — round-trip preservation of unsupported content.
const here = dirname(fileURLToPath(import.meta.url));
const chartsFixture = join(
  here,
  "..",
  "src",
  "lib",
  "docx",
  "fixtures",
  "word-charts-and-controls.docx",
);

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

async function openChartsFixture(page: Page): Promise<void> {
  await page.goto("/");
  await page.getByTestId("open-document-input").setInputFiles(chartsFixture);
  await expect(page).toHaveURL(/\/doc\/.+/);
}

async function download(page: Page, format: "docx" | "odt"): Promise<Download> {
  await page.getByTestId("download-menu").click();
  const downloadPromise = page.waitForEvent("download");
  await page.getByTestId(`download-${format}`).click();
  return downloadPromise;
}

test("a same-format save preserves the chart, content control and custom XML part byte-identical (AC-1, AC-2, AC-5)", async ({
  page,
}) => {
  await openChartsFixture(page);
  const buffer = await downloadedBuffer(await download(page, "docx"));

  const saved = unzipParts(buffer);
  const original = unzipParts(readFileSync(chartsFixture));
  for (const name of [
    "word/charts/chart1.xml",
    "customXml/item1.xml",
    "customXml/itemProps1.xml",
    "word/styles.xml",
  ]) {
    expect(saved.get(name), `${name} should survive the save`).toBeDefined();
    expect(saved.get(name)!.equals(original.get(name)!), `${name} should be byte-identical`).toBe(
      true,
    );
  }

  // The unmodelled document.xml elements are re-emitted in place.
  // (The serialiser may add an xmlns declaration on the fragment root, so
  // match the element name without its closing bracket.)
  const documentXml = saved.get("word/document.xml")!.toString("utf8");
  expect(documentXml).toContain("<w:sdt");
  expect(documentXml).toContain("Controlled block text");
  expect(documentXml).toContain("<w:drawing");

  // No conversion report for a same-format save: nothing was dropped.
  await expect(page.getByTestId("conversion-report")).toHaveCount(0);
});

test("a cross-format save shows the conversion report naming each dropped item and its location (AC-9, AC-10, AC-11, AC-12)", async ({
  page,
}) => {
  await openChartsFixture(page);
  await download(page, "odt");

  const report = page.getByTestId("conversion-report");
  await expect(report).toBeVisible();
  const items = page.getByTestId("conversion-report-item");
  await expect(items.filter({ hasText: "Chart" }).filter({ hasText: "paragraph 3" })).toHaveCount(
    1,
  );
  await expect(items.filter({ hasText: "Content control" })).toHaveCount(1);
  await expect(
    items.filter({ hasText: "Custom XML part" }).filter({ hasText: "customXml/item1.xml" }),
  ).toHaveCount(1);

  // Shown once per cross-format save: closing it and saving again brings it
  // back (the checkbox was not ticked).
  await page.getByTestId("conversion-report-close").click();
  await expect(report).toHaveCount(0);
  await download(page, "odt");
  await expect(report).toBeVisible();
});

test("'Don't show again for this document' suppresses future reports for that document (AC-13)", async ({
  page,
}) => {
  await openChartsFixture(page);
  await download(page, "odt");

  await expect(page.getByTestId("conversion-report")).toBeVisible();
  await page.getByTestId("conversion-report-dont-show").check();
  await page.getByTestId("conversion-report-close").click();
  await expect(page.getByTestId("conversion-report")).toHaveCount(0);

  await download(page, "odt");
  await expect(page.getByTestId("conversion-report")).toHaveCount(0);
});
