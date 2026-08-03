import { test, expect, type Download } from "@playwright/test";
import { readFileSync } from "node:fs";
import { inflateRawSync } from "node:zlib";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";

// SPEC-006 — save as .docx.
const here = dirname(fileURLToPath(import.meta.url));
const docxFixture = join(here, "..", "src", "lib", "docx", "fixtures", "word-authored.docx");
const odtFixture = join(here, "..", "src", "lib", "odf", "fixtures", "libreoffice-writer.odt");

/** Minimal central-directory walk over a downloaded archive — enough to
 * assert the OPC package structure from Node, where the app's DOMParser-based
 * validator can't run (the full schema validation lives in the vitest suite,
 * src/lib/docx/validate.spec.ts). */
function unzipParts(buffer: Buffer): Map<string, Buffer> {
  let eocd = buffer.length - 22;
  while (eocd >= 0 && buffer.readUInt32LE(eocd) !== 0x06054b50) {
    eocd--;
  }
  expect(eocd, "end-of-central-directory record should exist (AC-3)").toBeGreaterThanOrEqual(0);
  const totalEntries = buffer.readUInt16LE(eocd + 10);
  let offset = buffer.readUInt32LE(eocd + 16);
  const parts = new Map<string, Buffer>();
  for (let i = 0; i < totalEntries; i++) {
    expect(buffer.readUInt32LE(offset), "central directory entry signature (AC-3)").toBe(
      0x02014b50,
    );
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

test("the app bar offers Download > Word (.docx) and downloads <title>.docx (AC-1, AC-2)", async ({
  page,
}) => {
  await page.goto("/");
  await page.getByTestId("open-document-input").setInputFiles(docxFixture);
  await expect(page).toHaveURL(/\/doc\/.+/);

  const downloadButton = page.getByTestId("download-menu");
  await expect(downloadButton).toBeVisible();
  await expect(page.getByTestId("app-bar").getByTestId("download-menu")).toBeVisible();
  await downloadButton.click();
  const option = page.getByTestId("download-docx");
  await expect(option).toHaveText("Word (.docx)");

  const downloadPromise = page.waitForEvent("download");
  await option.click();
  const download = await downloadPromise;
  // The fixture opens titled after its file name, so the download round-trips it.
  expect(download.suggestedFilename()).toBe("word-authored.docx");
});

test("the downloaded package is a valid OPC ZIP with all required parts (AC-3..AC-8)", async ({
  page,
}) => {
  await page.goto("/");
  await page.getByTestId("open-document-input").setInputFiles(docxFixture);
  await expect(page).toHaveURL(/\/doc\/.+/);
  await page.getByTestId("download-menu").click();
  const downloadPromise = page.waitForEvent("download");
  await page.getByTestId("download-docx").click();
  const buffer = await downloadedBuffer(await downloadPromise);

  const parts = unzipParts(buffer);
  for (const required of [
    "[Content_Types].xml",
    "_rels/.rels",
    "word/document.xml",
    "word/styles.xml",
    "docProps/core.xml",
  ]) {
    expect(Array.from(parts.keys()), `package should contain ${required}`).toContain(required);
  }

  // No lost text: every paragraph of the source survives (AC-9, AC-16, AC-17).
  const documentXml = parts.get("word/document.xml")!.toString("utf8");
  expect(documentXml).toContain("Quarterly Report");
  expect(documentXml).toContain("Section Two");
  expect(documentXml).toContain("Centered paragraph with custom spacing and indent.");

  // The title and a dcterms:modified stamp reach core.xml (AC-14, AC-15).
  const coreXml = parts.get("docProps/core.xml")!.toString("utf8");
  expect(coreXml).toContain("<dc:title>word-authored</dc:title>");
  expect(coreXml).toMatch(/<dcterms:modified xsi:type="dcterms:W3CDTF">[^<]+<\/dcterms:modified>/);
});

test("a document opened from .odt downloads as converted WordprocessingML, not renamed ODF (AC-23)", async ({
  page,
}) => {
  await page.goto("/");
  await page.getByTestId("open-document-input").setInputFiles(odtFixture);
  await expect(page).toHaveURL(/\/doc\/.+/);
  await expect(page.getByTestId("format-badge")).toHaveText("ODT");

  await page.getByTestId("download-menu").click();
  const downloadPromise = page.waitForEvent("download");
  await page.getByTestId("download-docx").click();
  const download = await downloadPromise;
  expect(download.suggestedFilename()).toBe("libreoffice-writer.docx");

  const parts = unzipParts(await downloadedBuffer(download));
  expect(Array.from(parts.keys())).toContain("word/document.xml");
  // A renamed ODF package would still carry these.
  expect(Array.from(parts.keys())).not.toContain("mimetype");
  expect(Array.from(parts.keys())).not.toContain("content.xml");

  const documentXml = parts.get("word/document.xml")!.toString("utf8");
  expect(documentXml).toContain("schemas.openxmlformats.org/wordprocessingml");
  expect(documentXml).not.toContain("urn:oasis");
  expect(documentXml).toContain("Project Plan");
  expect(documentXml).toContain("Schedule");
});

test("downloading a blank document produces a valid one-paragraph file (AC-24)", async ({
  page,
}) => {
  await page.goto("/");
  await page.getByTestId("new-document").first().click();
  await expect(page).toHaveURL(/\/doc\/.+/);

  await page.getByTestId("download-menu").click();
  const downloadPromise = page.waitForEvent("download");
  await page.getByTestId("download-docx").click();
  const download = await downloadPromise;
  expect(download.suggestedFilename()).toBe("Untitled document.docx");

  const parts = unzipParts(await downloadedBuffer(download));
  const documentXml = parts.get("word/document.xml")!.toString("utf8");
  expect(documentXml).toMatch(/<w:p\/>|<w:p>/);
  expect(documentXml).toContain("<w:sectPr>");
});
