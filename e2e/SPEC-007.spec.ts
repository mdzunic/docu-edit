import { test, expect, type Download } from "@playwright/test";
import { readFileSync } from "node:fs";
import { inflateRawSync } from "node:zlib";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";

// SPEC-007 — save as .odt.
const here = dirname(fileURLToPath(import.meta.url));
const docxFixture = join(here, "..", "src", "lib", "docx", "fixtures", "word-authored.docx");
const odtFixture = join(here, "..", "src", "lib", "odf", "fixtures", "libreoffice-writer.odt");

interface ZipEntry {
  data: Buffer;
  method: number;
}

/** Minimal central-directory walk over a downloaded archive, recording each
 * entry's compression method and central-directory order — enough to assert
 * the ODF package structure from Node, where the app's DOMParser-based
 * parser can't run (the model-level round trips live in the vitest suite,
 * src/lib/odf/writer.spec.ts). */
function unzipEntries(buffer: Buffer): Map<string, ZipEntry> {
  let eocd = buffer.length - 22;
  while (eocd >= 0 && buffer.readUInt32LE(eocd) !== 0x06054b50) {
    eocd--;
  }
  expect(eocd, "end-of-central-directory record should exist").toBeGreaterThanOrEqual(0);
  const totalEntries = buffer.readUInt16LE(eocd + 10);
  let offset = buffer.readUInt32LE(eocd + 16);
  const entries = new Map<string, ZipEntry>();
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
    entries.set(name, {
      data: method === 8 ? inflateRawSync(data) : Buffer.from(data),
      method,
    });
    offset += 46 + nameLength + extraLength + commentLength;
  }
  return entries;
}

async function downloadedBuffer(download: Download): Promise<Buffer> {
  const path = await download.path();
  expect(path).toBeTruthy();
  return readFileSync(path);
}

test("the app bar offers Download > OpenDocument Text (.odt) and downloads <title>.odt (AC-1, AC-2)", async ({
  page,
}) => {
  await page.goto("/");
  await page.getByTestId("open-document-input").setInputFiles(odtFixture);
  await expect(page).toHaveURL(/\/doc\/.+/);

  await page.getByTestId("menu-title-file").click();

  await page.getByTestId("file-menu-export").click();
  const option = page.getByTestId("download-odt");
  await expect(option).toHaveText("OpenDocument Text (.odt)");

  const downloadPromise = page.waitForEvent("download");
  await option.click();
  const download = await downloadPromise;
  // The fixture opens titled after its file name, so the download round-trips it.
  expect(download.suggestedFilename()).toBe("libreoffice-writer.odt");
});

test("the downloaded package is a valid ODF ZIP: stored mimetype first, all required parts (AC-3..AC-10)", async ({
  page,
}) => {
  await page.goto("/");
  await page.getByTestId("open-document-input").setInputFiles(odtFixture);
  await expect(page).toHaveURL(/\/doc\/.+/);
  await page.getByTestId("menu-title-file").click();
  await page.getByTestId("file-menu-export").click();
  const downloadPromise = page.waitForEvent("download");
  await page.getByTestId("download-odt").click();
  const buffer = await downloadedBuffer(await downloadPromise);

  const entries = unzipEntries(buffer);
  for (const required of [
    "mimetype",
    "content.xml",
    "styles.xml",
    "meta.xml",
    "settings.xml",
    "META-INF/manifest.xml",
  ]) {
    expect(Array.from(entries.keys()), `package should contain ${required}`).toContain(required);
  }

  // The mimetype entry is written first (AC-3), with ZIP STORE (AC-10): the
  // first local file header sits at offset 0 and its raw bytes are the exact
  // media type string (AC-9).
  expect(buffer.readUInt32LE(0)).toBe(0x04034b50);
  const firstNameLength = buffer.readUInt16LE(26);
  expect(buffer.toString("utf8", 30, 30 + firstNameLength)).toBe("mimetype");
  expect(entries.get("mimetype")!.method).toBe(0);
  expect(entries.get("mimetype")!.data.toString("utf8")).toBe(
    "application/vnd.oasis.opendocument.text",
  );

  // No lost text: every paragraph of the source survives (AC-11).
  const contentXml = entries.get("content.xml")!.data.toString("utf8");
  expect(contentXml).toContain("Project Plan");
  expect(contentXml).toContain("Schedule");
  expect(contentXml).toContain("Centered indented paragraph.");

  // The title and a dc:date stamp reach meta.xml (AC-19, AC-20).
  const metaXml = entries.get("meta.xml")!.data.toString("utf8");
  expect(metaXml).toContain("<dc:title>libreoffice-writer</dc:title>");
  expect(metaXml).toMatch(/<dc:date>[^<]+<\/dc:date>/);
});

test("a document opened from .docx downloads as converted ODF, not renamed OPC (AC-24)", async ({
  page,
}) => {
  await page.goto("/");
  await page.getByTestId("open-document-input").setInputFiles(docxFixture);
  await expect(page).toHaveURL(/\/doc\/.+/);
  await expect(page.getByTestId("format-badge")).toHaveText("DOCX");

  await page.getByTestId("menu-title-file").click();

  await page.getByTestId("file-menu-export").click();
  const downloadPromise = page.waitForEvent("download");
  await page.getByTestId("download-odt").click();
  const download = await downloadPromise;
  expect(download.suggestedFilename()).toBe("word-authored.odt");

  const entries = unzipEntries(await downloadedBuffer(download));
  expect(Array.from(entries.keys())).toContain("content.xml");
  // A renamed OPC package would still carry these.
  expect(Array.from(entries.keys())).not.toContain("[Content_Types].xml");
  expect(Array.from(entries.keys())).not.toContain("word/document.xml");

  const contentXml = entries.get("content.xml")!.data.toString("utf8");
  expect(contentXml).toContain("urn:oasis:names:tc:opendocument");
  expect(contentXml).not.toContain("schemas.openxmlformats.org");
  expect(contentXml).toContain("Quarterly Report");
  expect(contentXml).toContain("Section Two");
});

test("downloading a blank document produces a valid one-paragraph file (AC-2)", async ({
  page,
}) => {
  await page.goto("/");
  await page.getByTestId("new-document").first().click();
  await expect(page).toHaveURL(/\/doc\/.+/);

  await page.getByTestId("menu-title-file").click();

  await page.getByTestId("file-menu-export").click();
  const downloadPromise = page.waitForEvent("download");
  await page.getByTestId("download-odt").click();
  const download = await downloadPromise;
  expect(download.suggestedFilename()).toBe("Untitled document.odt");

  const entries = unzipEntries(await downloadedBuffer(download));
  const contentXml = entries.get("content.xml")!.data.toString("utf8");
  expect(contentXml).toMatch(/<text:p\/>|<text:p>/);
  expect(entries.get("styles.xml")!.data.toString("utf8")).toContain("<style:page-layout");
});
