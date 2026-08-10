import { test, expect, type Page } from "@playwright/test";
import { inflateRawSync } from "node:zlib";
import { readFileSync, writeFileSync } from "node:fs";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import { tmpdir } from "node:os";

// SPEC-042 — the custom styles manager: a "Styles" side pane lists every
// style grouped by type with in-use tracking (AC-1, AC-2); New/Modify/
// Delete manage the catalogue (AC-3..AC-8, AC-11); character styles apply
// independently of the paragraph style and both show at the caret (AC-9,
// AC-10); a colliding name is rejected inline (AC-7 in the intake's own
// numbering — "name collision" AC); custom styles round-trip through
// .docx/.odt (AC-13, AC-14); and "Import styles from another document"
// merges a second file's catalogue with per-conflict resolution (AC-19..24).

async function createDocument(page: Page): Promise<void> {
  await page.goto("/");
  await page.getByTestId("new-document").first().click();
  await expect(page).toHaveURL(/\/doc\/.+/);
}

const editor = (page: Page) => page.locator('[contenteditable="true"]').first();

async function openStylesPane(page: Page): Promise<void> {
  await page.getByTestId("menu-title-view").click();
  await page.getByTestId("view-menu-styles-pane").click();
  await expect(page.getByTestId("styles-pane")).toBeVisible();
}

async function newStyle(
  page: Page,
  type: "paragraph" | "character" | "list" | "table",
  name: string,
  options: { basedOn?: string; bold?: boolean; italic?: boolean; color?: string } = {},
): Promise<void> {
  await page.getByTestId(`new-style-${type}`).click();
  await expect(page.getByTestId("style-editor-dialog")).toBeVisible();
  await page.getByTestId("style-name-input").fill(name);
  if (options.basedOn) {
    await page.getByTestId("style-based-on-select").selectOption({ label: options.basedOn });
  }
  if (options.bold) {
    await page.getByTestId("style-bold-checkbox").check();
  }
  if (options.italic) {
    await page.getByTestId("style-italic-checkbox").check();
  }
  if (options.color) {
    await page.getByTestId("style-color-input").fill(options.color);
  }
  await page.getByTestId("style-editor-save").click();
  await expect(page.getByTestId("style-editor-dialog")).not.toBeVisible();
}

function styleRow(page: Page, name: string) {
  return page.locator(`[data-testid^="style-row-"]:has-text("${name}")`);
}

// ---- Minimal ZIP helpers, only used by the AC-15..18 unmodeled-style test ----
// Read a package's parts (the same central-directory walk SPEC-006/019 use)...
function unzipParts(buffer: Buffer): Array<[string, Buffer]> {
  let eocd = buffer.length - 22;
  while (eocd >= 0 && buffer.readUInt32LE(eocd) !== 0x06054b50) eocd--;
  const total = buffer.readUInt16LE(eocd + 10);
  let off = buffer.readUInt32LE(eocd + 16);
  const parts: Array<[string, Buffer]> = [];
  for (let i = 0; i < total; i++) {
    const method = buffer.readUInt16LE(off + 10);
    const csize = buffer.readUInt32LE(off + 20);
    const nlen = buffer.readUInt16LE(off + 28);
    const elen = buffer.readUInt16LE(off + 30);
    const clen = buffer.readUInt16LE(off + 32);
    const lo = buffer.readUInt32LE(off + 42);
    const name = buffer.toString("utf8", off + 46, off + 46 + nlen);
    const lnlen = buffer.readUInt16LE(lo + 26);
    const lelen = buffer.readUInt16LE(lo + 28);
    const ds = lo + 30 + lnlen + lelen;
    const data = buffer.subarray(ds, ds + csize);
    parts.push([name, method === 8 ? inflateRawSync(data) : Buffer.from(data)]);
    off += 46 + nlen + elen + clen;
  }
  return parts;
}

const CRC_TABLE = (() => {
  const t = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    t[n] = c >>> 0;
  }
  return t;
})();
function crc32(buf: Buffer): number {
  let c = 0xffffffff;
  for (let i = 0; i < buf.length; i++) c = CRC_TABLE[(c ^ buf[i]) & 0xff] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}

// ...and repackage them uncompressed (COMPRESSION_STORE, which the app's OPC
// reader accepts) so we can hand-edit styles.xml/document.xml without a zip
// dependency (package.json is protected — no new deps for test scaffolding).
function storedZip(parts: Array<[string, Buffer]>): Buffer {
  const locals: Buffer[] = [];
  const centrals: Buffer[] = [];
  let offset = 0;
  for (const [name, data] of parts) {
    const nameBuf = Buffer.from(name, "utf8");
    const crc = crc32(data);
    const lh = Buffer.alloc(30);
    lh.writeUInt32LE(0x04034b50, 0);
    lh.writeUInt16LE(20, 4);
    lh.writeUInt32LE(crc, 14);
    lh.writeUInt32LE(data.length, 18);
    lh.writeUInt32LE(data.length, 22);
    lh.writeUInt16LE(nameBuf.length, 26);
    const local = Buffer.concat([lh, nameBuf, data]);
    locals.push(local);
    const ch = Buffer.alloc(46);
    ch.writeUInt32LE(0x02014b50, 0);
    ch.writeUInt16LE(20, 4);
    ch.writeUInt16LE(20, 6);
    ch.writeUInt32LE(crc, 16);
    ch.writeUInt32LE(data.length, 20);
    ch.writeUInt32LE(data.length, 24);
    ch.writeUInt16LE(nameBuf.length, 28);
    ch.writeUInt32LE(offset, 42);
    centrals.push(Buffer.concat([ch, nameBuf]));
    offset += local.length;
  }
  const localBlob = Buffer.concat(locals);
  const centralBlob = Buffer.concat(centrals);
  const eocd = Buffer.alloc(22);
  eocd.writeUInt32LE(0x06054b50, 0);
  eocd.writeUInt16LE(parts.length, 8);
  eocd.writeUInt16LE(parts.length, 10);
  eocd.writeUInt32LE(centralBlob.length, 12);
  eocd.writeUInt32LE(localBlob.length, 16);
  return Buffer.concat([localBlob, centralBlob, eocd]);
}

// Build a foreign .docx from the table fixture, injecting a table style the
// editor cannot fully model ("Foreign Grid"), applied to the table by
// w:tblStyle. Returns the path to the written file.
function buildForeignDocx(): string {
  const fixturesDir = join(dirname(fileURLToPath(import.meta.url)), "fixtures");
  const parts = unzipParts(readFileSync(join(fixturesDir, "table.docx")));
  const tableStyle =
    '<w:style w:type="table" w:styleId="ForeignGrid"><w:name w:val="Foreign Grid"/>' +
    '<w:basedOn w:val="Normal"/><w:tblPr><w:tblBorders>' +
    '<w:top w:val="single" w:sz="6" w:space="0" w:color="FF00FF"/>' +
    "</w:tblBorders></w:tblPr></w:style>";
  const out = parts.map(([name, data]): [string, Buffer] => {
    if (name === "word/styles.xml") {
      const xml = data.toString("utf8").replace("</w:styles>", tableStyle + "</w:styles>");
      return [name, Buffer.from(xml, "utf8")];
    }
    if (name === "word/document.xml") {
      const xml = data.toString("utf8").replace("<w:tblPr>", '<w:tblPr><w:tblStyle w:val="ForeignGrid"/>');
      return [name, Buffer.from(xml, "utf8")];
    }
    return [name, data];
  });
  const path = join(tmpdir(), "spec042-foreign-styles.docx");
  writeFileSync(path, storedZip(out));
  return path;
}

test.describe("Styles pane (AC-1, AC-2)", () => {
  test("lists styles grouped by type, and marks an applied style in use", async ({ page }) => {
    await createDocument(page);
    await editor(page).click();
    await page.keyboard.type("Some text");
    await openStylesPane(page);

    await expect(page.getByTestId("styles-group-paragraph")).toBeVisible();
    await expect(page.getByTestId("styles-group-character")).toBeVisible();
    await expect(page.getByTestId("styles-group-list")).toBeVisible();
    await expect(page.getByTestId("styles-group-table")).toBeVisible();
    // Every built-in paragraph style is listed even before anything custom
    // is created.
    await expect(styleRow(page, "Heading 1")).toBeVisible();

    await newStyle(page, "paragraph", "My Body");
    await expect(styleRow(page, "My Body").getByTestId(/style-in-use-/)).toBeVisible();
  });
});

test.describe("New / Modify style (AC-3, AC-4, AC-5, AC-11)", () => {
  test("modifying a parent style re-bakes a child style's formatting without breaking its identity", async ({
    page,
  }) => {
    await createDocument(page);
    await editor(page).click();
    await page.keyboard.type("Parent paragraph.");
    await page.keyboard.press("Enter");
    await page.keyboard.type("Child paragraph.");
    await openStylesPane(page);

    await page.keyboard.press("Control+Home");
    await editor(page).click();
    await page.keyboard.press("Control+Home");
    await newStyle(page, "paragraph", "Parent Style", { bold: true });

    await page.keyboard.press("Control+End");
    await newStyle(page, "paragraph", "Child Style", { basedOn: "Parent Style", italic: true });

    // Both styles are in use before the modify.
    await expect(styleRow(page, "Parent Style").getByTestId(/style-in-use-/)).toBeVisible();
    await expect(styleRow(page, "Child Style").getByTestId(/style-in-use-/)).toBeVisible();

    await styleRow(page, "Parent Style").getByTestId(/style-modify-/).click();
    await expect(page.getByTestId("style-editor-dialog")).toBeVisible();
    await page.getByTestId("style-color-input").fill("#0000ff");
    await page.getByTestId("style-editor-save").click();
    await expect(page.getByTestId("style-editor-dialog")).not.toBeVisible();

    // The regression under test: Child Style must still be in use — its
    // paragraph must not have been silently re-pointed at Parent Style.
    await expect(styleRow(page, "Child Style").getByTestId(/style-in-use-/)).toBeVisible();
    await expect(styleRow(page, "Parent Style").getByTestId(/style-in-use-/)).toBeVisible();

    // The caret's readout on the child paragraph still names Child Style.
    await editor(page).click();
    await page.keyboard.press("Control+End");
    await expect(page.getByTestId("caret-paragraph-style")).toHaveText("Child Style");
  });

  test("rejects a name that collides with an existing style, inline", async ({ page }) => {
    await createDocument(page);
    await editor(page).click();
    await page.keyboard.type("Text");
    await openStylesPane(page);

    await page.getByTestId("new-style-paragraph").click();
    await page.getByTestId("style-name-input").fill("Normal");
    await expect(page.getByTestId("style-name-error")).toBeVisible();
    await expect(page.getByTestId("style-editor-save")).toBeDisabled();
    await page.getByTestId("style-editor-cancel").click();
  });
});

test.describe("Delete style (AC-7, AC-8)", () => {
  test("asks for confirmation and reverts paragraphs to Normal", async ({ page }) => {
    await createDocument(page);
    await editor(page).click();
    await page.keyboard.type("Styled paragraph.");
    await openStylesPane(page);

    await page.keyboard.press("Control+Home");
    await newStyle(page, "paragraph", "Temp Style", { bold: true });

    await styleRow(page, "Temp Style").getByTestId(/style-delete-/).click();
    await expect(page.getByTestId("style-delete-dialog")).toBeVisible();
    await expect(page.getByTestId("style-delete-in-use-warning")).toBeVisible();
    await page.getByTestId("style-delete-confirm").click();
    await expect(page.getByTestId("style-delete-dialog")).not.toBeVisible();

    await expect(styleRow(page, "Temp Style")).toHaveCount(0);
    await expect(page.getByTestId("caret-paragraph-style")).toHaveText("Normal");
  });
});

test.describe("Character styles (AC-9, AC-10)", () => {
  test("applies independently of the paragraph style and both show at the caret", async ({ page }) => {
    await createDocument(page);
    await editor(page).click();
    await page.keyboard.type("A run of text.");
    await openStylesPane(page);

    await page.keyboard.press("Home");
    await page.keyboard.down("Shift");
    for (let i = 0; i < 6; i++) {
      await page.keyboard.press("ArrowRight");
    }
    await page.keyboard.up("Shift");

    await newStyle(page, "character", "Emphasis Style", { italic: true, color: "#cc0000" });
    await expect(styleRow(page, "Emphasis Style").getByTestId(/style-in-use-/)).toBeVisible();

    await page.keyboard.press("Home");
    await page.keyboard.press("ArrowRight");
    await expect(page.getByTestId("caret-character-style")).toHaveText("Emphasis Style");
    await expect(page.getByTestId("caret-paragraph-style")).toHaveText("Normal");

    await page.keyboard.press("End");
    await expect(page.getByTestId("caret-character-style")).toHaveText("None");
  });
});

test.describe("Custom style round trip (AC-13, AC-14)", () => {
  test("a custom paragraph style survives a .docx download and reopen", async ({ page }) => {
    await createDocument(page);
    await editor(page).click();
    await page.keyboard.type("Custom styled paragraph.");
    await openStylesPane(page);
    await page.keyboard.press("Control+Home");
    await newStyle(page, "paragraph", "RoundTrip Style", { bold: true });

    await page.getByTestId("download-menu").click();
    const [download] = await Promise.all([
      page.waitForEvent("download"),
      page.getByTestId("download-docx").click(),
    ]);
    const path = await download.path();
    expect(path).toBeTruthy();

    await page.goto("/");
    await page.getByTestId("open-document-input").setInputFiles(path!);
    await expect(page).toHaveURL(/\/doc\/.+/);
    await openStylesPane(page);
    await expect(styleRow(page, "RoundTrip Style")).toBeVisible();
    await page.keyboard.press("Control+Home");
    await expect(page.getByTestId("caret-paragraph-style")).toHaveText("RoundTrip Style");
  });
});

test.describe("Import styles from another document (AC-19..24)", () => {
  test("merges non-conflicting styles automatically and lets the user resolve a conflict", async ({ page }) => {
    // Source document: a custom style whose name collides with the target's
    // own "Shared Name", plus one that does not.
    await createDocument(page);
    await editor(page).click();
    await page.keyboard.type("Source doc text.");
    await openStylesPane(page);
    await page.keyboard.press("Control+Home");
    await newStyle(page, "paragraph", "Shared Name", { bold: true });
    await newStyle(page, "paragraph", "Imported Unique Style", { italic: true });

    await page.getByTestId("download-menu").click();
    const [sourceDownload] = await Promise.all([
      page.waitForEvent("download"),
      page.getByTestId("download-docx").click(),
    ]);
    const sourcePath = await sourceDownload.path();
    expect(sourcePath).toBeTruthy();

    // Target document: its own "Shared Name" with different formatting.
    await createDocument(page);
    await editor(page).click();
    await page.keyboard.type("Target doc text.");
    await openStylesPane(page);
    await page.keyboard.press("Control+Home");
    await newStyle(page, "paragraph", "Shared Name", { italic: true });

    await page.getByTestId("menu-title-format").click();
    await page.getByTestId("format-menu-import-styles").click();
    await expect(page.getByTestId("import-styles-dialog")).toBeVisible();
    await page.getByTestId("import-styles-file-input").setInputFiles(sourcePath!);
    await expect(page.getByTestId("import-styles-non-conflicting-count")).toContainText("1");
    const conflictRow = page.locator('[data-testid^="import-conflict-"]', { hasText: "Shared Name" });
    await expect(conflictRow).toBeVisible();
    await conflictRow.getByTestId(/-replace$/).check();

    await page.getByTestId("import-styles-import").click();
    // Both documents also carry the full built-in set (materialized into
    // the catalogue the moment any custom style is created) — those collide
    // by name too and default to "keep", hence 11 skipped alongside the one
    // real conflict this test resolves as "replace".
    await expect(page.getByTestId("import-styles-summary")).toContainText("Added 1, replaced 1, skipped 11");
    await page.getByTestId("import-styles-done").click();
    await expect(page.getByTestId("import-styles-dialog")).not.toBeVisible();

    await expect(styleRow(page, "Imported Unique Style")).toBeVisible();
  });
});

test.describe("Unmodeled styles (AC-15, AC-16, AC-17, AC-18)", () => {
  test("a foreign table style is listed, apply-able by reference, rendered, and preserved verbatim", async ({
    page,
  }) => {
    const foreign = buildForeignDocx();

    await page.goto("/");
    await page.getByTestId("open-document-input").setInputFiles(foreign);
    await expect(page).toHaveURL(/\/doc\/.+/);
    await openStylesPane(page);

    // AC-15: the style the editor cannot fully model is still listed — under
    // the Table group — and flagged as unmodeled.
    await expect(page.getByTestId("styles-group-table")).toContainText("Foreign Grid");
    await expect(page.getByTestId("style-row-ForeignGrid")).toBeVisible();
    await expect(page.getByTestId("style-unmodeled-ForeignGrid")).toBeVisible();

    // AC-16: the reference is a first-class citizen — it can be applied by
    // reference (apply affordance present); it cannot be edited as if modeled
    // (Modify is disabled), so its reference is preserved rather than rewritten.
    await expect(page.getByTestId("style-apply-ForeignGrid")).toBeVisible();
    await expect(page.getByTestId("style-modify-ForeignGrid")).toBeDisabled();

    // AC-17: text/table using the unmodeled style still renders (resolved via
    // the parent / document default) rather than disappearing.
    await expect(page.getByTestId("document-content").locator("table").first()).toBeVisible();

    // AC-18: on export the unmodeled style is written back unchanged — same id,
    // same visible name, same border colour it was read in with.
    await page.getByTestId("download-menu").click();
    const [dl] = await Promise.all([
      page.waitForEvent("download"),
      page.getByTestId("download-docx").click(),
    ]);
    const path = await dl.path();
    const parts = new Map(unzipParts(readFileSync(path!)));
    const stylesXml = parts.get("word/styles.xml")!.toString("utf8");
    expect(stylesXml).toContain('w:styleId="ForeignGrid"');
    expect(stylesXml).toContain('w:val="Foreign Grid"');
    expect(stylesXml).toContain("FF00FF");
  });
});
