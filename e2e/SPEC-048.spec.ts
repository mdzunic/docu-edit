import { test, expect, type Page } from "@playwright/test";
import { readFileSync } from "node:fs";
import { embedFontsIntoDocx } from "../src/lib/fonts/embedDocx";
import { unzip } from "../src/lib/core/zip";
import { createDocument as buildDocumentModel } from "../src/lib/documentModel";
import { writeDocx } from "../src/lib/docx/writer";
import { DOCUMENTED_DEFAULT_FAMILY, fontStack } from "../src/lib/fonts/substitution";

// SPEC-048 — font fallback & embedding. Substitution keeps the original family
// name in the model while rendering (and measuring, and exporting) on a bundled
// metric twin; the "Tools -> Fonts" panel discloses each family's status; opened
// embedded fonts are used and preserved; and "Embed fonts in this file" writes
// subset faces, states the size cost, and skips + names licence-restricted ones.
// The substitution maths, subsetting and extraction are proven exhaustively at
// the unit level (src/lib/fonts/*.spec.ts); this drives the user-visible
// behaviour end to end in a real browser.

const DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const carlito = new Uint8Array(readFileSync("public/fonts/Carlito-Regular.ttf"));
const wordAuthored = new Uint8Array(readFileSync("src/lib/docx/fixtures/word-authored.docx"));

/** Set the OS/2 fsType u16 to `value` (offset+8 of the OS/2 table) so a face
 * can be marked licence-restricted for the AC-17/18 skip path. */
function withFsType(src: Uint8Array, value: number): Uint8Array {
  const bytes = src.slice();
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const numTables = view.getUint16(4);
  for (let i = 0; i < numTables; i += 1) {
    const base = 12 + i * 16;
    const tag = String.fromCharCode(
      view.getUint8(base), view.getUint8(base + 1), view.getUint8(base + 2), view.getUint8(base + 3),
    );
    if (tag === "OS/2") {
      view.setUint16(view.getUint32(base + 8) + 8, value);
      return bytes;
    }
  }
  throw new Error("no OS/2 table in the fixture face");
}

/** A family that is genuinely missing on every host: nothing ships it, we do
 * not bundle it, and `substituteFor` is an exact-name table so it has no
 * metric twin. Spelled with a serif hint so `genericOf` classifies it serif
 * and the fallback contrast below is measured against the host's serif — on a
 * Linux container bare `sans-serif` can itself resolve to Liberation Sans,
 * which would make the contrast vacuous.
 *
 * This used to be "Georgia", with the comment "not installed on this host".
 * That is true of the CI container and false of macOS, where Georgia ships
 * with the OS — so the test failed for everyone running it locally, asserting
 * a fact about the machine rather than about the app. */
const MISSING_FAMILY = "Nonesuch Serif";

/** A .docx whose body text names `MISSING_FAMILY`. Built here rather than
 * chosen from the toolbar's font list, which only offers real families. */
async function missingFontDocx(): Promise<Buffer> {
  const model = buildDocumentModel("spec-048-missing-font");
  model.format = "docx";
  model.sections = [
    {
      blocks: [
        {
          type: "paragraph",
          runs: [{ text: "Text in a missing font", formatting: { fontFamily: MISSING_FAMILY } }],
        },
      ],
    },
  ];
  const { bytes } = await writeDocx(model);
  return Buffer.from(bytes);
}

/** word-authored.docx (uses Calibri) carrying an embedded Calibri face —
 * installable by default, or licence-restricted when `restricted`. */
async function embeddedCalibriDocx(restricted: boolean): Promise<Buffer> {
  const faceBytes = restricted ? withFsType(carlito, 0x0002) : carlito;
  const out = await embedFontsIntoDocx(wordAuthored, [
    { family: "Calibri", bold: false, italic: false, bytes: faceBytes },
  ]);
  return Buffer.from(out);
}

const fontsReady = (page: Page) =>
  page.evaluate(() => (document as unknown as { fonts: FontFaceSet }).fonts.ready);

async function newDocument(page: Page): Promise<void> {
  await page.goto("/");
  await page.getByTestId("new-document").click();
  await expect(page).toHaveURL(/\/doc\/.+/);
  await fontsReady(page);
}

async function openBuffer(page: Page, name: string, buffer: Buffer): Promise<void> {
  await page.goto("/");
  await page.getByTestId("open-document-input").setInputFiles({ name, mimeType: DOCX_MIME, buffer });
  await expect(page).toHaveURL(/\/doc\/.+/);
  await fontsReady(page);
}

async function openPath(page: Page, path: string): Promise<void> {
  await page.goto("/");
  await page.getByTestId("open-document-input").setInputFiles(path);
  await expect(page).toHaveURL(/\/doc\/.+/);
  await fontsReady(page);
}

const pane = (page: Page) => page.getByTestId("fonts-pane");

async function openFontsPane(page: Page): Promise<void> {
  await page.getByTestId("menu-title-tools").click();
  await page.getByTestId("tools-menu-fonts").click();
  await expect(pane(page)).toBeVisible();
}

async function download(page: Page, testid: string) {
  const [dl] = await Promise.all([
    page.waitForEvent("download"),
    page.getByTestId("download-menu").click().then(() => page.getByTestId(testid).click()),
  ]);
  return dl;
}

/** Canvas text width for a CSS font shorthand family list. */
async function widthOf(page: Page, family: string): Promise<number> {
  return page.evaluate((fam) => {
    const ctx = document.createElement("canvas").getContext("2d")!;
    ctx.font = `40px ${fam}`;
    return ctx.measureText("The quick brown fox jumps over 0123456789").width;
  }, family);
}

test("bundled metric twins are registered and loaded (AC-1)", async ({ page }) => {
  await newDocument(page);
  const loaded = await page.evaluate(() => {
    const fs = document as unknown as { fonts: FontFaceSet };
    return ["Carlito", "Caladea", "Liberation Sans", "Liberation Serif", "Liberation Mono"].map(
      (fam) => ({ fam, ok: fs.fonts.check(`16px "${fam}"`) }),
    );
  });
  for (const l of loaded) expect(l.ok, `${l.fam} loaded`).toBe(true);
});

test("an unavailable font is rendered on its metric twin (AC-2/6)", async ({ page }) => {
  await newDocument(page);
  const stack = await widthOf(page, `"Calibri", "Carlito", sans-serif`);
  const twin = await widthOf(page, `"Carlito", sans-serif`);
  const generic = await widthOf(page, `sans-serif`);
  const calibriOnly = await widthOf(page, `"Calibri", sans-serif`);
  // Calibri is unavailable here, so the app's stack paints on Carlito — the same
  // metric twin layout measurement and PDF export resolve to (AC-2, AC-6).
  expect(stack).toBeCloseTo(twin, 1);
  expect(Math.abs(stack - generic)).toBeGreaterThan(1);
  expect(calibriOnly).toBeCloseTo(generic, 1); // Calibri itself is absent
});

test("the original family name is kept in the model and on save; content is unchanged (AC-3/4/5)", async ({
  page,
}) => {
  await newDocument(page);
  await page.getByTestId("document-content").click();
  await page.keyboard.type("Round trip content");

  const saved = new Uint8Array(readFileSync((await download(page, "download-docx").then((d) => d.path()))!));
  const parts = await unzip(saved);
  const dec = new TextDecoder();
  const doc = dec.decode(parts.get("word/document.xml")!);
  const styles = parts.has("word/styles.xml") ? dec.decode(parts.get("word/styles.xml")!) : "";
  const all = doc + styles;
  expect(all).toContain("Calibri"); // AC-4: original name written on save
  expect(all).not.toContain("Carlito"); // the twin never leaks into the document
  expect(doc).toContain("Round trip content"); // AC-5: content preserved verbatim

  // AC-3: the model keeps the original name — the panel lists Calibri, never Carlito.
  await openFontsPane(page);
  await expect(page.getByTestId("fonts-item-Calibri")).toBeVisible();
  await expect(page.getByTestId("fonts-item-Carlito")).toHaveCount(0);
});

test("the Fonts panel lists document fonts with a status (AC-7/8/10)", async ({ page }) => {
  await newDocument(page);
  await page.getByTestId("document-content").click();
  await page.keyboard.type("Hello fonts");

  await openFontsPane(page);

  // The panel summarises how many fonts are in use and lists them.
  await expect(page.getByTestId("fonts-summary")).toContainText(/font/i);
  const items = pane(page).locator(".fonts-item");
  await expect(items.first()).toBeVisible();

  // The document default (Calibri) appears with a resolved status; on a host
  // without Calibri it renders on its bundled twin, Carlito.
  const calibriStatus = page.getByTestId("fonts-status-Calibri");
  await expect(calibriStatus).toBeVisible();
  await expect(calibriStatus).toHaveText(/Available|Substituted|Missing/);

  // Closing the panel hides it again.
  await page.getByTestId("fonts-pane-close").click();
  await expect(pane(page)).toBeHidden();
});

test("a genuinely missing font falls back to the documented default and reads Missing (AC-9/10)", async ({
  page,
}) => {
  await openBuffer(page, "missing-font.docx", await missingFontDocx());

  // AC-10: read the panel status BEFORE any canvas measurement — measuring a
  // family on a canvas first warms fontconfig alias resolution and can flip
  // the install probe, so the cold read is the faithful one.
  await openFontsPane(page);
  await expect(page.getByTestId(`fonts-status-${MISSING_FAMILY}`)).toHaveText("Missing");

  // AC-9: the app's fallback stack is the documented default (Liberation
  // Sans), measurably distinct from the bare generic. Measured through
  // `fontStack` itself, so this follows the stack the app actually renders
  // and measures with rather than a copy of it typed out here.
  const stack = await widthOf(page, fontStack(MISSING_FAMILY));
  const dflt = await widthOf(page, `"${DOCUMENTED_DEFAULT_FAMILY}", serif`);
  const bareSerif = await widthOf(page, `serif`);
  expect(stack).toBeCloseTo(dflt, 1);
  expect(Math.abs(stack - bareSerif)).toBeGreaterThan(1);
});

test("an embedded .docx font is used on render and preserved on a plain save (AC-11/12)", async ({
  page,
}) => {
  // Installable embedded Calibri → Available (only possible via the embedded
  // face; Calibri is neither bundled nor installed here).
  await openBuffer(page, "embedded.docx", await embeddedCalibriDocx(false));
  await openFontsPane(page);
  await expect(page.getByTestId("fonts-status-Calibri")).toHaveText("Available");

  // A licence-restricted embedded face is NOT registered → stays Substituted.
  await openBuffer(page, "restricted.docx", await embeddedCalibriDocx(true));
  await openFontsPane(page);
  await expect(page.getByTestId("fonts-status-Calibri")).toHaveText("Substituted — Carlito");

  // AC-12: a plain (non-embed) save preserves the embedded face — reopening the
  // saved file shows Calibri Available again.
  await openBuffer(page, "embedded.docx", await embeddedCalibriDocx(false));
  const saved = await download(page, "download-docx").then((d) => d.path());
  await openPath(page, saved!);
  await openFontsPane(page);
  await expect(page.getByTestId("fonts-status-Calibri")).toHaveText("Available");
});

test("an embedded .odt font is used on render and preserved on a plain save (AC-13/14)", async ({
  page,
}) => {
  await openPath(page, "src/lib/odf/fixtures/lo-embedded.odt");
  await openFontsPane(page);
  await expect(pane(page).locator(".fonts-item").first()).toBeVisible();
  const before = await pane(page).locator('[data-testid^="fonts-status-"]').allTextContents();
  expect(before.some((s) => /Available/.test(s))).toBe(true);

  // AC-14: a plain save preserves the embedded font — reopen still Available.
  const saved = await download(page, "download-odt").then((d) => d.path());
  await openPath(page, saved!);
  await openFontsPane(page);
  const after = await pane(page).locator('[data-testid^="fonts-status-"]').allTextContents();
  expect(after.some((s) => /Available/.test(s))).toBe(true);
});

test("embed-on-save produces a reopenable .odt and states its size (AC-15/16)", async ({ page }) => {
  const errors: string[] = [];
  page.on("console", (m) => {
    if (m.type() === "error") errors.push(m.text());
  });

  await openPath(page, "e2e/fixtures/styles.odt");
  await openFontsPane(page);
  await page.getByTestId("fonts-embed").click();
  await expect(page.getByTestId("embed-fonts-dialog")).toBeVisible();
  // AC-16: the size cost is stated before saving.
  await expect(page.getByTestId("embed-fonts-size")).toContainText(/adds about/i);

  // AC-15: confirm embed, capture the download, and reopen the app's OWN output.
  const [dl] = await Promise.all([
    page.waitForEvent("download"),
    page.getByTestId("embed-fonts-confirm").click(),
  ]);
  expect(dl.suggestedFilename()).toMatch(/\.odt$/);
  const saved = await dl.path();

  await page.goto("/");
  await page.getByTestId("open-document-input").setInputFiles(saved!);
  // Before the fix the app rejected its own output as ill-formed and never
  // reached /doc/. It must now open cleanly.
  await expect(page).toHaveURL(/\/doc\/.+/, { timeout: 15000 });
  await expect(page.getByText(/isn't an OpenDocument text file/i)).toHaveCount(0);

  await openFontsPane(page);
  await expect(page.getByTestId("fonts-summary")).toContainText(/font/i);
  expect(errors.filter((e) => /OpenDocument|Namespace|font-face/i.test(e))).toEqual([]);
});

test("embed-on-save skips and names licence-restricted fonts (AC-17/18)", async ({ page }) => {
  await openBuffer(page, "restricted.docx", await embeddedCalibriDocx(true));
  await openFontsPane(page);
  await page.getByTestId("fonts-embed").click();
  await expect(page.getByTestId("embed-fonts-dialog")).toBeVisible();
  // AC-17/18: the restricted face is skipped and named to the user.
  await expect(page.getByTestId("embed-fonts-skipped")).toContainText("Calibri");
});

test("fonts resolve without a flash of unstyled text (AC-19)", async ({ page }) => {
  await page.goto("/");
  await page.getByTestId("new-document").click();
  await expect(page).toHaveURL(/\/doc\/.+/);
  const ready = await page.evaluate(async () => {
    const fs = document as unknown as { fonts: FontFaceSet };
    await fs.fonts.ready;
    return {
      status: (fs.fonts as unknown as { status: string }).status,
      carlito: fs.fonts.check(`16px "Carlito"`),
    };
  });
  expect(ready.status).toBe("loaded");
  expect(ready.carlito).toBe(true);
});

test("PDF export embeds the same resolved twin the screen renders (AC-20)", async ({ page }) => {
  await newDocument(page);
  await page.getByTestId("document-content").click();
  await page.keyboard.type("Cambria run text");
  await page.keyboard.press("ControlOrMeta+a");
  await page.getByTestId("font-family-select").selectOption("Cambria");
  const pdf = readFileSync((await download(page, "download-pdf").then((d) => d.path()))!).toString("latin1");
  // Cambria resolves to its metric twin Caladea both on screen and in the PDF.
  expect(pdf).toMatch(/Caladea/);
});
