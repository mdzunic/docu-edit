import { test, expect, type Download, type Locator, type Page } from "@playwright/test";
import { inflateRawSync } from "node:zlib";
import { readFileSync } from "node:fs";

// SPEC-027 — the Page setup dialog: paper size (standard + Custom),
// orientation, the four margins, the document-level measurement-unit
// preference, live re-render/reflow on OK, margin validation, .docx/.odt
// round trips of page geometry, accurate text-area rendering, and the
// Show/hide ruler toggle with margin markers.

async function createDocument(page: Page): Promise<string> {
  await page.goto("/");
  await page.getByTestId("new-document").click();
  await expect(page).toHaveURL(/\/doc\/.+/);
  const url = page.url();
  return url.slice(url.lastIndexOf("/") + 1);
}

async function openFixture(page: Page, path: string): Promise<string> {
  await page.goto("/");
  await page.getByTestId("open-document-input").setInputFiles(path);
  await expect(page).toHaveURL(/\/doc\/.+/);
  const url = page.url();
  return url.slice(url.lastIndexOf("/") + 1);
}

async function download(page: Page, format: "docx" | "odt"): Promise<string> {
  await page.getByTestId("download-menu").click();
  const file: Download = await Promise.all([
    page.waitForEvent("download"),
    page.getByTestId(`download-${format}`).click(),
  ]).then(([d]) => d);
  const path = await file.path();
  expect(path).toBeTruthy();
  return path!;
}

const content = (page: Page) => page.getByTestId("document-content");

async function openDialog(page: Page): Promise<void> {
  await page.getByTestId("menu-title-file").click();
  await page.getByTestId("page-setup-button").click();
  await expect(page.getByTestId("page-setup-dialog")).toBeVisible();
}

async function apply(page: Page): Promise<void> {
  await page.getByTestId("page-setup-apply").click();
  await expect(page.getByTestId("page-setup-dialog")).toHaveCount(0);
}

/** The cm value of an inline style property set directly (not computed) —
 * `.page`/`.page-margin` are sized/positioned in real "cm" CSS units from
 * the model, so reading the specified value (not getComputedStyle, which
 * would resolve to px) round-trips the exact number the app wrote. */
async function styleCm(locator: Locator, prop: string): Promise<number> {
  const raw = await locator.evaluate(
    (el, p) => (el as HTMLElement).style.getPropertyValue(p),
    prop,
  );
  return Number.parseFloat(raw);
}

const page_ = (page: Page) => page.getByTestId("document-page");

/** SPEC-028 moved the page margins onto the page content box's padding
 * (the page-margin element now marks the resulting content area), so the
 * specified margin value lives on the parent's padding-<side>. */
async function marginStyleCm(page: Page, side: string): Promise<number> {
  const raw = await pageMargin(page).evaluate(
    (el, s) => ((el as HTMLElement).parentElement as HTMLElement).style.getPropertyValue(`padding-${s}`),
    side,
  );
  return Number.parseFloat(raw);
}
const pageMargin = (page: Page) => page.getByTestId("page-margin");

/** Minimal central-directory walk over a downloaded archive — same helper
 * as e2e/SPEC-023.spec.ts. */
function unzipParts(buffer: Buffer): Map<string, Buffer> {
  let eocd = buffer.length - 22;
  while (eocd >= 0 && buffer.readUInt32LE(eocd) !== 0x06054b50) {
    eocd--;
  }
  expect(eocd).toBeGreaterThanOrEqual(0);
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

test("the dialog offers standard paper sizes plus Custom, orientation and four margins (AC-1, AC-2, AC-3)", async ({
  page,
}) => {
  await createDocument(page);
  await openDialog(page);
  const sizes = await page
    .getByTestId("page-setup-paper-size")
    .locator("option")
    .allTextContents();
  expect(sizes).toEqual(["A4", "Letter", "Legal", "A3", "A5", "Tabloid", "Custom"]);
  await expect(page.getByTestId("page-setup-orientation")).toBeVisible();
  for (const id of ["page-setup-margin-top", "page-setup-margin-right", "page-setup-margin-bottom", "page-setup-margin-left"]) {
    await expect(page.getByTestId(id)).toBeVisible();
  }
});

test("choosing Legal landscape re-renders the page immediately on OK (AC-1, AC-2, AC-5)", async ({ page }) => {
  await createDocument(page);
  await openDialog(page);
  await page.getByTestId("page-setup-paper-size").selectOption("Legal");
  await page.getByTestId("page-setup-orientation").selectOption("landscape");
  await apply(page);

  // Legal is 21.59 x 35.56cm portrait; landscape swaps them — and the
  // change is visible with no reload (AC-5's "immediately").
  expect(await styleCm(page_(page), "width")).toBeCloseTo(35.56, 1);
  expect(await styleCm(page_(page), "height")).toBeCloseTo(21.59, 1);
});

test("Custom size accepts width/height with a cm/mm/in/pt unit selector (AC-4)", async ({ page }) => {
  await createDocument(page);
  await openDialog(page);
  await page.getByTestId("page-setup-paper-size").selectOption("Custom");
  const units = await page.getByTestId("page-setup-unit").locator("option").allTextContents();
  expect(units).toEqual(["cm", "mm", "in", "pt"]);

  await page.getByTestId("page-setup-unit").selectOption("in");
  await page.getByTestId("page-setup-width").fill("8.5");
  await page.getByTestId("page-setup-height").fill("11");
  await apply(page);

  expect(await styleCm(page_(page), "width")).toBeCloseTo(21.59, 1);
  expect(await styleCm(page_(page), "height")).toBeCloseTo(27.94, 1);
});

test("OK sets the four margins, which render exactly and the text wraps inside them (AC-3, AC-5, AC-12, AC-13)", async ({
  page,
}) => {
  await createDocument(page);
  await openDialog(page);
  // Inch-based values so twip-grid snapping (see commands.ts's setPageSetup)
  // does not perturb the numbers this test checks.
  await page.getByTestId("page-setup-margin-top").fill("2.54");
  await page.getByTestId("page-setup-margin-right").fill("5.08");
  await page.getByTestId("page-setup-margin-bottom").fill("2.54");
  await page.getByTestId("page-setup-margin-left").fill("1.27");
  await apply(page);

  expect(await marginStyleCm(page, "top")).toBeCloseTo(2.54, 2);
  expect(await marginStyleCm(page, "right")).toBeCloseTo(5.08, 2);
  expect(await marginStyleCm(page, "bottom")).toBeCloseTo(2.54, 2);
  expect(await marginStyleCm(page, "left")).toBeCloseTo(1.27, 2);

  // Content starts exactly at the top margin: the first paragraph's top
  // edge coincides with the margin box's top edge (AC-12; the editable
  // root itself wraps whole pages since SPEC-028).
  const marginBox = (await pageMargin(page).boundingBox())!;
  const contentBox = (await page.locator('[data-section="0"][data-block="0"]').first().boundingBox())!;
  expect(Math.abs(contentBox.y - marginBox.y)).toBeLessThanOrEqual(1);
  expect(Math.abs(contentBox.x - marginBox.x)).toBeLessThanOrEqual(1);

  // Text wraps at the right margin: a long unbroken line still fits inside
  // the margin box's width, never overflowing it (AC-13).
  await content(page).click();
  await page.keyboard.type(
    "one two three four five six seven eight nine ten eleven twelve thirteen fourteen fifteen sixteen",
  );
  const lineCount = await content(page).evaluate((el) => {
    const range = document.createRange();
    range.selectNodeContents(el);
    return range.getClientRects().length;
  });
  expect(lineCount).toBeGreaterThan(1);
  // The wrapped paragraph never overflows the margin box (the editable
  // root itself wraps whole pages since SPEC-028, so measure the block).
  const textBox = (await page.locator('[data-section="0"][data-block="0"]').first().boundingBox())!;
  expect(textBox.width).toBeLessThanOrEqual(marginBox.width + 1);
});

test("the dialog's unit follows the document preference, which is settable and persists across reload (AC-6, AC-7)", async ({
  page,
}) => {
  await createDocument(page);
  await openDialog(page);
  expect(await page.getByTestId("page-setup-unit").inputValue()).toBe("cm");
  await page.getByTestId("page-setup-unit").selectOption("in");
  await apply(page);

  // Re-opening reflects the just-set preference (AC-6).
  await openDialog(page);
  expect(await page.getByTestId("page-setup-unit").inputValue()).toBe("in");
  await page.getByTestId("page-setup-cancel").click();

  // A real reload reads the preference back from storage (AC-7).
  await page.reload();
  await openDialog(page);
  expect(await page.getByTestId("page-setup-unit").inputValue()).toBe("in");
});

test("values from the opened document populate the dialog, including a non-standard custom size (AC-8)", async ({
  page,
}) => {
  await createDocument(page);
  await openDialog(page);
  await page.getByTestId("page-setup-paper-size").selectOption("Custom");
  await page.getByTestId("page-setup-width").fill("12");
  await page.getByTestId("page-setup-height").fill("18");
  await page.getByTestId("page-setup-margin-top").fill("1");
  await page.getByTestId("page-setup-margin-right").fill("1");
  await page.getByTestId("page-setup-margin-bottom").fill("1");
  await page.getByTestId("page-setup-margin-left").fill("1");
  await apply(page);

  // Reload so the dialog is re-populated from the persisted document, not
  // from whatever the dialog happened to have in memory a moment ago.
  await page.reload();
  await openDialog(page);
  expect(await page.getByTestId("page-setup-paper-size").inputValue()).toBe("Custom");
  expect(Number(await page.getByTestId("page-setup-width").inputValue())).toBeCloseTo(12, 1);
  expect(Number(await page.getByTestId("page-setup-height").inputValue())).toBeCloseTo(18, 1);
});

test("margins that leave under 1cm of text area are rejected with the exact message (AC-9)", async ({ page }) => {
  await createDocument(page);
  await openDialog(page);
  await page.getByTestId("page-setup-margin-left").fill("10.5");
  await page.getByTestId("page-setup-margin-right").fill("10");
  await expect(page.getByTestId("page-setup-margins-error")).toHaveText(
    "Margins are too large for this page size",
  );
  await expect(page.getByTestId("page-setup-apply")).toBeDisabled();
  // The rejected values never reach the model — the page keeps its old
  // (valid) margins.
  await page.getByTestId("page-setup-cancel").click();
  expect(await marginStyleCm(page, "left")).toBeCloseTo(2.54, 1);
});

test("page setup round-trips through .docx (w:pgSz, w:pgMar) with no unit drift (AC-10)", async ({ page }) => {
  await createDocument(page);
  await openDialog(page);
  await page.getByTestId("page-setup-paper-size").selectOption("Custom");
  await page.getByTestId("page-setup-width").fill("20.32"); // 8in
  await page.getByTestId("page-setup-height").fill("25.4"); // 10in
  await page.getByTestId("page-setup-margin-top").fill("2.54");
  await page.getByTestId("page-setup-margin-right").fill("2.54");
  await page.getByTestId("page-setup-margin-bottom").fill("2.54");
  await page.getByTestId("page-setup-margin-left").fill("2.54");
  await apply(page);

  const path = await download(page, "docx");
  const parts = unzipParts(readFileSync(path));
  const documentXml = parts.get("word/document.xml")!.toString("utf8");
  // 8in/10in/1in are all whole numbers of twips (1440/in) — exact.
  expect(documentXml).toContain('<w:pgSz w:w="11520" w:h="14400"/>');
  expect(documentXml).toContain('w:top="1440"');
  expect(documentXml).toContain('w:left="1440"');

  await openFixture(page, path);
  expect(await styleCm(page_(page), "width")).toBeCloseTo(20.32, 2);
  expect(await styleCm(page_(page), "height")).toBeCloseTo(25.4, 2);
  expect(await marginStyleCm(page, "top")).toBeCloseTo(2.54, 2);
});

test("page setup round-trips through .odt (style:page-layout-properties) with no unit drift (AC-11)", async ({
  page,
}) => {
  await createDocument(page);
  await openDialog(page);
  await page.getByTestId("page-setup-paper-size").selectOption("A5");
  await page.getByTestId("page-setup-margin-top").fill("1.5");
  await page.getByTestId("page-setup-margin-right").fill("1.5");
  await page.getByTestId("page-setup-margin-bottom").fill("1.5");
  await page.getByTestId("page-setup-margin-left").fill("1.5");
  await apply(page);

  const path = await download(page, "odt");
  const parts = unzipParts(readFileSync(path));
  const stylesXml = parts.get("styles.xml")!.toString("utf8");
  expect(stylesXml).toContain("style:page-layout-properties");
  // A5's 14.8/21cm are themselves snapped onto the twip grid on entry (like
  // the .docx path — see commands.ts's setPageSetup), so the written length
  // is not the idealized decimal; check the numeric value instead.
  const width = Number(/fo:page-width="([\d.]+)cm"/.exec(stylesXml)?.[1]);
  const height = Number(/fo:page-height="([\d.]+)cm"/.exec(stylesXml)?.[1]);
  expect(width).toBeCloseTo(14.8, 2);
  expect(height).toBeCloseTo(21, 2);

  await openFixture(page, path);
  expect(await styleCm(page_(page), "width")).toBeCloseTo(14.8, 2);
  expect(await styleCm(page_(page), "height")).toBeCloseTo(21, 2);
  expect(await marginStyleCm(page, "top")).toBeCloseTo(1.5, 2);
});

test("the Show/hide ruler toggle displays and hides the ruler (AC-14)", async ({ page }) => {
  await createDocument(page);
  await expect(page.getByTestId("document-ruler")).toBeVisible();
  await page.getByTestId("menu-title-view").click();
  await page.getByTestId("toggle-ruler-button").click();
  await expect(page.getByTestId("document-ruler")).toHaveCount(0);
  await page.getByTestId("menu-title-view").click();
  await page.getByTestId("toggle-ruler-button").click();
  await expect(page.getByTestId("document-ruler")).toBeVisible();
});

test("the ruler's margin markers reflect the current page setup (AC-15)", async ({ page }) => {
  await createDocument(page);
  await openDialog(page);
  await page.getByTestId("page-setup-margin-left").fill("5.08"); // 2in
  await page.getByTestId("page-setup-margin-right").fill("5.08");
  await apply(page);

  const ruler = page.getByTestId("document-ruler");
  await ruler.scrollIntoViewIfNeeded();
  const rulerBox = (await ruler.boundingBox())!;
  const pageWidthCm = await styleCm(page_(page), "width");
  const pxPerCm = rulerBox.width / pageWidthCm;

  const leftMarker = (await page.getByTestId("ruler-marker-margin-left").boundingBox())!;
  const rightMarker = (await page.getByTestId("ruler-marker-margin-right").boundingBox())!;
  const leftCm = (leftMarker.x + leftMarker.width / 2 - rulerBox.x) / pxPerCm;
  const rightCm = (rightMarker.x + rightMarker.width / 2 - rulerBox.x) / pxPerCm;
  expect(leftCm).toBeCloseTo(5.08, 0);
  expect(pageWidthCm - rightCm).toBeCloseTo(5.08, 0);
});
