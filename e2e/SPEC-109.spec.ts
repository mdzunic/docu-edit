import { test, expect, type Locator, type Page } from "@playwright/test";
import { readFileSync } from "node:fs";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";

// SPEC-109 — combine images into one image, end to end (AC-52): open the
// library with spec107-shot.png (64 × 40) and spec107-photo.jpg (160 × 120)
// → select both → "Combine…" → grid, 2 columns, A4 landscape, Narrow margins
// → the preview reports two placed rectangles → Save to library → a new PNG
// row whose details panel reports the sheet's pixel size → back in the
// workspace (restored from the persisted plan), reorder the tray and save
// again → the second output differs from the first at the sampled centre
// pixels.
//
// Rendered BYTES are never compared (canvas encoders differ across
// platforms): the outputs are read back as pixels — their dimensions, the
// alpha of a background sample, and the colour at each placed item's centre.

const fixturesDir = join(dirname(fileURLToPath(import.meta.url)), "fixtures");

async function dropFile(page: Page, name: string, type: string) {
  const transfer = await page.evaluateHandle(
    ({ base64, fileName, mime }) => {
      const dt = new DataTransfer();
      const bytes = Uint8Array.from(atob(base64), (c) => c.charCodeAt(0));
      dt.items.add(new File([bytes], fileName, { type: mime }));
      return dt;
    },
    { base64: readFileSync(join(fixturesDir, name)).toString("base64"), fileName: name, mime: type },
  );
  await page.dispatchEvent("body", "drop", { dataTransfer: transfer });
  await expect(page).toHaveURL(/\/media\/.+/);
  await page.getByTestId("close-document").click();
  await expect(page.getByTestId("library")).toBeVisible();
}

function rowByTitle(page: Page, title: string) {
  return page
    .getByTestId("library-row")
    .filter({ has: page.getByTestId("row-title").getByText(title, { exact: true }) });
}

interface PlacedRect {
  xCm: number;
  yCm: number;
  widthCm: number;
  heightCm: number;
}

/** The preview's ARIA list of placed rectangles, read off its entries. */
async function placedRects(page: Page): Promise<PlacedRect[]> {
  return page.getByTestId("assembly-placed-item").evaluateAll((items) =>
    items.map((item) => ({
      xCm: Number(item.getAttribute("data-x-cm")),
      yCm: Number(item.getAttribute("data-y-cm")),
      widthCm: Number(item.getAttribute("data-width-cm")),
      heightCm: Number(item.getAttribute("data-height-cm")),
    })),
  );
}

/** One pixel of an <img>, read back through a canvas. */
async function samplePixels(image: Locator, points: Array<[number, number]>): Promise<number[][]> {
  await expect.poll(() => image.evaluate((img: HTMLImageElement) => (img.complete ? img.naturalWidth : 0))).toBeGreaterThan(0);
  return image.evaluate((img: HTMLImageElement, at) => {
    const canvas = document.createElement("canvas");
    canvas.width = img.naturalWidth;
    canvas.height = img.naturalHeight;
    const ctx = canvas.getContext("2d")!;
    ctx.drawImage(img, 0, 0);
    return at.map(([x, y]) => Array.from(ctx.getImageData(x, y, 1, 1).data));
  }, points);
}

// A4 landscape at 150 dpi: 29.7 × 21 cm → 1754 × 1240 px.
const SHEET_PX: [number, number] = [1754, 1240];
const SHEET_CM: [number, number] = [29.7, 21];

/** Each placed rectangle's centre, in output pixels. */
function centres(rects: PlacedRect[]): Array<[number, number]> {
  return rects.map((rect) => [
    Math.round(((rect.xCm + rect.widthCm / 2) / SHEET_CM[0]) * SHEET_PX[0]),
    Math.round(((rect.yCm + rect.heightCm / 2) / SHEET_CM[1]) * SHEET_PX[1]),
  ]);
}

/** Opens a library row in the media viewer; checks it is a PNG of the
 * sheet's pixel size; samples the background corner and `points`. */
async function readOutput(page: Page, title: string, points: Array<[number, number]>): Promise<number[][]> {
  const row = rowByTitle(page, title);
  await expect(row.getByTestId("row-format")).toHaveText("PNG");
  await row.dblclick();
  await expect(page).toHaveURL(/\/media\/.+/);
  await expect(page.getByTestId("detail-dimensions")).toHaveText(`${SHEET_PX[0]} × ${SHEET_PX[1]} px`);
  const samples = await samplePixels(page.getByTestId("media-image"), [[3, 3], ...points]);
  await page.getByTestId("close-document").click();
  await expect(page.getByTestId("library")).toBeVisible();
  return samples;
}

/** RGB within `tolerance` of `expected`. */
function expectNear(actual: number[], expected: number[], tolerance = 12) {
  for (let i = 0; i < 3; i += 1) {
    expect(Math.abs(actual[i] - expected[i]), `channel ${i}: ${actual.join(",")} vs ${expected.join(",")}`).toBeLessThanOrEqual(tolerance);
  }
}

test("combine two images in a 2-column grid on A4 landscape, save, reorder, save again (AC-52)", async ({ page }) => {
  await page.goto("/");
  await expect(page.getByTestId("app-bar")).toBeVisible();
  await dropFile(page, "spec107-shot.png", "image/png");
  await dropFile(page, "spec107-photo.jpg", "image/jpeg");

  // --- select both → Combine… ------------------------------------------------
  const combine = page.getByTestId("bulk-combine");
  await rowByTitle(page, "spec107-shot").getByTestId("row-title").click({ modifiers: ["ControlOrMeta"] });
  await expect(combine).toBeDisabled();
  await rowByTitle(page, "spec107-photo").getByTestId("row-title").click({ modifiers: ["ControlOrMeta"] });
  await expect(page.getByTestId("selection-count")).toHaveText("2 selected");
  await expect(combine).toBeEnabled();
  await combine.click();
  await expect(page).toHaveURL(/\/assemble$/);

  // The tray: the selection, in order, with thumbnails, titles and sizes.
  const items = page.getByTestId("assembly-tray").getByTestId("assembly-item");
  await expect(items).toHaveCount(2);
  await expect(items.getByTestId("item-title")).toHaveText(["spec107-shot", "spec107-photo"]);
  await expect(items.getByTestId("item-size")).toHaveText(["64 × 40 px", "160 × 120 px"]);
  await expect(items.first().getByTestId("row-thumbnail")).toBeVisible();

  // --- grid, 2 columns, A4 landscape, Narrow ----------------------------------
  await page.getByTestId("assembly-layout").selectOption("grid");
  await page.getByTestId("assembly-columns").selectOption("2");
  await page.getByTestId("sheet-paper").selectOption("A4");
  await page.getByTestId("sheet-landscape").click();
  await page.getByTestId("sheet-margin-preset").selectOption("Narrow");
  await expect(page.getByTestId("sheet-margin-left")).toHaveValue("1.27");

  // Two placed rectangles: the 27.16 × 18.46 cm content box splits into two
  // 13.48 cm cells around the 0.2 cm gap; each image is contained in its own.
  const placed = page.getByTestId("assembly-placed-item");
  await expect(placed).toHaveCount(2);
  await expect(page.getByTestId("assembly-preview")).toBeVisible();
  const first = await placedRects(page);
  expect(first).toEqual([
    { xCm: 1.27, yCm: 6.288, widthCm: 13.48, heightCm: 8.425 },
    { xCm: 14.95, yCm: 5.445, widthCm: 13.48, heightCm: 10.11 },
  ]);
  await expect(page.getByTestId("assembly-estimate")).toContainText(`${SHEET_PX[0]} × ${SHEET_PX[1]} px at 150 dpi`);
  await expect(page.getByTestId("assembly-format")).toHaveValue("png");
  await expect(page.getByTestId("assembly-quality")).toHaveCount(0);
  await expect(page.getByTestId("assembly-title")).toHaveValue("Combined image");

  // --- Save to library ------------------------------------------------------
  await page.getByTestId("assembly-save").click();
  await expect(page.getByTestId("assembly-saved")).toHaveText("Saved “Combined image” to the library.");
  await page.getByTestId("close-document").click();
  await expect(page.getByTestId("library")).toBeVisible();
  // The sources are untouched.
  await expect(rowByTitle(page, "spec107-shot").getByTestId("row-format")).toHaveText("PNG");
  await expect(rowByTitle(page, "spec107-photo").getByTestId("row-format")).toHaveText("JPEG");

  const points = centres(first);
  const [background, ...before] = await readOutput(page, "Combined image", points);
  // The PNG keeps the sheet's transparent background.
  expect(background[3]).toBe(0);
  // Cell 1's centre is the screenshot's centre: its gradient's middle.
  expectNear(before[0], [128, 128, 96]);
  expect(before[0][3]).toBe(255);
  expect(before[1][3]).toBe(255);

  // --- back to the workspace: the plan was persisted --------------------------
  await page.goto("/assemble");
  await expect(items).toHaveCount(2);
  await expect(items.getByTestId("item-title")).toHaveText(["spec107-shot", "spec107-photo"]);
  await expect(page.getByTestId("assembly-columns")).toHaveValue("2");
  await expect(page.getByTestId("sheet-landscape")).toHaveAttribute("aria-pressed", "true");
  await expect(page.getByTestId("sheet-margin-preset")).toHaveValue("Narrow");

  // Reorder by keyboard: Alt+↓ on the first row, announced.
  await items.first().focus();
  await page.keyboard.press("Alt+ArrowDown");
  await expect(page.getByTestId("assembly-tray-status")).toHaveText("“spec107-shot” moved to position 2 of 2.");
  await expect(items.getByTestId("item-title")).toHaveText(["spec107-photo", "spec107-shot"]);
  await expect(items.nth(1)).toBeFocused();
  await expect(placed).toHaveCount(2);

  await page.getByTestId("assembly-save").click();
  await expect(page.getByTestId("assembly-saved")).toHaveText("Saved “Combined image (2)” to the library.");
  await page.getByTestId("close-document").click();
  await expect(page.getByTestId("library")).toBeVisible();

  const [, ...after] = await readOutput(page, "Combined image (2)", points);
  // The same centres now hold the other image: the outputs differ there.
  for (let i = 0; i < 2; i += 1) {
    const difference = [0, 1, 2].reduce((sum, c) => sum + Math.abs(after[i][c] - before[i][c]), 0);
    expect(difference, `centre ${i}: ${before[i].join(",")} → ${after[i].join(",")}`).toBeGreaterThan(30);
  }
  // And what moved into cell 2 is the screenshot's centre.
  expectNear(after[1], [128, 128, 96]);
});

test("'New combined image' opens an empty tray, and 'Add from library…' offers only images", async ({ page }) => {
  await page.goto("/");
  await expect(page.getByTestId("app-bar")).toBeVisible();
  await dropFile(page, "spec107-shot.png", "image/png");

  await page.getByTestId("new-combined-image").click();
  await expect(page).toHaveURL(/\/assemble$/);
  await expect(page.getByTestId("assembly-tray").getByTestId("assembly-item")).toHaveCount(0);
  await expect(page.getByTestId("assembly-tray-empty")).toBeVisible();
  await expect(page.getByTestId("assembly-save")).toBeDisabled();

  await page.getByTestId("assembly-add").click();
  const picker = page.getByTestId("assembly-picker");
  await expect(picker).toBeVisible();
  await expect(picker.getByTestId("picker-item")).toHaveCount(1);
  await picker.getByTestId("picker-check").check();
  await picker.getByTestId("assembly-picker-add").click();
  await expect(picker).toBeHidden();
  await expect(page.getByTestId("assembly-tray").getByTestId("item-title")).toHaveText(["spec107-shot"]);

  // Duplicate, then remove the copy.
  await page.getByTestId("item-duplicate").first().click();
  await expect(page.getByTestId("assembly-item")).toHaveCount(2);
  await expect(page.getByTestId("assembly-placed-item")).toHaveCount(2);
  await page.getByTestId("item-remove").nth(1).click();
  await expect(page.getByTestId("assembly-item")).toHaveCount(1);

  // JPEG: quality appears, and the transparent sheet is refused up front.
  await page.getByTestId("assembly-format").selectOption("jpeg");
  await expect(page.getByTestId("assembly-quality")).toHaveValue("90");
  await expect(page.getByTestId("assembly-background-notice")).toContainText("JPEG has no transparency");
  await page.getByTestId("assembly-format").selectOption("png");
  await expect(page.getByTestId("assembly-quality")).toHaveCount(0);
  await expect(page.getByTestId("assembly-background-notice")).toHaveCount(0);

  // Start over clears the tray.
  await page.getByTestId("assembly-reset").click();
  await expect(page.getByTestId("assembly-item")).toHaveCount(0);
  await page.reload();
  await expect(page.getByTestId("assembly-tray-empty")).toBeVisible();
});
