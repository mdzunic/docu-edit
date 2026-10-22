import { test, expect, type Locator, type Page } from "@playwright/test";
import { readFileSync } from "node:fs";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";

// SPEC-112 — edit an image, end to end (AC-44..48): open spec107-photo.jpg
// (160 × 120, four flat quadrants) → Edit → crop to 16:9 → rotate right →
// brightness +30 → Save → the library thumbnail updates and the details
// panel reports the new pixel size → reopen → the history lists three edits
// and the crop is still adjustable → Revert to original → the pixel size
// returns to the source's → redact a rectangle → the notice appears →
// Flatten and discard original → the history is empty, the size matches the
// redacted render, and Undo does not bring the original back.
//
// Rendered BYTES are never compared (canvas encoders differ across
// platforms); rendered PIXELS are, by sampling, with a tolerance for JPEG.

const fixturesDir = join(dirname(fileURLToPath(import.meta.url)), "fixtures");
const photo = join(fixturesDir, "spec107-photo.jpg");

const REDACTION_NOTICE = "The original is still stored. Choose Flatten and discard original to remove the hidden pixels.";

async function dropPhoto(page: Page) {
  const transfer = await page.evaluateHandle(
    (base64) => {
      const dt = new DataTransfer();
      const bytes = Uint8Array.from(atob(base64), (c) => c.charCodeAt(0));
      dt.items.add(new File([bytes], "spec107-photo.jpg", { type: "image/jpeg" }));
      return dt;
    },
    readFileSync(photo).toString("base64"),
  );
  await page.dispatchEvent("body", "drop", { dataTransfer: transfer });
  await expect(page).toHaveURL(/\/media\/.+/);
}

function rowByTitle(page: Page, title: string) {
  return page
    .getByTestId("library-row")
    .filter({ has: page.getByTestId("row-title").getByText(title, { exact: true }) });
}

/** An <img>'s natural size once it has loaded. */
async function naturalSize(image: Locator): Promise<[number, number]> {
  await expect
    .poll(() => image.evaluate((img: HTMLImageElement) => (img.complete ? img.naturalWidth : 0)))
    .toBeGreaterThan(0);
  return image.evaluate((img: HTMLImageElement) => [img.naturalWidth, img.naturalHeight] as [number, number]);
}

/** One pixel of an <img>, read back through a canvas. */
async function samplePixel(image: Locator, x: number, y: number): Promise<number[]> {
  await naturalSize(image);
  return image.evaluate(
    (img: HTMLImageElement, at) => {
      const canvas = document.createElement("canvas");
      canvas.width = img.naturalWidth;
      canvas.height = img.naturalHeight;
      const ctx = canvas.getContext("2d")!;
      ctx.drawImage(img, 0, 0);
      return Array.from(ctx.getImageData(at.x, at.y, 1, 1).data);
    },
    { x, y },
  );
}

/** RGB within `tolerance` of `expected` (JPEG round trips drift a little). */
function expectNear(actual: number[], expected: number[], tolerance = 14) {
  for (let i = 0; i < 3; i += 1) {
    expect(Math.abs(actual[i] - expected[i]), `channel ${i}: ${actual.join(",")} vs ${expected.join(",")}`).toBeLessThanOrEqual(tolerance);
  }
}

async function openEditor(page: Page) {
  await page.getByTestId("media-edit-open").click();
  await expect(page.getByTestId("media-edit")).toBeVisible();
  await expect(page.getByTestId("edit-preview")).toBeVisible();
}

test("crop 16:9, rotate, brighten, save, reopen, revert, redact, flatten (AC-44..48)", async ({ page }) => {
  await page.goto("/");
  await expect(page.getByTestId("app-bar")).toBeVisible();
  await dropPhoto(page);
  await expect(page.getByTestId("detail-dimensions")).toHaveText("160 × 120 px");

  // The library card's thumbnail before any edit: the photo's own shape.
  await page.getByTestId("close-document").click();
  await expect(page.getByTestId("library")).toBeVisible();
  await page.getByTestId("view-grid").click();
  const thumbnail = rowByTitle(page, "spec107-photo").getByTestId("media-thumbnail");
  await expect(thumbnail).toBeVisible();
  expect(await naturalSize(thumbnail)).toEqual([160, 120]);
  const thumbnailBefore = await thumbnail.getAttribute("src");

  // --- AC-44: Edit → crop 16:9 → rotate right → brightness +30 → Save ------
  await rowByTitle(page, "spec107-photo").dblclick();
  await expect(page).toHaveURL(/\/media\/.+/);
  await openEditor(page);

  await page.getByTestId("tool-crop").click();
  await page.getByTestId("crop-aspect").selectOption("16:9");
  await expect(page.getByTestId("edit-crop-announce")).toHaveText("0, 15, 160 × 90");

  await page.getByTestId("tool-rotate").click();
  await page.getByTestId("edit-rotate-right").click();

  await page.getByTestId("tool-adjust").click();
  await page.getByTestId("adjust-brightness-number").fill("30");
  await page.getByTestId("adjust-brightness-number").press("Enter");
  await expect(page.getByTestId("adjust-brightness")).toHaveValue("30");

  const rows = page.getByTestId("media-edit-history").getByTestId("edit-history-row");
  await expect(rows).toHaveCount(3);
  await expect(page.getByTestId("edit-undo")).toHaveAttribute("title", "Undo Adjust brightness");
  // During the session the preview is the current image.
  await expect(page.getByTestId("detail-dimensions")).toHaveText("90 × 160 px");

  await page.getByTestId("edit-save").click();
  await expect(page.getByTestId("edit-notice")).toHaveText("Saved.");
  await page.getByTestId("edit-done").click();
  await expect(page.getByTestId("media-edit")).toBeHidden();

  // The details panel reports the new pixel size, and the viewer shows the
  // render: sampled pixels, not bytes. Turned right, output (70, 40) reads
  // the photo's red top-left quadrant (200, 60, 60) and (20, 40) its blue
  // bottom-left (60, 80, 200), each brightened by 30 · 2.55 = 76.5.
  await expect(page.getByTestId("detail-dimensions")).toHaveText("90 × 160 px");
  await expect(page.getByTestId("detail-original-dimensions")).toContainText("160 × 120 px");
  const viewed = page.getByTestId("media-image");
  expect(await naturalSize(viewed)).toEqual([90, 160]);
  expectNear(await samplePixel(viewed, 70, 40), [255, 137, 137]);
  expectNear(await samplePixel(viewed, 20, 40), [137, 157, 255]);

  // The library thumbnail updates.
  await page.getByTestId("close-document").click();
  await expect(page.getByTestId("library")).toBeVisible();
  await expect(thumbnail).toBeVisible();
  await expect.poll(() => thumbnail.getAttribute("src")).not.toBe(thumbnailBefore);
  expect(await naturalSize(thumbnail)).toEqual([90, 160]);

  // --- AC-45: reopen → three edits in the history, the crop adjustable ----
  await rowByTitle(page, "spec107-photo").dblclick();
  await expect(page).toHaveURL(/\/media\/.+/);
  await expect(page.getByTestId("detail-dimensions")).toHaveText("90 × 160 px");
  await openEditor(page);
  await expect(rows).toHaveCount(3);
  await expect(rows.nth(0)).toContainText("Crop");
  await expect(rows.nth(1)).toContainText("Rotate right");
  await expect(rows.nth(2)).toContainText("Adjust brightness");

  await page.getByTestId("tool-crop").click();
  await expect(page.getByTestId("edit-crop-announce")).toHaveText("0, 15, 160 × 90");
  const bottomEdge = page.getByTestId("crop-handle-s");
  await bottomEdge.focus();
  await bottomEdge.press("ArrowUp");
  await expect(page.getByTestId("edit-crop-announce")).toHaveText("0, 15, 160 × 89");
  await bottomEdge.press("Shift+ArrowUp");
  await expect(page.getByTestId("edit-crop-announce")).toHaveText("0, 15, 160 × 79");
  // Each key press is one undo step.
  await page.keyboard.press("ControlOrMeta+z");
  await expect(page.getByTestId("edit-crop-announce")).toHaveText("0, 15, 160 × 89");
  await page.keyboard.press("ControlOrMeta+z");
  await expect(page.getByTestId("edit-crop-announce")).toHaveText("0, 15, 160 × 90");

  // --- AC-46: Revert to original → the source's pixel size ---------------
  await page.getByTestId("edit-revert").click();
  const revert = page.getByTestId("edit-revert-dialog");
  await expect(revert).toBeVisible();
  await expect(page.getByTestId("edit-revert-message")).toContainText("Crop, Rotate right, Adjust brightness");
  await page.getByTestId("edit-revert-confirm").click();
  await expect(revert).toBeHidden();
  await expect(rows).toHaveCount(0);
  await expect(page.getByTestId("detail-dimensions")).toHaveText("160 × 120 px");

  // --- AC-47: redact a rectangle → the notice ------------------------------
  await expect(page.getByTestId("edit-redaction-notice")).toHaveCount(0);
  await page.getByTestId("tool-redact").click();
  const surface = page.getByTestId("edit-draw-surface");
  const box = (await surface.boundingBox())!;
  await page.mouse.move(box.x + box.width * 0.1, box.y + box.height * 0.1);
  await page.mouse.down();
  await page.mouse.move(box.x + box.width * 0.25, box.y + box.height * 0.25, { steps: 4 });
  await page.mouse.move(box.x + box.width * 0.4, box.y + box.height * 0.45, { steps: 4 });
  await page.mouse.up();
  await expect(rows).toHaveCount(1);
  await expect(rows.first()).toContainText("Redact");
  // A drag is one undo entry, not one per move.
  await expect(page.getByTestId("edit-undo")).toHaveAttribute("title", "Undo Redact");
  await expect(page.getByTestId("edit-redaction-notice")).toHaveText(REDACTION_NOTICE);

  // --- AC-48: Flatten and discard original ---------------------------------
  await page.getByTestId("edit-flatten").click();
  const flatten = page.getByTestId("edit-flatten-dialog");
  await expect(flatten).toBeVisible();
  await expect(page.getByTestId("edit-flatten-message")).toContainText("permanent");
  await expect(page.getByTestId("edit-flatten-message")).toContainText("cannot be undone");
  await page.getByTestId("edit-flatten-confirm").click();
  await expect(flatten).toBeHidden();
  await expect(page.getByTestId("edit-notice")).toContainText("Flattened");
  await expect(rows).toHaveCount(0);
  await expect(page.getByTestId("edit-history-empty")).toBeVisible();
  await expect(page.getByTestId("detail-dimensions")).toHaveText("160 × 120 px");
  await expect(page.getByTestId("detail-original-dimensions")).toHaveCount(0);
  await expect(page.getByTestId("edit-redaction-notice")).toHaveCount(0);
  await expect(page.getByTestId("edit-undo")).toBeDisabled();
  await page.keyboard.press("ControlOrMeta+z");
  await expect(rows).toHaveCount(0);

  // The stored source IS the redacted render now: black under the cover,
  // the photo's grey quadrant elsewhere — and it stays so after reopening.
  await page.getByTestId("edit-done").click();
  await expect(page.getByTestId("media-edit")).toBeHidden();
  expect(await naturalSize(viewed)).toEqual([160, 120]);
  expectNear(await samplePixel(viewed, 40, 30), [0, 0, 0]);
  expectNear(await samplePixel(viewed, 140, 100), [128, 128, 128]);
  await page.getByTestId("close-document").click();
  await rowByTitle(page, "spec107-photo").dblclick();
  await expect(page).toHaveURL(/\/media\/.+/);
  expectNear(await samplePixel(page.getByTestId("media-image"), 40, 30), [0, 0, 0]);
});

test("Escape leaves the editor; Cmd/Ctrl+] rotates; unsaved edits are confirmed before they are dropped", async ({ page }) => {
  await page.goto("/");
  await expect(page.getByTestId("app-bar")).toBeVisible();
  await dropPhoto(page);
  await openEditor(page);
  await page.getByTestId("tool-rotate").focus();
  await page.keyboard.press("ControlOrMeta+]");
  await expect(page.getByTestId("detail-dimensions")).toHaveText("120 × 160 px");
  await page.keyboard.press("ControlOrMeta+[");
  await expect(page.getByTestId("detail-dimensions")).toHaveText("160 × 120 px");
  await page.keyboard.press("ControlOrMeta+[");
  await expect(page.getByTestId("detail-dimensions")).toHaveText("120 × 160 px");
  // Three turns are one rotate entry in the history: a record of intent
  // collapsed per kind, not a log of key presses.
  await expect(page.getByTestId("edit-history-row")).toHaveCount(1);
  await expect(page.getByTestId("edit-history-row")).toContainText("Rotate left");

  await page.keyboard.press("Escape");
  await expect(page.getByTestId("edit-discard-dialog")).toBeVisible();
  await page.getByTestId("edit-discard-confirm").click();
  await expect(page.getByTestId("media-edit")).toBeHidden();
  // Nothing was saved: the stored photo is untouched.
  await expect(page.getByTestId("detail-dimensions")).toHaveText("160 × 120 px");
  await openEditor(page);
  await expect(page.getByTestId("edit-history-row")).toHaveCount(0);
  await page.keyboard.press("Escape");
  await expect(page.getByTestId("media-edit")).toBeHidden();
});
