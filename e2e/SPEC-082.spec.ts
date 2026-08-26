import { test, expect, type Page } from "@playwright/test";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

const here = dirname(fileURLToPath(import.meta.url));
const FIXTURE_PNG = join(here, "fixtures", "fixture-image.png");

// SPEC-082 — slide objects: from a blank deck, insert a text box and type in
// it, draw a shape, nudge it, send it to back, insert an image via the picker,
// then download, re-open, and confirm all three new frames survive with the
// expected stacking (shape at the back, image on top) and the image centred.

function stageFrames(page: Page) {
  return page.getByTestId("stage-workspace").getByTestId("slide-frame");
}

async function newDeck(page: Page): Promise<void> {
  await page.goto("/");
  await page.getByTestId("new-presentation").click();
  await expect(page).toHaveURL(/\/slides\/.+/);
  await expect(page.getByTestId("presentation-view")).toBeVisible();
}

async function stageBox(page: Page) {
  const box = await page.getByTestId("stage-workspace").getByTestId("slide-stage").boundingBox();
  if (!box) throw new Error("no stage box");
  return box;
}

test("insert a text box, a shape and an image; nudge, restack, round-trip", async ({ page }) => {
  await newDeck(page);
  const before = await stageFrames(page).count(); // 2 placeholders

  // --- text box: arm, click to place the default box, type into it ----------
  await page.getByTestId("insert-text-box").click();
  let box = await stageBox(page);
  await page.mouse.click(box.x + box.width * 0.3, box.y + box.height * 0.2);
  await expect(stageFrames(page)).toHaveCount(before + 1);
  await expect(page.getByTestId("slide-text-edit")).toBeVisible();
  await page.keyboard.type("Hello slide");
  await page.keyboard.press("Escape");
  await expect(page.getByTestId("stage-workspace")).toContainText("Hello slide");

  // --- shape: arm from the gallery, draw it by drag -------------------------
  await page.getByTestId("insert-shape").click();
  await page.getByTestId("shape-gallery-rect").click();
  box = await stageBox(page);
  // Draw in the empty band below the placeholders so the gesture is not stolen
  // by a frame under the pointer.
  await page.mouse.move(box.x + box.width * 0.2, box.y + box.height * 0.72);
  await page.mouse.down();
  await page.mouse.move(box.x + box.width * 0.42, box.y + box.height * 0.9, { steps: 8 });
  await page.mouse.up();
  await expect(stageFrames(page)).toHaveCount(before + 2);
  // The freshly drawn shape is selected and named from the catalogue.
  await expect(page.getByRole("group", { name: "Rectangle" })).toBeVisible();

  // --- nudge the shape with the arrow keys ----------------------------------
  const shape = page.getByTestId("stage-workspace").getByRole("group", { name: "Rectangle" });
  await shape.click();
  const nudgedFrom = await shape.boundingBox();
  await page.keyboard.press("Shift+ArrowRight");
  await page.keyboard.press("Shift+ArrowRight");
  await expect
    .poll(async () => (await shape.boundingBox())?.x ?? 0)
    .toBeGreaterThan((nudgedFrom?.x ?? 0) + 5);

  // --- send the shape to back -----------------------------------------------
  await page.getByTestId("arrange-menu").click();
  await page.locator('[data-arrange="back"]').click();
  // It is now the first frame in z-order.
  await expect(stageFrames(page).first()).toHaveAttribute("aria-label", "Rectangle");

  // --- image: insert the fixture PNG through the picker ---------------------
  await page.getByTestId("slides-image-input").setInputFiles(FIXTURE_PNG);
  await expect(stageFrames(page)).toHaveCount(before + 3);
  await expect(page.getByRole("group", { name: "Image" })).toBeVisible();

  // --- download, re-open, verify the three frames + stacking + position -----
  const file = await downloadTo(page, "spec082-deck.pptx");
  await page.goto("/");
  await page.getByTestId("open-document-input").setInputFiles(file);
  await expect(page).toHaveURL(/\/slides\/.+/);
  await expect(page.getByTestId("presentation-view")).toBeVisible();

  await expect(stageFrames(page)).toHaveCount(before + 3);
  await expect(page.getByTestId("stage-workspace")).toContainText("Hello slide");
  // Stacking survives: shape at the back (first), image on top (last).
  await expect(stageFrames(page).first()).toHaveAttribute("aria-label", "Rectangle");
  await expect(stageFrames(page).last()).toHaveAttribute("aria-label", "Image");

  // Position survives: the inserted image sits centred on the slide.
  const stage = await stageBox(page);
  const image = await page.getByRole("group", { name: "Image" }).boundingBox();
  if (!image) throw new Error("no image box");
  const imageCx = image.x + image.width / 2;
  const imageCy = image.y + image.height / 2;
  expect(Math.abs(imageCx - (stage.x + stage.width / 2))).toBeLessThan(stage.width * 0.1);
  expect(Math.abs(imageCy - (stage.y + stage.height / 2))).toBeLessThan(stage.height * 0.1);
});

async function downloadTo(page: Page, name: string): Promise<string> {
  const downloadPromise = page.waitForEvent("download");
  await page.getByTestId("presentation-download").click();
  const download = await downloadPromise;
  const filePath = join(tmpdir(), name);
  await download.saveAs(filePath);
  return filePath;
}
