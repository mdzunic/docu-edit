import { test, expect, type Page } from "@playwright/test";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { openInsertTab } from "./slides";

const here = dirname(fileURLToPath(import.meta.url));
const FIXTURE_PNG = join(here, "fixtures", "fixture-image.png");
const FIXTURE_PPTX = join(here, "fixtures", "mini-deck.pptx");

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
  await openInsertTab(page);

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
  const imageGroup = page.getByTestId("stage-workspace").getByRole("group", { name: "Image" });
  await expect(imageGroup).toBeVisible();
  // Renders immediately: the <img> resolves to a data URI without a reload (AC-8).
  await expect(imageGroup.locator("img")).toHaveAttribute("src", /^data:image\//);

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

test("select and ungroup a deck's group whose grpSpPr carries no own xfrm (AC-31/32)", async ({ page }) => {
  await page.goto("/");
  await page.getByTestId("open-document-input").setInputFiles(FIXTURE_PPTX);
  await expect(page).toHaveURL(/\/slides\/.+/);
  await expect(page.getByTestId("presentation-view")).toBeVisible();

  // Slide 3 carries a plain text box plus 'Group 3', whose <p:grpSpPr/> has no
  // xfrm of its own; the reader must derive the group's box from the union of
  // its children (AC-31) or clicking where the group is drawn selects nothing.
  await page.getByTestId("filmstrip-slide").nth(2).click();
  const stage = page.getByTestId("stage-workspace");
  await expect(stage.getByRole("group", { name: "Group" })).toHaveCount(1);

  // Click the group's rendered content to select it; the Ungroup control only
  // appears for a group selection.
  await stage.getByText("Grouped note").click();
  await openInsertTab(page);
  await expect(page.getByTestId("ungroup-button")).toBeVisible();

  // The selection chrome is sized from the group's own box; without the
  // union-of-children fallback that box is 0×0 and the group is unusable — no
  // visible handles, no move/resize target (AC-31).
  const selBox = await stage.getByTestId("frame-selection").boundingBox();
  if (!selBox) throw new Error("no selection box");
  expect(selBox.width).toBeGreaterThan(1);
  expect(selBox.height).toBeGreaterThan(1);

  // Ungroup promotes the child in place: the group vanishes, its text survives,
  // and the selection is now a non-group child (so Ungroup disappears) (AC-32).
  await page.getByTestId("ungroup-button").click();
  await expect(stage.getByRole("group", { name: "Group" })).toHaveCount(0);
  await expect(stage).toContainText("Grouped note");
  await expect(page.getByTestId("ungroup-button")).toHaveCount(0);
});

async function downloadTo(page: Page, name: string): Promise<string> {
  const downloadPromise = page.waitForEvent("download");
  await page.getByTestId("presentation-download").click();
  const download = await downloadPromise;
  const filePath = join(tmpdir(), name);
  await download.saveAs(filePath);
  return filePath;
}
