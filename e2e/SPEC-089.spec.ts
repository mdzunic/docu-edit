import { test, expect, type Page } from "@playwright/test";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

// SPEC-089 — slide layouts. Open the layouts fixture, add a Section-header slide
// through the filmstrip's "New slide with layout…" gallery, type a title, change
// the slide to Two Content through the toolbar Layout gallery (the title text
// survives and a second body placeholder appears), drag the title away and Reset
// slide (it snaps back), then download, reopen and confirm the slide still
// reports "Two Content" as its layout.
const here = dirname(fileURLToPath(import.meta.url));
const FIXTURE = join(here, "fixtures", "spec089-layouts.pptx");

/** The text frames on the MAIN stage (not the filmstrip thumbnails). */
function stageFrames(page: Page) {
  return page.getByTestId("stage-workspace").getByTestId("slide-frame");
}
function gallery(page: Page) {
  return page.getByTestId("slides-layout-gallery");
}

async function openFixture(page: Page, file = FIXTURE): Promise<void> {
  await page.goto("/");
  await page.getByTestId("open-document-input").setInputFiles(file);
  await expect(page).toHaveURL(/\/slides\/.+/);
  await expect(page.getByTestId("presentation-view")).toBeVisible();
}

async function downloadTo(page: Page, name: string): Promise<string> {
  const downloadPromise = page.waitForEvent("download");
  await page.getByTestId("presentation-download").click();
  const download = await downloadPromise;
  const filePath = join(tmpdir(), name);
  await download.saveAs(filePath);
  return filePath;
}

test("pick a layout for a new slide, change layout keeping text, reset, and round-trip", async ({ page }) => {
  await openFixture(page);
  await expect(page.getByTestId("filmstrip-slide")).toHaveCount(2);

  // Filmstrip menu → "New slide with layout…" → the gallery of real thumbnails.
  await page.getByTestId("presentation-slide-menu").first().click();
  await page.getByTestId("slide-menu-new-with-layout").click();
  await expect(gallery(page)).toBeVisible();
  await expect(gallery(page).getByRole("option")).toHaveCount(7);
  // Each card carries a slide-stage thumbnail.
  await expect(gallery(page).getByTestId("slide-stage").first()).toBeVisible();

  // Choose "Section Header" and add the slide.
  await gallery(page).getByRole("option", { name: "Section Header" }).click();
  await page.getByTestId("layout-apply").click();
  await expect(gallery(page)).toBeHidden();
  await expect(page.getByTestId("filmstrip-slide")).toHaveCount(3);
  // The new slide carries exactly the section-header placeholders (title + body).
  await expect(stageFrames(page)).toHaveCount(2);

  // Type a title into the title placeholder.
  await stageFrames(page).first().dblclick();
  await page.keyboard.type("Kickoff");
  await expect(stageFrames(page).first()).toContainText("Kickoff");

  // Toolbar Layout → change to "Two Content".
  await page.getByTestId("slides-layout-toggle").click();
  await expect(gallery(page)).toBeVisible();
  await gallery(page).getByRole("option", { name: "Two Content" }).click();
  await page.getByTestId("layout-apply").click();
  await expect(gallery(page)).toBeHidden();
  // The title text survives and two body placeholders appear (3 frames total).
  await expect(stageFrames(page).first()).toContainText("Kickoff");
  await expect(stageFrames(page)).toHaveCount(3);

  // Drag the title away, then Reset slide — it snaps back to the layout slot.
  const title = stageFrames(page).first();
  await title.click();
  const before = await title.boundingBox();
  if (!before) throw new Error("no title box");
  await page.mouse.move(before.x + before.width / 2, before.y + before.height / 2);
  await page.mouse.down();
  await page.mouse.move(before.x + before.width / 2 + 40, before.y + before.height / 2 + 120, { steps: 8 });
  await page.mouse.up();
  await expect
    .poll(async () => (await title.boundingBox())?.y ?? 0)
    .toBeGreaterThan(before.y + 40);

  await page.getByTestId("slides-layout-toggle").click();
  await expect(gallery(page)).toBeVisible();
  await page.getByTestId("layout-reset-slide").click();
  await expect
    .poll(async () => Math.abs(((await title.boundingBox())?.y ?? 0) - before.y))
    .toBeLessThan(12);

  // Download, reopen, and confirm the slide still reports "Two Content".
  const file = await downloadTo(page, "spec089-layouts-edited.pptx");
  await page.goto("/");
  await page.getByTestId("open-document-input").setInputFiles(file);
  await expect(page.getByTestId("presentation-view")).toBeVisible();
  await page.getByTestId("filmstrip-slide").nth(1).click();
  await expect(stageFrames(page).first()).toContainText("Kickoff");
  await page.getByTestId("slides-layout-toggle").click();
  await expect(
    gallery(page).getByRole("option", { name: "Two Content", selected: true }),
  ).toBeVisible();
});
