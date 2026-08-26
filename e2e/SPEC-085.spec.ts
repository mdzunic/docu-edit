import { test, expect, type Page } from "@playwright/test";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

// SPEC-085 — slide transitions. Open the pptx fixture, set Push/Left/1s on
// slide 1 through the Transition panel (Apply-to-all NOT used), Present, and
// confirm advancing to slide 2 shows movement (an intermediate transform
// sampled via page.evaluate, then a settled final state). Exit, download,
// reopen and confirm the panel round-trips the transition (pptx duration
// quantises 1s → the "med" bucket, still 1s here as the model keeps exact ms
// until reopen; slow/med/fast map back to 1000/600/300 so 1s survives).
const here = dirname(fileURLToPath(import.meta.url));
const FIXTURE_PPTX = join(here, "fixtures", "mini-deck.pptx");

async function openFixture(page: Page, file = FIXTURE_PPTX): Promise<void> {
  await page.goto("/");
  await page.getByTestId("open-document-input").setInputFiles(file);
  await expect(page).toHaveURL(/\/slides\/.+/);
  await expect(page.getByTestId("presentation-view")).toBeVisible();
}

async function setPushLeftOneSecond(page: Page): Promise<void> {
  await page.getByTestId("slides-transition-toggle").click();
  await expect(page.getByTestId("slides-transition-panel")).toBeVisible();
  await page.getByTestId("transition-effect-push").check();
  await page.getByTestId("transition-direction").selectOption("left");
  await page.getByTestId("transition-duration").fill("1");
  await page.getByTestId("transition-duration").blur();
}

test("set a transition, see movement in playback, and round-trip it", async ({ page }) => {
  await openFixture(page);
  const deckId = page.url().match(/\/slides\/([^/?]+)/)?.[1] ?? "";
  expect(deckId).not.toBe("");

  await setPushLeftOneSecond(page);
  // Apply-to-all is deliberately NOT clicked (only slide 1 gets the transition).

  // Present from slide 1.
  await page.getByTestId("slides-present").click();
  await expect(page).toHaveURL(new RegExp(`/slides/${deckId}/present/1$`));
  const play = page.getByTestId("presentation-play");
  await expect(play).toBeVisible();

  // Advance to slide 2 — the transition layer appears and animates.
  await page.keyboard.press("ArrowRight");
  const layer = page.getByTestId("slide-transition-layer");
  await expect(layer).toBeVisible();
  await expect(layer).toHaveAttribute("data-effect", "push");

  // Sample an INTERMEDIATE transform: the incoming stage is mid-translate
  // (neither fully off-screen nor settled) while the transition runs.
  const incoming = page.getByTestId("transition-incoming");
  const mid = await incoming.evaluate((el) => getComputedStyle(el).transform);
  expect(mid).not.toBe("none"); // a matrix() while animating, not the identity

  // …then it settles: the layer is gone and slide 2 is shown plainly.
  await expect(layer).toHaveCount(0);
  await expect(page.getByTestId("play-slide-counter")).toHaveText("2 / 4");

  // Exit back to the editor.
  await page.keyboard.press("Escape");
  await expect(page.getByTestId("presentation-view")).toBeVisible();

  // Download and reopen.
  const downloadPromise = page.waitForEvent("download");
  await page.getByTestId("presentation-download").click();
  const download = await downloadPromise;
  const filePath = join(tmpdir(), "spec085-deck.pptx");
  await download.saveAs(filePath);

  await openFixture(page, filePath);
  // The panel-open preference persists per device, so open it only if the
  // reopened deck shows it collapsed.
  const panel = page.getByTestId("slides-transition-panel");
  if ((await panel.count()) === 0) {
    await page.getByTestId("slides-transition-toggle").click();
  }
  await expect(panel).toBeVisible();
  // Slide 1 still reads Push / Left / 1s.
  await expect(page.getByTestId("transition-effect-push")).toBeChecked();
  await expect(page.getByTestId("transition-direction")).toHaveValue("left");
  await expect(page.getByTestId("transition-duration")).toHaveValue("1");
});
