import { test, expect, type Locator, type Page } from "@playwright/test";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

// SPEC-105 — animation depth. Open the animation-depth fixture, give the
// bullet body "Fade in, by paragraph" (the timeline shows four bars in four
// groups), Present and build the four bullets one advance at a time; back in
// the editor add a "Spin" on the title triggered "On click of" the shape (a
// rounded rectangle) and a "Line" motion path on the shape; Present again and
// click the shape — the title spins and the slide does not advance; then exit,
// download, reopen, and find the build, the spin and the path listed in order
// with their fields intact.
const here = dirname(fileURLToPath(import.meta.url));
const FIXTURE = join(here, "fixtures", "spec105-animation-depth.pptx");
const BULLETS = ["First point", "Second point", "Third point", "Fourth point"];

async function openFixture(page: Page, file = FIXTURE): Promise<void> {
  await page.goto("/");
  await page.getByTestId("open-document-input").setInputFiles(file);
  await expect(page).toHaveURL(/\/slides\/.+/);
  await expect(page.getByTestId("presentation-view")).toBeVisible();
}

async function openPanel(page: Page): Promise<void> {
  const panel = page.getByTestId("slides-animation-panel");
  if ((await panel.count()) === 0) {
    await page.getByTestId("slides-animate-toggle").click();
  }
  await expect(panel).toBeVisible();
}

/** Every bullet from `shown` on is hidden; the ones before it are visible. */
async function expectBulletsShown(stage: Locator, shown: number): Promise<void> {
  for (const [index, text] of BULLETS.entries()) {
    if (index < shown) {
      await expect(stage.getByText(text)).toBeVisible();
    } else {
      await expect(stage.getByText(text)).toBeHidden();
    }
  }
}

test("by-paragraph build, click-a-shape spin and a motion path round-trip", async ({ page }) => {
  await openFixture(page);
  const deckId = page.url().match(/\/slides\/([^/?]+)/)?.[1] ?? "";
  expect(deckId).not.toBe("");

  await openPanel(page);
  const workspace = page.getByTestId("stage-workspace");

  // Select the bullet body, Animate it: "Fade in", "By paragraph".
  await workspace.getByText("First point").click();
  await expect(page.getByTestId("animation-add")).toBeEnabled();
  await page.getByTestId("animation-add").click();
  await expect(page.getByTestId("animation-row")).toHaveCount(1);
  await page.getByTestId("animation-effect").selectOption("fadeIn");
  await expect(page.getByTestId("animation-build-by")).toBeEnabled();
  await page.getByTestId("animation-build-by").selectOption("byParagraph");
  await expect(page.getByTestId("animation-build-trigger")).toBeVisible();

  // The timeline shows four bars in four click groups.
  const toggle = page.getByTestId("animation-timeline-toggle");
  if ((await toggle.getAttribute("aria-pressed")) !== "true") {
    await toggle.click();
  }
  await expect(toggle).toHaveAttribute("aria-pressed", "true");
  const timeline = page.getByTestId("animation-timeline");
  await expect(timeline).toBeVisible();
  await expect(timeline.getByTestId("animation-timeline-ruler")).toBeVisible();
  await expect(timeline.locator('[data-testid^="animation-timeline-group-"]')).toHaveCount(4);
  await expect(timeline.locator('[data-testid^="animation-timeline-bar-"]')).toHaveCount(4);
  for (let g = 0; g < 4; g++) {
    await expect(timeline.getByTestId(`animation-timeline-group-${g}`).locator('[data-testid^="animation-timeline-bar-"]')).toHaveCount(1);
  }

  // Present: the slide shows with no bullets; each advance builds one.
  await page.getByTestId("slides-present").click();
  await expect(page).toHaveURL(new RegExp(`/slides/${deckId}/present/1$`));
  const play = page.getByTestId("presentation-play");
  const stage = play.getByTestId("slide-stage");
  await expect(stage.getByText("Bullets")).toBeVisible();
  await expectBulletsShown(stage, 0);
  for (let shown = 1; shown <= BULLETS.length; shown++) {
    await page.keyboard.press("ArrowRight");
    await expectBulletsShown(stage, shown);
    await expect(play).toHaveAttribute("data-building", "false");
    await expect(page).toHaveURL(new RegExp(`/slides/${deckId}/present/1$`));
  }

  // Back in the editor.
  await page.keyboard.press("Escape");
  await expect(page.getByTestId("presentation-view")).toBeVisible();
  await openPanel(page);

  // An emphasis "Spin" on the title, triggered "On click of" the shape.
  await workspace.getByText("Bullets").click();
  await page.getByTestId("animation-add").click();
  await expect(page.getByTestId("animation-row")).toHaveCount(2);
  await page.getByTestId("animation-effect").selectOption("spin");
  await expect(page.getByTestId("animation-spin-degrees")).toBeVisible();
  await expect(page.getByTestId("animation-trigger-frame")).toBeDisabled();
  await page.getByTestId("animation-trigger").selectOption("onClickOf");
  await expect(page.getByTestId("animation-trigger-frame")).toBeEnabled();
  await page.getByTestId("animation-trigger-frame").selectOption({ label: "Rounded rectangle" });

  // A "Line" motion path on the shape.
  // (The shape paints its fill on a canvas over its text: click the frame.)
  await workspace.getByTestId("slide-frame").filter({ hasText: "Press me" }).click();
  await page.getByTestId("animation-add").click();
  await expect(page.getByTestId("animation-row")).toHaveCount(3);
  await page.getByTestId("animation-effect").selectOption("motionPath");
  await page.getByTestId("animation-path-preset").selectOption("line");
  await expect(page.getByTestId("animation-path-distance")).toHaveValue("25");

  // Present: clicking the shape spins the title without advancing.
  await page.getByTestId("slides-present").click();
  await expect(page).toHaveURL(new RegExp(`/slides/${deckId}/present/1$`));
  await expectBulletsShown(stage, 0);
  const title = stage.locator(".frame").filter({ hasText: "Bullets" });
  await stage.locator(".frame").filter({ hasText: "Press me" }).click();
  // The spin ran and settled a full turn round; no build group advanced.
  await expect(title).toHaveCSS("rotate", "360deg");
  await expect(play).toHaveAttribute("data-building", "false");
  await expectBulletsShown(stage, 0);
  await expect(page).toHaveURL(new RegExp(`/slides/${deckId}/present/1$`));
  await expect(play.getByTestId("play-slide-counter")).toHaveText("1 / 4");

  // Exit, download and reopen.
  await page.keyboard.press("Escape");
  await expect(page.getByTestId("presentation-view")).toBeVisible();
  const downloadPromise = page.waitForEvent("download");
  await page.getByTestId("presentation-download").click();
  const download = await downloadPromise;
  const filePath = join(tmpdir(), "spec105-deck.pptx");
  await download.saveAs(filePath);

  await openFixture(page, filePath);
  await openPanel(page);
  // The build, the spin and the path, in order, their fields intact.
  const rows = page.getByTestId("animation-row");
  await expect(rows).toHaveCount(3);
  await expect(rows.nth(0)).toContainText("Fade in");
  await expect(rows.nth(0)).toContainText("By paragraph");
  await expect(rows.nth(1)).toContainText("Spin");
  await expect(rows.nth(2)).toContainText("Motion path");

  await rows.nth(0).click();
  await expect(page.getByTestId("animation-effect")).toHaveValue("fadeIn");
  await expect(page.getByTestId("animation-build-by")).toHaveValue("byParagraph");

  await rows.nth(1).click();
  await expect(page.getByTestId("animation-effect")).toHaveValue("spin");
  await expect(page.getByTestId("animation-trigger")).toHaveValue("onClickOf");
  await expect(page.getByTestId("animation-trigger-frame").locator("option:checked")).toHaveText("Rounded rectangle");

  await rows.nth(2).click();
  await expect(page.getByTestId("animation-effect")).toHaveValue("motionPath");
  await expect(page.getByTestId("animation-path-preset")).toHaveValue("line");
  await expect(page.getByTestId("animation-path-direction")).toHaveValue("right");
  await expect(page.getByTestId("animation-path-distance")).toHaveValue("25");
});

test("the motion-path endpoint and the timeline edges drag, take focus and step by arrow", async ({ page }) => {
  await openFixture(page);
  await openPanel(page);
  const workspace = page.getByTestId("stage-workspace");

  // A "Line" motion path on the shape (right, 25%), its editor armed.
  const shape = workspace.getByTestId("slide-frame").filter({ hasText: "Press me" });
  await shape.click();
  await page.getByTestId("animation-add").click();
  await expect(page.getByTestId("animation-row")).toHaveCount(1);
  await page.getByTestId("animation-effect").selectOption("motionPath");
  await page.getByTestId("animation-path-preset").selectOption("line");
  await expect(page.getByTestId("animation-path-direction")).toHaveValue("right");
  await expect(page.getByTestId("animation-path-distance")).toHaveValue("25");
  await page.getByTestId("animation-path-edit").click();
  const endpoint = page.getByTestId("animation-path-endpoint");
  await expect(endpoint).toBeVisible();

  // Drag the endpoint straight down: the path turns down, the panel follows,
  // and the endpoint keeps focus for the arrows.
  const box = await endpoint.boundingBox();
  if (!box) throw new Error("no endpoint box");
  const x = box.x + box.width / 2;
  const y = box.y + box.height / 2;
  const stageBox = await page.getByTestId("animation-path-overlay").boundingBox();
  if (!stageBox) throw new Error("no overlay box");
  await page.mouse.move(x, y);
  await page.mouse.down();
  await page.mouse.move(x - stageBox.width * 0.25, y + stageBox.height * 0.4, { steps: 12 });
  await page.mouse.up();
  await expect(page.getByTestId("animation-path-direction")).toHaveValue("down");
  await expect(page.getByTestId("animation-path-distance")).toHaveValue("40");
  await expect(endpoint).toHaveAttribute("aria-label", "Path end: down, 40%");
  await expect(endpoint).toBeFocused();
  // The arrows step the endpoint, never the (still selected) shape too.
  const shapeStyle = await shape.getAttribute("style");
  await page.keyboard.press("ArrowDown");
  await expect(page.getByTestId("animation-path-distance")).toHaveValue("41");
  await expect(endpoint).toHaveAttribute("aria-label", "Path end: down, 41%");
  expect(await shape.getAttribute("style")).toBe(shapeStyle);
  // One undo entry per gesture: the key press, then the whole drag.
  await page.keyboard.press("Control+z");
  await expect(page.getByTestId("animation-path-distance")).toHaveValue("40");
  await page.keyboard.press("Control+z");
  await expect(page.getByTestId("animation-path-direction")).toHaveValue("right");
  await expect(page.getByTestId("animation-path-distance")).toHaveValue("25");
  await page.getByTestId("animation-path-edit").click();
  await expect(endpoint).toHaveCount(0);

  // The timeline: dragging the bar's right edge lengthens it in 100ms steps,
  // the edge takes focus, and Right then steps it once more.
  const toggle = page.getByTestId("animation-timeline-toggle");
  if ((await toggle.getAttribute("aria-pressed")) !== "true") {
    await toggle.click();
  }
  const timeline = page.getByTestId("animation-timeline");
  const durationEdge = timeline.getByTestId("animation-timeline-duration-0");
  await expect(durationEdge).toHaveAttribute("aria-valuenow", "500");
  const edge = await durationEdge.boundingBox();
  if (!edge) throw new Error("no edge box");
  await page.mouse.move(edge.x + edge.width / 2, edge.y + edge.height / 2);
  await page.mouse.down();
  // 60px per second: 30px is half a second.
  await page.mouse.move(edge.x + edge.width / 2 + 30, edge.y + edge.height / 2, { steps: 6 });
  await page.mouse.up();
  await expect(durationEdge).toHaveAttribute("aria-valuenow", "1000");
  await expect(durationEdge).toBeFocused();
  await page.keyboard.press("ArrowRight");
  await expect(durationEdge).toHaveAttribute("aria-valuenow", "1100");
  expect(await shape.getAttribute("style")).toBe(shapeStyle);

  // A click on the delay edge focuses it too; Right then steps the delay.
  const delayEdge = timeline.getByTestId("animation-timeline-delay-0");
  await expect(delayEdge).toHaveAttribute("aria-valuenow", "0");
  await delayEdge.click();
  await expect(delayEdge).toBeFocused();
  await page.keyboard.press("ArrowRight");
  await expect(delayEdge).toHaveAttribute("aria-valuenow", "100");
  await expect(durationEdge).toHaveAttribute("aria-valuenow", "1100");
  expect(await shape.getAttribute("style")).toBe(shapeStyle);
});
