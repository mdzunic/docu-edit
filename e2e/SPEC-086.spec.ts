import { test, expect, type Page } from "@playwright/test";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

// SPEC-086 — slide animations. Open the animation fixture, add "Fade in, on
// click" to a body frame and "Fly in from left, after previous" to a shape
// through the Animations panel, Present, and confirm both frames start hidden
// then build in on one advance (the after-previous shape chains after the body
// with no further input), a second advance moves to the next slide, then exit,
// download, reopen and confirm the panel lists both effects in order.
const here = dirname(fileURLToPath(import.meta.url));
const FIXTURE = join(here, "fixtures", "spec086-animations.pptx");

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

test("add animations, build them in playback, and round-trip the panel", async ({ page }) => {
  await openFixture(page);
  const deckId = page.url().match(/\/slides\/([^/?]+)/)?.[1] ?? "";
  expect(deckId).not.toBe("");

  await openPanel(page);
  const workspace = page.getByTestId("stage-workspace");

  // Add "Fade in, on click" to the body frame (Add defaults to fade-in/click).
  await workspace.getByText("Body line").click();
  await expect(page.getByTestId("animation-add")).toBeEnabled();
  await page.getByTestId("animation-add").click();
  await expect(page.getByTestId("animation-row")).toHaveCount(1);

  // Add the shape, then set Fly in / from left / after previous.
  await workspace.getByText("Shape text").click();
  await page.getByTestId("animation-add").click();
  await expect(page.getByTestId("animation-row")).toHaveCount(2);
  await page.getByTestId("animation-effect").selectOption("flyIn");
  await page.getByTestId("animation-direction").selectOption("left");
  await page.getByTestId("animation-trigger").selectOption("afterPrevious");

  // Present from slide 1.
  await page.getByTestId("slides-present").click();
  await expect(page).toHaveURL(new RegExp(`/slides/${deckId}/present/1$`));
  const play = page.getByTestId("presentation-play");
  await expect(play).toBeVisible();
  const stage = play.getByTestId("slide-stage");

  // The slide appears with both animated frames hidden; the title is visible.
  await expect(stage.getByText("Animate me")).toBeVisible();
  await expect(stage.getByText("Body line")).toBeHidden();
  await expect(stage.getByText("Shape text")).toBeHidden();

  // One advance plays the whole click group: the body fades in, then the shape
  // flies in after it — no further input. The slide index does not change.
  await page.keyboard.press("ArrowRight");
  await expect(stage.getByText("Body line")).toBeVisible();
  await expect(stage.getByText("Shape text")).toBeVisible();
  await expect(page).toHaveURL(new RegExp(`/slides/${deckId}/present/1$`));

  // The build settles: the flown-in shape comes to rest (no residual transform).
  const shapeFrame = stage.locator(".frame").filter({ hasText: "Shape text" });
  await expect(shapeFrame).toHaveCSS("transform", "none");

  // Advance again — now that the group has fully built, this moves to the next
  // slide. (A mid-build advance instead completes the group and holds — AC-24,
  // covered in the useBuildPreview unit suite.)
  await page.keyboard.press("ArrowRight");
  await expect(page).toHaveURL(new RegExp(`/slides/${deckId}/present/2$`));

  // Exit playback.
  await page.keyboard.press("Escape");
  await expect(page.getByTestId("presentation-view")).toBeVisible();

  // Download and reopen.
  const downloadPromise = page.waitForEvent("download");
  await page.getByTestId("presentation-download").click();
  const download = await downloadPromise;
  const filePath = join(tmpdir(), "spec086-deck.pptx");
  await download.saveAs(filePath);

  await openFixture(page, filePath);
  await openPanel(page);
  // The reopened deck's first slide lists both effects in order.
  const rows = page.getByTestId("animation-row");
  await expect(rows).toHaveCount(2);
  await expect(rows.nth(0)).toContainText("Fade in");
  await expect(rows.nth(1)).toContainText("Fly in");
});

test("an advance mid-build completes the current group instantly and holds (AC-24)", async ({ page }) => {
  await openFixture(page);
  const deckId = page.url().match(/\/slides\/([^/?]+)/)?.[1] ?? "";
  expect(deckId).not.toBe("");

  await openPanel(page);
  const workspace = page.getByTestId("stage-workspace");

  // Body fades in on click over a long 5s window so the first advance is caught
  // with the group still animating.
  await workspace.getByText("Body line").click();
  await page.getByTestId("animation-add").click();
  await expect(page.getByTestId("animation-row")).toHaveCount(1);
  await page.getByTestId("animation-duration").fill("5");

  // The shape flies in after the body — but only after a 10s delay, so the click
  // group stays in flight for well over ten seconds. This makes the mid-build
  // window deterministic: a second advance during it cannot race the settle.
  await workspace.getByText("Shape text").click();
  await page.getByTestId("animation-add").click();
  await expect(page.getByTestId("animation-row")).toHaveCount(2);
  await page.getByTestId("animation-effect").selectOption("flyIn");
  await page.getByTestId("animation-direction").selectOption("left");
  await page.getByTestId("animation-trigger").selectOption("afterPrevious");
  await page.getByTestId("animation-delay").fill("10");

  // Present from slide 1; both animated frames start hidden.
  await page.getByTestId("slides-present").click();
  await expect(page).toHaveURL(new RegExp(`/slides/${deckId}/present/1$`));
  const play = page.getByTestId("presentation-play");
  const stage = play.getByTestId("slide-stage");
  await expect(stage.getByText("Body line")).toBeHidden();
  await expect(stage.getByText("Shape text")).toBeHidden();

  // First advance starts the click group building. The after-previous shape is
  // still hidden (its 10s delay has not elapsed) — the group is mid-build.
  await page.keyboard.press("ArrowRight");
  await expect(stage.getByText("Shape text")).toBeHidden();
  await expect(page).toHaveURL(new RegExp(`/slides/${deckId}/present/1$`));

  // AC-24: an advance arriving mid-build completes the group INSTANTLY — the
  // shape snaps straight to its rest state without waiting out the 10s delay —
  // and HOLDS on the slide rather than walking on to the next one.
  await page.keyboard.press("ArrowRight");
  await expect(stage.getByText("Body line")).toBeVisible();
  await expect(stage.getByText("Shape text")).toBeVisible();
  const shapeFrame = stage.locator(".frame").filter({ hasText: "Shape text" });
  await expect(shapeFrame).toHaveCSS("transform", "none");
  await expect(page).toHaveURL(new RegExp(`/slides/${deckId}/present/1$`));

  // Only once the group is fully built does a further advance move on.
  await page.keyboard.press("ArrowRight");
  await expect(page).toHaveURL(new RegExp(`/slides/${deckId}/present/2$`));
});
