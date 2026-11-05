import { test, expect, type Page } from "@playwright/test";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

// SPEC-146 — the show timer: a Stopwatch or a Countdown that starts with the
// show in the full-screen play view and in presenter view, is started, paused,
// resumed and stopped by key or button, survives a reload, crosses the
// presenter/audience pair and carries over when the audience promotes itself.
// Time is driven by page.clock.setFixedTime (the clock is per context, so it
// covers the audience popup too): every "time passes" step is another
// setFixedTime, and the auto-waiting toHaveText covers the next one-second
// tick. No timer assertion waits on the real clock. The minus sign of overtime
// is U+2212. Fullscreen grants are flaky headless, so nothing asserts them.
const here = dirname(fileURLToPath(import.meta.url));
const miniDeck = join(here, "fixtures", "mini-deck.pptx");
const T0 = new Date(2026, 10, 5, 9, 0, 0).getTime();

async function openFixture(page: Page, file: string): Promise<string> {
  await page.goto("/");
  await page.getByTestId("open-document-input").setInputFiles(file);
  await expect(page).toHaveURL(/\/slides\/.+/);
  await expect(page.getByTestId("presentation-view")).toBeVisible();
  const deckId = page.url().match(/\/slides\/([^/?]+)/)?.[1] ?? "";
  expect(deckId).not.toBe("");
  return deckId;
}

/** Present menu → "Presenter view": the presenter route here, the audience
 * window as the popup (as SPEC-100's e2e does it). */
async function startPresenterView(page: Page, deckId: string): Promise<Page> {
  await page.getByTestId("presentation-present-menu").click();
  await expect(page.getByTestId("presentation-present-menu-popup")).toBeVisible();
  const popup = page.waitForEvent("popup");
  await page.getByRole("menuitem", { name: "Presenter view" }).click();
  const audience = await popup;
  await expect(page).toHaveURL(new RegExp(`/slides/${deckId}/presenter/1$`));
  await expect(page.getByTestId("presenter-view")).toBeVisible();
  await expect(audience).toHaveURL(new RegExp(`/slides/${deckId}/present/1\\?follow=[^&]+`));
  await expect(audience.getByTestId("presentation-play")).toBeVisible();
  return audience;
}

/** Library → Preferences, so the show timer's settings are in place before the deck opens. */
async function withPreferences(page: Page, set: (dialog: ReturnType<Page["getByTestId"]>) => Promise<void>): Promise<void> {
  await page.goto("/");
  await page.getByTestId("library-preferences").click();
  const dialog = page.getByTestId("preferences-dialog");
  await expect(dialog).toBeVisible();
  await set(dialog);
  await page.getByTestId("preferences-close").click();
  await expect(dialog).toHaveCount(0);
}

/** Moves the pointer over the stage so the idle-fading HUD shows. */
async function revealHud(page: Page): Promise<void> {
  await page.mouse.move(120, 120);
  await page.mouse.move(300, 300);
}

test("Stopwatch in the full-screen show: starts with the show, pauses, stops, survives a reload, fresh after Escape", async ({ page }) => {
  await page.clock.setFixedTime(T0);
  const deckId = await openFixture(page, miniDeck);
  await page.getByTestId("slides-present").click();
  await expect(page).toHaveURL(new RegExp(`/slides/${deckId}/present/1$`));
  await revealHud(page);
  const timer = page.getByTestId("play-timer");
  const toggle = page.getByTestId("play-timer-toggle");
  const stop = page.getByTestId("play-timer-stop");
  const counter = page.getByTestId("play-slide-counter");
  await expect(timer).toHaveText("0:00");
  await expect(timer).toHaveAttribute("data-state", "running");
  await expect(timer).toHaveAttribute("data-mode", "stopwatch");
  await expect(timer).toHaveAttribute("role", "timer");
  await expect(toggle).toHaveText("Pause (T)");
  // The HUD timer is read-only: no mode or length controls in front of the room.
  await expect(page.getByTestId("play-timer-mode-countdown")).toHaveCount(0);
  await expect(page.getByTestId("play-timer-length")).toHaveCount(0);

  await page.clock.setFixedTime(T0 + 75_000);
  await expect(timer).toHaveText("1:15");

  // t pauses; the paused time is not counted, and ArrowRight still advances.
  await page.keyboard.press("t");
  await expect(timer).toHaveAttribute("data-state", "paused");
  await expect(toggle).toHaveText("Resume (T)");
  await expect(page.getByTestId("play-timer-status")).toHaveText("Timer paused");
  await page.clock.setFixedTime(T0 + 675_000);
  await expect(timer).toHaveText("1:15");
  await page.keyboard.press("ArrowRight");
  await expect(counter).toHaveText("2 / 4");
  await expect(timer).toHaveText("1:15");

  // Stop zeroes and halts it, and the click does not advance the show.
  await revealHud(page);
  await stop.click();
  await expect(timer).toHaveText("0:00");
  await expect(timer).toHaveAttribute("data-state", "stopped");
  await expect(toggle).toHaveText("Start (T)");
  await expect(stop).toBeDisabled();
  await expect(counter).toHaveText("2 / 4");

  // t starts a fresh count.
  await page.keyboard.press("t");
  await expect(timer).toHaveAttribute("data-state", "running");
  await expect(timer).toHaveText("0:00");

  // A reload restores it, the reload's own time counted.
  await page.clock.setFixedTime(T0 + 705_000);
  await page.reload();
  await expect(page.getByTestId("presentation-play")).toBeVisible();
  await expect(timer).toHaveText("0:30");
  await expect(timer).toHaveAttribute("data-state", "running");

  // Escape leaves the show (removing the record); Present again starts fresh.
  await page.keyboard.press("Escape");
  await expect(page).toHaveURL(new RegExp(`/slides/${deckId}\\?return=2$`));
  await expect(page.getByTestId("presentation-view")).toBeVisible();
  await page.getByTestId("slides-present").click();
  await expect(page.getByTestId("presentation-play")).toBeVisible();
  await expect(timer).toHaveText("0:00");
  await expect(timer).toHaveAttribute("data-state", "running");
});

test("Countdown in the full-screen show: set in Preferences, it counts down into overtime and announces it", async ({ page }) => {
  await page.clock.setFixedTime(T0);
  await withPreferences(page, async (dialog) => {
    await dialog.getByTestId("preferences-presentation-timer-mode").selectOption("countdown");
    const length = dialog.getByTestId("preferences-presentation-timer-length");
    await length.fill("2");
    await length.press("Tab"); // commits on change
    await expect(length).toHaveValue("2");
  });
  await openFixture(page, miniDeck);
  await page.getByTestId("slides-present").click();
  await revealHud(page);
  const timer = page.getByTestId("play-timer");
  await expect(timer).toHaveText("2:00");
  await expect(timer).toHaveAttribute("data-state", "running");
  await expect(timer).toHaveAttribute("data-mode", "countdown");
  await expect(timer).toHaveAttribute("aria-label", "Time remaining");

  await page.clock.setFixedTime(T0 + 75_000);
  await expect(timer).toHaveText("0:45");
  await expect(timer).toHaveAttribute("data-overtime", "false");

  await page.clock.setFixedTime(T0 + 162_000);
  await expect(timer).toHaveText("−0:42");
  await expect(timer).toHaveAttribute("data-overtime", "true");
  await expect(timer).toHaveAttribute("aria-label", "Time over");
  await expect(page.getByTestId("play-timer-status")).toHaveText("Time is up");

  await page.keyboard.press("s");
  await expect(timer).toHaveText("2:00");
  await expect(timer).toHaveAttribute("data-state", "stopped");
});

test("Presenter view: Countdown mid-show keeps the count, the audience's t pauses the presenter's timer, and the promoted audience inherits it", async ({ page }) => {
  await page.clock.setFixedTime(T0);
  const deckId = await openFixture(page, miniDeck);
  const audience = await startPresenterView(page, deckId);
  const timer = page.getByTestId("presenter-timer");
  await expect(timer).toHaveAttribute("data-state", "running");
  await expect(timer).toHaveAttribute("data-mode", "stopwatch");

  await page.clock.setFixedTime(T0 + 75_000);
  await expect(timer).toHaveText("1:15");

  // Switching to Countdown shows the time remaining of the default 20
  // minutes, without losing the count; a 1-minute talk is already over.
  await page.getByTestId("presenter-timer-mode-countdown").click();
  await expect(page.getByTestId("presenter-timer-mode-countdown")).toHaveAttribute("aria-pressed", "true");
  await expect(timer).toHaveText("18:45");
  await expect(timer).toHaveAttribute("data-state", "running");
  const length = page.getByTestId("presenter-timer-length");
  await length.fill("1");
  await length.press("Enter");
  await expect(timer).toHaveText("−0:15");
  await expect(timer).toHaveAttribute("data-overtime", "true");
  await page.getByTestId("presenter-timer-mode-stopwatch").click();
  await expect(timer).toHaveText("1:15");

  // t in the AUDIENCE window pauses the presenter's timer; the audience,
  // which follows, has no timer of its own.
  await audience.keyboard.press("t");
  await expect(timer).toHaveAttribute("data-state", "paused");
  await expect(timer).toHaveText("1:15");
  await expect(audience.getByTestId("play-timer")).toHaveCount(0);

  // The presenter window closes: the audience promotes itself to a plain
  // show (no ?follow=) and its HUD shows the inherited, still-paused timer.
  await page.clock.setFixedTime(T0 + 200_000);
  await page.close();
  await expect(audience).toHaveURL(new RegExp(`/slides/${deckId}/present/1$`));
  await revealHud(audience);
  const inherited = audience.getByTestId("play-timer");
  await expect(inherited).toHaveText("1:15");
  await expect(inherited).toHaveAttribute("data-state", "paused");
});

test("Auto-start off: Present leaves the timer stopped until t starts it", async ({ page }) => {
  await page.clock.setFixedTime(T0);
  await withPreferences(page, async (dialog) => {
    const autostart = dialog.getByTestId("preferences-presentation-timer-autostart");
    await expect(autostart).toBeChecked();
    await autostart.uncheck();
    await expect(autostart).not.toBeChecked();
  });
  await openFixture(page, miniDeck);
  await page.getByTestId("slides-present").click();
  await revealHud(page);
  const timer = page.getByTestId("play-timer");
  await expect(timer).toHaveText("0:00");
  await expect(timer).toHaveAttribute("data-state", "stopped");
  await expect(page.getByTestId("play-timer-toggle")).toHaveText("Start (T)");
  await expect(page.getByTestId("play-timer-stop")).toBeDisabled();
  await page.keyboard.press("t");
  await expect(timer).toHaveAttribute("data-state", "running");
  await expect(page.getByTestId("play-timer-toggle")).toHaveText("Pause (T)");
});
