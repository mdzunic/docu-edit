import { type Page } from "@playwright/test";

// Shared helpers for the slides editor's chrome.

/** Open the ribbon's Insert tab (SPEC-061 Toolbar Redesign, option 2c).
 * The deck editor's insert/arrange strip used to be a permanently visible bar
 * stacked under the formatting toolbar; it now lives behind this tab, so a
 * test that reaches for `insert-shape`, `arrange-menu`, `align-menu`,
 * `group-button`, `style-button` or `slides-position-toggle` opens the tab
 * first, exactly as a user does. Calling it twice is harmless. */
export async function openInsertTab(page: Page): Promise<void> {
  await page.getByTestId("slides-tab-insert").click();
}

/** Open the ribbon's Home tab — the formatting strip (font, B/I/U, colour,
 * alignment, lists). Home is selected on load, so this is only needed after
 * a test has switched to Insert. */
export async function openHomeTab(page: Page): Promise<void> {
  await page.getByTestId("slides-tab-home").click();
}
