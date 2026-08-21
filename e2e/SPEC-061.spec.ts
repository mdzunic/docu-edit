import { test, expect, type Page } from "@playwright/test";

// SPEC-061 — the redesigned document toolbar: fixed bands, priority overflow
// measured from real width, a single roving tab stop, the removed duplicate
// track-changes toggle, and the contextual text-selection toolbar. These
// checks measure the inner .doc-toolbar-strip (not the outer [data-testid=
// toolbar] wrapper) so a clipped-but-not-scrolled band cannot pass unseen.

async function newDoc(page: Page): Promise<void> {
  await page.goto("/");
  await page.getByTestId("new-document").click();
  await expect(page).toHaveURL(/\/doc\/.+/);
}

const strip = (page: Page) => page.getByTestId("doc-toolbar-strip");
const overflowButton = (page: Page) => page.getByTestId("toolbar-overflow-button");

test("controls are grouped into fixed bands separated by rules (AC-3, AC-4)", async ({ page }) => {
  await page.setViewportSize({ width: 1400, height: 900 });
  await newDoc(page);
  // The eight bands render in fixed left-to-right order with the character and
  // paragraph bands present.
  for (const band of [1, 2, 3, 4, 5, 6]) {
    await expect(strip(page).locator(`.band[data-band="${band}"]`)).toBeVisible();
  }
  // Bands past the first carry the 1px subtle separator rule.
  const hasRule = await strip(page)
    .locator('.band[data-band="4"]')
    .evaluate((el) => getComputedStyle(el).borderInlineStartWidth);
  expect(hasRule).toBe("1px");
});

test("the strip never scrolls horizontally (AC-13)", async ({ page }) => {
  await page.setViewportSize({ width: 800, height: 800 });
  await newDoc(page);
  const scrolls = await strip(page).evaluate((el) => el.scrollWidth > el.clientWidth + 1);
  expect(scrolls).toBe(false);
});

test("at 1024px the character and paragraph bands show without overflow (AC-36)", async ({ page }) => {
  await page.setViewportSize({ width: 1024, height: 900 });
  await newDoc(page);
  // Give the settle-timer re-measures time to run after the UI font loads.
  await page.waitForTimeout(1600);
  await expect(strip(page).locator('.band[data-band="4"]')).toBeVisible(); // Character
  await expect(strip(page).locator('.band[data-band="6"]')).toBeVisible(); // Paragraph
  const clipped = await strip(page).evaluate((el) => el.scrollWidth > el.clientWidth + 1);
  expect(clipped).toBe(false);
});

test("narrowing to 1024px keeps the paragraph band without over-collapsing (AC-36)", async ({ page }) => {
  // Start wide (nothing overflows, so the overflow button is display:none and
  // its width is not measured inline), then narrow to 1024px. The resize path
  // must re-plan against the real overflow-button width, not a stale default,
  // or band 6 collapses even though it fits.
  await page.setViewportSize({ width: 1440, height: 900 });
  await newDoc(page);
  await page.waitForTimeout(1600);
  await expect(overflowButton(page)).toBeHidden();
  await page.setViewportSize({ width: 1024, height: 900 });
  await expect(strip(page).locator('.band[data-band="4"]')).toBeVisible(); // Character
  await expect(strip(page).locator('.band[data-band="6"]')).toBeVisible(); // Paragraph
  const clipped = await strip(page).evaluate((el) => el.scrollWidth > el.clientWidth + 1);
  expect(clipped).toBe(false);
});

test("narrow width collapses low-priority bands into the overflow menu (AC-12, AC-38)", async ({ page }) => {
  await page.setViewportSize({ width: 720, height: 800 });
  await newDoc(page);
  await expect(overflowButton(page)).toBeVisible();
  // A collapsed control (pagination, band 8) is not in the strip but is found
  // by the same testid once the overflow menu opens.
  await overflowButton(page).click();
  await expect(page.getByTestId("toolbar-overflow-menu")).toBeVisible();
  await expect(page.getByTestId("paragraph-pagination-button")).toBeVisible();
});

test("opening a side pane collapses the toolbar like narrowing the window (AC-15)", async ({ page }) => {
  await page.setViewportSize({ width: 1340, height: 900 });
  await newDoc(page);
  await page.waitForTimeout(1600);
  await expect(overflowButton(page)).toBeHidden();
  // Open the Styles pane (a right-docked pane) — the reserved width collapses
  // the strip exactly as narrowing would.
  await page.getByTestId("menu-title-view").click();
  await page.getByTestId("view-menu-styles-pane").click();
  await expect(overflowButton(page)).toBeVisible();
});

test("the strip is one roving tab stop ending on the overflow button (AC-32, AC-34)", async ({ page }) => {
  await page.setViewportSize({ width: 720, height: 800 });
  await newDoc(page);
  await expect(overflowButton(page)).toBeVisible();
  // Focus the first enabled strip control (undo is disabled on a new doc),
  // then End jumps to the overflow button.
  await strip(page).locator("button:not([disabled]), input, select").first().focus();
  await page.keyboard.press("End");
  const onOverflow = await overflowButton(page).evaluate((el) => el === document.activeElement);
  expect(onOverflow).toBe(true);
});

test("exactly one strip control is a tab stop, even after the size steppers enable (AC-32)", async ({ page }) => {
  await page.setViewportSize({ width: 1400, height: 900 });
  await newDoc(page);
  // A caret resolves a font size, enabling the −/+ steppers that mount disabled.
  await page.getByTestId("document-content").click();
  await expect(page.getByTestId("font-size-decrease")).toBeEnabled();
  const tabStops = await strip(page).evaluate(
    (el) => el.querySelectorAll('[tabindex="0"]').length,
  );
  expect(tabStops).toBe(1);
});

test("arrow roving reaches the font-size − stepper between the select and the field (AC-31)", async ({ page }) => {
  await page.setViewportSize({ width: 1400, height: 900 });
  await newDoc(page);
  await page.getByTestId("document-content").click();
  await expect(page.getByTestId("font-size-decrease")).toBeEnabled();
  await page.getByTestId("style-picker-button").focus();
  await page.keyboard.press("ArrowRight");
  await expect(page.getByTestId("font-family-select")).toBeFocused();
  await page.keyboard.press("ArrowRight");
  await expect(page.getByTestId("font-size-decrease")).toBeFocused();
  await page.keyboard.press("ArrowRight");
  await expect(page.getByTestId("font-size-input")).toBeFocused();
  // The field owns Left/Right for its caret only until the boundary; entering it
  // rightward drops the caret at the end so one more ArrowRight ejects it.
  await page.keyboard.press("ArrowRight");
  await expect(page.getByTestId("font-size-increase")).toBeFocused();
});

test("End reaches the last visible control when nothing overflows (AC-33)", async ({ page }) => {
  await page.setViewportSize({ width: 1440, height: 900 });
  await newDoc(page);
  await page.waitForTimeout(1600);
  await expect(overflowButton(page)).toBeHidden();
  await page.getByTestId("style-picker-button").focus();
  await page.keyboard.press("End");
  // Focus lands on a real, visible strip control — not the display:none overflow
  // button and not stuck on the first control.
  const landed = await strip(page).evaluate((el) => {
    const active = document.activeElement;
    if (!(active instanceof HTMLElement) || !el.contains(active)) {
      return { ok: false, testid: null as string | null };
    }
    return {
      ok: active.offsetParent !== null,
      testid: active.getAttribute("data-testid"),
    };
  });
  expect(landed.ok).toBe(true);
  expect(landed.testid).not.toBe("style-picker-button");
  expect(landed.testid).not.toBe("toolbar-overflow-button");
});

test("the duplicate track-changes toolbar button is gone; the status-bar one remains (AC-19, AC-20)", async ({ page }) => {
  await newDoc(page);
  await expect(page.getByTestId("toolbar-track-changes-toggle")).toHaveCount(0);
  await expect(page.getByTestId("status-bar-track-changes")).toBeVisible();
});

test("the font-size control is one unit with a common-sizes dropdown (AC-21)", async ({ page }) => {
  await page.setViewportSize({ width: 1400, height: 900 });
  await newDoc(page);
  await page.getByTestId("font-size-menu-button").click();
  const menu = page.getByTestId("font-size-menu");
  await expect(menu).toBeVisible();
  await menu.getByTestId("font-size-option-24").click();
  await expect(page.getByTestId("font-size-input")).toHaveValue("24");
});

test("selecting body text raises a contextual toolbar; Escape hides it, keeping the selection (AC-24, AC-26, AC-30)", async ({ page }) => {
  await page.setViewportSize({ width: 1400, height: 900 });
  await newDoc(page);
  const content = page.getByTestId("document-content");
  await content.click();
  await page.keyboard.type("Contextual selection toolbar");
  // Collapsed caret: no toolbar (AC-26).
  await expect(page.getByTestId("text-selection-toolbar")).toHaveCount(0);
  await page.keyboard.press("Home");
  for (let i = 0; i < 10; i++) {
    await page.keyboard.press("Shift+ArrowRight");
  }
  await expect(page.getByTestId("text-selection-toolbar")).toBeVisible();
  await expect(page.getByTestId("selection-bold")).toBeVisible();
  await expect(page.getByTestId("selection-comment")).toBeVisible();
  await page.keyboard.press("Escape");
  await expect(page.getByTestId("text-selection-toolbar")).toBeHidden();
  const stillSelected = await page.evaluate(() => (window.getSelection()?.toString() ?? "").length > 0);
  expect(stillSelected).toBe(true);
});

test("the text-selection toolbar is suppressed while a menu is open (AC-28)", async ({ page }) => {
  await page.setViewportSize({ width: 1400, height: 900 });
  await newDoc(page);
  const content = page.getByTestId("document-content");
  await content.click();
  await page.keyboard.type("Selection under a menu");
  await page.keyboard.press("Home");
  for (let i = 0; i < 9; i++) {
    await page.keyboard.press("Shift+ArrowRight");
  }
  await expect(page.getByTestId("text-selection-toolbar")).toBeVisible();
  await page.getByTestId("menu-title-insert").click();
  await expect(page.getByTestId("text-selection-toolbar")).toBeHidden();
});
