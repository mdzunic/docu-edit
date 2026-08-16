import { test, expect, type Page } from "@playwright/test";

// SPEC-048 — font fallback & embedding. "Tools -> Fonts" opens a panel that
// lists every family the document uses, each marked available / substituted
// (naming the twin) / missing. The substitution maths, embedded-font
// extraction and subset embedding are proven exhaustively at the unit level
// (src/lib/fonts/*.spec.ts); this exercises the panel/menu wiring end to end.

async function createDocument(page: Page): Promise<void> {
  await page.goto("/");
  await page.getByTestId("new-document").click();
  await expect(page).toHaveURL(/\/doc\/.+/);
}

const pane = (page: Page) => page.getByTestId("fonts-pane");

async function openFontsPane(page: Page): Promise<void> {
  await page.getByTestId("menu-title-tools").click();
  await page.getByTestId("tools-menu-fonts").click();
  await expect(pane(page)).toBeVisible();
}

test("the Fonts panel lists document fonts with a status (AC-7/8/10)", async ({ page }) => {
  await createDocument(page);
  await page.getByTestId("document-content").click();
  await page.keyboard.type("Hello fonts");

  await openFontsPane(page);

  // The panel summarises how many fonts are in use and lists them.
  await expect(page.getByTestId("fonts-summary")).toContainText(/font/i);
  const items = pane(page).locator(".fonts-item");
  await expect(items.first()).toBeVisible();

  // The document default (Calibri) appears with a resolved status; on a host
  // without Calibri it renders on its bundled twin, Carlito.
  const calibriStatus = page.getByTestId("fonts-status-Calibri");
  await expect(calibriStatus).toBeVisible();
  await expect(calibriStatus).toHaveText(/Available|Substituted|Missing/);

  // Closing the panel hides it again.
  await page.getByTestId("fonts-pane-close").click();
  await expect(pane(page)).toBeHidden();
});
