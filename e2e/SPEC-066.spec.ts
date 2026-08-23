import { test, expect, type Page } from "@playwright/test";
import { MOD } from "./keys";

// SPEC-066 — the centralized Preferences dialog: reachable from the editor
// (File menu + Cmd/Ctrl+,) and the library, applying live, persisting under
// one key and surviving a reload. Unit tests cover the store's schema,
// fallback, migration and cross-tab sync (src/lib/preferences.spec.ts); this
// file exercises what only a real browser confirms: the menu/shortcut wiring,
// that a change survives a reload, and reset.

async function createDocument(page: Page): Promise<void> {
  await page.goto("/");
  await page.getByTestId("new-document").first().click();
  await expect(page).toHaveURL(/\/doc\/.+/);
}

async function openFileMenu(page: Page): Promise<void> {
  const title = page.getByTestId("menu-title-file");
  if ((await title.getAttribute("aria-expanded")) !== "true") {
    await title.click();
  }
  await expect(title).toHaveAttribute("aria-expanded", "true");
}

const dialog = (page: Page) => page.getByTestId("preferences-dialog");
const htmlTheme = (page: Page) =>
  page.evaluate(() => document.documentElement.getAttribute("data-theme"));

test("File → Preferences… opens the dialog in the editor (AC-1/40)", async ({ page }) => {
  await createDocument(page);
  await openFileMenu(page);
  await page.getByTestId("file-menu-preferences").click();
  await expect(dialog(page)).toBeVisible();
});

test("Cmd/Ctrl+, opens the dialog from the editor (AC-2)", async ({ page }) => {
  await createDocument(page);
  await page.keyboard.press(`${MOD}+Comma`);
  await expect(dialog(page)).toBeVisible();
});

test("the dialog is reachable from the library view (AC-39)", async ({ page }) => {
  await page.goto("/");
  await page.getByTestId("library-preferences").click();
  await expect(dialog(page)).toBeVisible();
  // …and via the shortcut there too, since AppShell hosts it in both views.
  await page.getByTestId("preferences-close").click();
  await expect(dialog(page)).toHaveCount(0);
  await page.keyboard.press(`${MOD}+Comma`);
  await expect(dialog(page)).toBeVisible();
});

test("changing a setting applies live and the Close button just closes (AC-7/8)", async ({
  page,
}) => {
  await createDocument(page);
  await page.keyboard.press(`${MOD}+Comma`);
  await expect(dialog(page)).toBeVisible();

  // Live-apply: choosing Dark sets data-theme immediately, with no OK step.
  await page.getByTestId("preferences-theme").selectOption("dark");
  expect(await htmlTheme(page)).toBe("dark");

  // Close just closes — the choice already applied stays applied.
  await page.getByTestId("preferences-close").click();
  await expect(dialog(page)).toHaveCount(0);
  expect(await htmlTheme(page)).toBe("dark");
});

test("theme, measurement unit and default font survive a reload (AC-41/42/43)", async ({
  page,
}) => {
  await createDocument(page);
  await page.keyboard.press(`${MOD}+Comma`);
  await expect(dialog(page)).toBeVisible();

  await page.getByTestId("preferences-theme").selectOption("dark");
  await page.getByTestId("preferences-measurement-unit").selectOption("in");
  await page.getByTestId("preferences-default-font").selectOption("Georgia");
  await page.getByTestId("preferences-close").click();

  // Everything lives under one key.
  const stored = await page.evaluate(() =>
    JSON.parse(localStorage.getItem("docu-edit:preferences") || "{}"),
  );
  expect(stored.theme).toBe("dark");
  expect(stored.measurementUnit).toBe("in");
  expect(stored.defaultFontFamily).toBe("Georgia");

  await page.reload();
  // The theme re-applies before the app owns it (index.html no-flash script).
  expect(await htmlTheme(page)).toBe("dark");

  // Wait for the editor (and its keydown listener) to be ready after reload.
  await expect(page.getByTestId("menu-title-file")).toBeVisible();
  await page.keyboard.press(`${MOD}+Comma`);
  await expect(dialog(page)).toBeVisible();
  await expect(page.getByTestId("preferences-theme")).toHaveValue("dark");
  await expect(page.getByTestId("preferences-measurement-unit")).toHaveValue("in");
  await expect(page.getByTestId("preferences-default-font")).toHaveValue("Georgia");
});

test("Reset all preferences restores defaults after a confirmation (AC-28)", async ({ page }) => {
  await createDocument(page);
  await page.keyboard.press(`${MOD}+Comma`);
  await expect(dialog(page)).toBeVisible();

  await page.getByTestId("preferences-theme").selectOption("dark");
  await page.getByTestId("preferences-measurement-unit").selectOption("in");
  expect(await htmlTheme(page)).toBe("dark");

  // Reset is gated behind a confirmation.
  await page.getByTestId("preferences-reset").click();
  await page.getByTestId("preferences-reset-confirm").click();

  await expect(page.getByTestId("preferences-theme")).toHaveValue("system");
  await expect(page.getByTestId("preferences-measurement-unit")).toHaveValue("cm");
  expect(await htmlTheme(page)).toBeNull();
});
