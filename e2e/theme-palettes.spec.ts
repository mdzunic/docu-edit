import { test, expect, type Page } from "@playwright/test";
import { MOD } from "./keys";

// Design themes (the library-modernist handoff): Modernist is the base :root
// palette, and the Preferences dialog's "Color theme" select offers Nocturne,
// Classic and the accent options as value-only [data-palette="…"]
// redefinitions. Unit tests cover the preference schema (preferences.spec.ts),
// the data-palette application (useTheme.spec.ts) and the WCAG pairs over
// every palette (tokens.spec.ts); this file exercises what only a real
// browser confirms: the dialog wiring, the CSS actually re-resolving, the
// no-flash first paint after a reload, and the interplay with the dark theme.

const htmlPalette = (page: Page) =>
  page.evaluate(() => document.documentElement.getAttribute("data-palette"));

const token = (page: Page, name: string) =>
  page.evaluate(
    (t) => getComputedStyle(document.documentElement).getPropertyValue(t).trim(),
    name,
  );

async function openPreferences(page: Page): Promise<void> {
  await page.getByTestId("library-preferences").click();
  await expect(page.getByTestId("preferences-dialog")).toBeVisible();
}

test("the default is Modernist: no data-palette attribute, brick accent, Archivo-first stack", async ({
  page,
}) => {
  await page.goto("/");
  await expect(page.getByTestId("app-bar")).toBeVisible();
  expect(await htmlPalette(page)).toBeNull();
  expect(await token(page, "--accent")).toBe("#9c4a34");
  expect(await token(page, "--font-ui")).toContain("Archivo");
  await openPreferences(page);
  await expect(page.getByTestId("preferences-palette-modernist")).toHaveAttribute(
    "aria-pressed",
    "true",
  );
});

test("the Color theme select offers every declared palette", async ({ page }) => {
  await page.goto("/");
  await openPreferences(page);
  // Every palette renders as a chip with its dominant-colour swatch square.
  const values = await page
    .getByTestId("preferences-palette")
    .locator(".palette-chip")
    .evaluateAll((chips) =>
      chips.map((chip) => chip.getAttribute("data-testid")!.replace("preferences-palette-", "")),
    );
  expect(values).toEqual([
    "modernist",
    "nocturne",
    "classic",
    "petrol",
    "plum",
    "ink",
    "signal",
    "pine",
    "ochre",
    "graphite",
  ]);
});

test("choosing Nocturne applies live, persists, and paints flash-free on reload", async ({
  page,
}) => {
  await page.goto("/");
  const modernistBase = await token(page, "--surface-base");

  await openPreferences(page);
  await page.getByTestId("preferences-palette-nocturne").click();
  // Live-apply: the attribute and the resolved tokens change with the dialog
  // still open — there is no OK step.
  expect(await htmlPalette(page)).toBe("nocturne");
  expect(await token(page, "--surface-base")).not.toBe(modernistBase);
  expect(await token(page, "--accent")).toBe("#6f5fc4");
  expect(await token(page, "--font-ui")).toContain("Inter");

  // Persisted under the one preferences key…
  expect(
    await page.evaluate(
      () => JSON.parse(localStorage.getItem("docu-edit:preferences") || "{}").palette,
    ),
  ).toBe("nocturne");

  // …and re-applied by the no-flash script before the app bundle runs.
  await page.reload();
  expect(await htmlPalette(page)).toBe("nocturne");
  expect(await token(page, "--accent")).toBe("#6f5fc4");
});

test("a palette composes with the dark theme: each palette's own dark values win", async ({
  page,
}) => {
  await page.goto("/");
  await openPreferences(page);
  await page.getByTestId("preferences-palette-petrol").click();
  await page.getByTestId("preferences-theme-dark").click();

  expect(await htmlPalette(page)).toBe("petrol");
  expect(
    await page.evaluate(() => document.documentElement.getAttribute("data-theme")),
  ).toBe("dark");
  // Petrol's dark accent, not its light one and not Modernist's dark brick.
  expect(await token(page, "--accent")).toBe("#297f91");
  // The ground is Modernist's dark ground — accent palettes swap only the
  // accent-facing tokens.
  expect(await token(page, "--surface-base")).toBe("#2d2b2b");

  // Back to light: the palette's light accent returns.
  await page.getByTestId("preferences-theme-light").click();
  expect(await token(page, "--accent")).toBe("#0e6f83");
});

test("the palette follows preferences reset and the editor is themed too", async ({ page }) => {
  await page.goto("/");
  await openPreferences(page);
  await page.getByTestId("preferences-palette-plum").click();
  expect(await htmlPalette(page)).toBe("plum");
  await page.getByTestId("preferences-close").click();

  // The editor inherits the palette (the tokens live on :root).
  await page.getByTestId("new-document").first().click();
  await expect(page).toHaveURL(/\/doc\/.+/);
  expect(await htmlPalette(page)).toBe("plum");
  expect(await token(page, "--accent")).toBe("#7a2d59");

  // Reset all preferences restores the default palette and drops the attribute.
  await page.keyboard.press(`${MOD}+Comma`);
  await expect(page.getByTestId("preferences-dialog")).toBeVisible();
  await page.getByTestId("preferences-reset").click();
  await page.getByTestId("preferences-reset-confirm").click();
  await expect(page.getByTestId("preferences-palette-modernist")).toHaveAttribute(
    "aria-pressed",
    "true",
  );
  expect(await htmlPalette(page)).toBeNull();
  expect(await token(page, "--accent")).toBe("#9c4a34");
});
