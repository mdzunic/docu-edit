import { test, expect, type Page, type Download } from "@playwright/test";
import { readFileSync } from "node:fs";

// SPEC-059 — the three-state theme (system / light / dark), its dark token
// palette, the dim-page opt-in and the guarantee that neither ever touches
// saved files or printed/PDF output. Unit tests cover the palette contrast
// (tokens.spec.ts over both palettes), the persistence guard (themePrefs) and
// the composable (useTheme); this file exercises what only a real browser
// confirms: the menu wiring, the data-theme attribute that gates the CSS, that
// every chrome surface resolves to a dark-palette colour, and that a PDF is
// byte-identical whatever the theme.

async function createDocument(page: Page): Promise<void> {
  await page.goto("/");
  await page.getByTestId("new-document").first().click();
  await expect(page).toHaveURL(/\/doc\/.+/);
}

/** Opens a top-level menu, tolerating one already being open (clicking an open
 * menu's title toggles it shut). */
async function openMenu(page: Page, id: string): Promise<void> {
  const title = page.getByTestId(`menu-title-${id}`);
  if ((await title.getAttribute("aria-expanded")) !== "true") {
    await title.click();
  }
  await expect(title).toHaveAttribute("aria-expanded", "true");
}

/** Opens View → Appearance and leaves the submenu flyout open — idempotent, so
 * a test may call it whether or not the menu/submenu is already showing. */
async function openAppearance(page: Page): Promise<void> {
  await openMenu(page, "view");
  const flyout = page.getByTestId("view-menu-appearance-flyout");
  if (!(await flyout.isVisible())) {
    await page.getByTestId("view-menu-appearance").click();
  }
  await expect(flyout).toBeVisible();
}

async function chooseTheme(page: Page, which: "system" | "light" | "dark"): Promise<void> {
  await openAppearance(page);
  await page.getByTestId(`appearance-${which}`).click();
}

const htmlTheme = (page: Page) =>
  page.evaluate(() => document.documentElement.getAttribute("data-theme"));

test("View → Appearance is a submenu of three radios showing the current choice (AC-1/48)", async ({
  page,
}) => {
  await createDocument(page);
  await openAppearance(page);

  const system = page.getByTestId("appearance-system");
  const light = page.getByTestId("appearance-light");
  const dark = page.getByTestId("appearance-dark");
  for (const item of [system, light, dark]) {
    await expect(item).toHaveAttribute("role", "menuitemradio");
  }
  // Default is system, so it reads as the selected radio.
  await expect(system).toHaveAttribute("aria-checked", "true");
  await expect(dark).toHaveAttribute("aria-checked", "false");

  await dark.click();
  await openAppearance(page);
  await expect(page.getByTestId("appearance-dark")).toHaveAttribute("aria-checked", "true");
  await expect(page.getByTestId("appearance-system")).toHaveAttribute("aria-checked", "false");
});

test("choosing a theme sets data-theme on <html>, and system removes it (AC-4/6)", async ({
  page,
}) => {
  await createDocument(page);

  await chooseTheme(page, "dark");
  expect(await htmlTheme(page)).toBe("dark");

  await chooseTheme(page, "light");
  expect(await htmlTheme(page)).toBe("light");

  await chooseTheme(page, "system");
  expect(await htmlTheme(page)).toBeNull();
});

test("the preference persists under docu-edit:theme and survives a reload (AC-2/5/54/57)", async ({
  page,
}) => {
  await createDocument(page);
  await chooseTheme(page, "dark");
  expect(await page.evaluate(() => localStorage.getItem("docu-edit:theme"))).toBe("dark");

  // A reload re-applies before the app owns the theme: the no-flash script in
  // index.html sets data-theme straight from storage, so it is present the
  // moment the document element exists.
  await page.reload();
  expect(await htmlTheme(page)).toBe("dark");
});

test("Dim page is default-off, dark-only, screen-only and persists (AC-13/14/49/55/70/71/72)", async ({
  page,
}) => {
  await createDocument(page);

  // Light (the default resolved theme in the test browser): the dim item is
  // present but disabled (AC-49/70).
  await openAppearance(page);
  const dim = page.getByTestId("appearance-dim-page");
  await expect(dim).toHaveAttribute("role", "menuitemcheckbox");
  await expect(dim).toHaveAttribute("aria-checked", "false"); // default off (AC-14)
  await expect(dim).toBeDisabled();
  await expect(dim).toContainText("Screen only"); // marked screen-only (AC-71)

  // Switch to dark: the item enables and toggles on, persisting under its key.
  await chooseTheme(page, "dark");
  await openAppearance(page);
  await expect(page.getByTestId("appearance-dim-page")).toBeEnabled();
  await page.getByTestId("appearance-dim-page").click();
  expect(await page.evaluate(() => localStorage.getItem("docu-edit:dim-page"))).toBe("1");
  expect(
    await page.evaluate(() => document.documentElement.getAttribute("data-dim-page")),
  ).toBe("1");
});

// The colour tokens whose values must be the only ones any chrome surface
// resolves to in dark theme. Document content (white pages, model colours) is
// out of scope by design (AC-12), so this asserts over chrome only.
const PALETTE_TOKENS = [
  "--surface-canvas", "--surface-base", "--surface-raised", "--surface-sunken",
  "--surface-brand", "--surface-selected", "--surface-overlay", "--text-primary",
  "--text-secondary", "--text-disabled", "--text-on-brand", "--text-accent",
  "--text-danger", "--border-subtle", "--border-default", "--border-strong",
  "--accent", "--accent-hover", "--accent-subtle", "--danger", "--success",
  "--warning", "--state-hover", "--state-active", "--focus-ring",
];

/** In-page: builds the set of allowed computed colours from the resolved dark
 * tokens, then reports any surface whose background sits outside it. */
async function surfaceViolations(page: Page, testids: string[]): Promise<string[]> {
  return page.evaluate(
    ({ tokens, ids }) => {
      const cs = getComputedStyle(document.documentElement);
      const probe = document.createElement("span");
      document.body.appendChild(probe);
      const allowed = new Set<string>(["rgba(0, 0, 0, 0)", "transparent"]);
      for (const name of tokens) {
        probe.style.color = "";
        probe.style.color = cs.getPropertyValue(name).trim();
        allowed.add(getComputedStyle(probe).color);
      }
      probe.remove();
      const bad: string[] = [];
      for (const id of ids) {
        const el = document.querySelector(`[data-testid="${id}"]`);
        if (!el) {
          bad.push(`${id}: missing`);
          continue;
        }
        const bg = getComputedStyle(el).backgroundColor;
        if (!allowed.has(bg)) {
          bad.push(`${id}: ${bg}`);
        }
      }
      return bad;
    },
    { tokens: PALETTE_TOKENS, ids: testids },
  );
}

test("every top-level surface resolves to a dark-palette colour (AC-47)", async ({ page }) => {
  // The library first (AC-52): themed with no menu bar of its own.
  await page.goto("/");
  await page.evaluate(() => localStorage.setItem("docu-edit:theme", "dark"));
  await page.reload();
  await expect(page.getByTestId("app-bar")).toBeVisible();
  expect(await htmlTheme(page)).toBe("dark");
  expect(await surfaceViolations(page, ["app-bar"])).toEqual([]);

  // The editor, its chrome and its status bar.
  await page.getByTestId("new-document").first().click();
  await expect(page).toHaveURL(/\/doc\/.+/);
  expect(await htmlTheme(page)).toBe("dark");
  expect(await surfaceViolations(page, ["app-bar", "menu-bar", "toolbar", "canvas", "status-bar"])).toEqual([]);

  // Each of the seven panes, opened from its own menu.
  const panes: Array<[string, string, string]> = [
    ["view", "view-menu-outline", "outline-pane"],
    ["view", "view-menu-styles-pane", "styles-pane"],
    ["view", "view-menu-comments-pane", "comments-pane"],
    ["view", "view-menu-track-changes-pane", "track-changes-pane"],
    ["file", "file-menu-version-history", "version-history-pane"],
    ["tools", "tools-menu-accessibility", "accessibility-pane"],
    ["tools", "tools-menu-fonts", "fonts-pane"],
  ];
  for (const [menu, item, paneId] of panes) {
    const pane = page.getByTestId(paneId);
    // Some panes (the outline pane) open by default; only toggle a closed one.
    if (!(await pane.isVisible())) {
      await openMenu(page, menu);
      await page.getByTestId(item).click();
    }
    await expect(pane).toBeVisible();
    expect(await surfaceViolations(page, [paneId])).toEqual([]);
  }

  // A dialog and its scrim (AC-42).
  await openMenu(page, "file");
  await page.getByTestId("page-setup-button").click();
  await expect(page.getByTestId("page-setup-dialog")).toBeVisible();
  expect(await surfaceViolations(page, ["page-setup-dialog"])).toEqual([]);
});

async function downloadedBuffer(download: Download): Promise<Buffer> {
  const path = await download.path();
  expect(path).toBeTruthy();
  return readFileSync(path);
}

async function exportPdf(page: Page): Promise<Buffer> {
  const downloadPromise = page.waitForEvent("download");
  await page.getByTestId("download-menu").click();
  await page.getByTestId("download-pdf").click();
  return downloadedBuffer(await downloadPromise);
}

test("the exported PDF is byte-identical in light and dark theme (AC-16/29..33)", async ({
  page,
}) => {
  await createDocument(page);

  await chooseTheme(page, "light");
  const lightPdf = await exportPdf(page);

  await chooseTheme(page, "dark");
  const darkPdf = await exportPdf(page);

  expect(darkPdf.equals(lightPdf)).toBe(true);
});
