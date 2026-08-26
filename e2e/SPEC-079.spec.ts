import { test, expect, type Locator, type Page } from "@playwright/test";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

// SPEC-079 — the WYSIWYG slide renderer at /slides/:id. Opening the fixture
// deck lands on a stage that renders slide 1 (title larger than body), a
// filmstrip that lists every slide, keyboard navigation, and no more of the
// preparation milestone's read-only note. The fixtures are generated from
// src/lib/presentation/fixtures/build.ts, so these assertions and the unit
// suite describe the same deck.
const here = dirname(fileURLToPath(import.meta.url));
const pptxFixture = join(here, "fixtures", "mini-deck.pptx");

async function openFixture(page: Page): Promise<void> {
  await page.goto("/");
  await page.getByTestId("open-document-input").setInputFiles(pptxFixture);
  await expect(page).toHaveURL(/\/slides\/.+/);
  await expect(page.getByTestId("presentation-view")).toBeVisible();
}

/** The rendered font size (px) of a rendered text run — relative comparisons
 * only (Mac and CI fonts differ). */
async function fontSizePx(locator: Locator): Promise<number> {
  return locator.evaluate((el) => parseFloat(getComputedStyle(el as HTMLElement).fontSize));
}

test("the stage renders slide 1 with its title larger than its body text", async ({ page }) => {
  await openFixture(page);
  const stage = page.getByTestId("stage-workspace");
  await expect(stage.getByText("Quarterly Review")).toBeVisible();

  const titleSize = await fontSizePx(stage.getByText("Quarterly Review"));
  const bodySize = await fontSizePx(stage.getByText("FY26 results and outlook"));
  expect(titleSize).toBeGreaterThan(bodySize);

  // The text frame must occupy a real, legible box — not collapse to a
  // zero-width column (the failure mode when a placeholder had no geometry).
  const titleBox = await stage.getByTestId("slide-frame").first().boundingBox();
  expect(titleBox?.width ?? 0).toBeGreaterThan(100);
});

test("the filmstrip lists the fixture's slide count", async ({ page }) => {
  await openFixture(page);
  await expect(page.getByTestId("filmstrip-slide")).toHaveCount(4);
  await expect(page.getByTestId("slide-filmstrip")).toHaveAttribute("role", "listbox");
});

test("keyboard-navigating to slide 2 shows its body line on the stage", async ({ page }) => {
  await openFixture(page);
  const stage = page.getByTestId("stage-workspace");
  await expect(stage.getByText("Ship presentations MVP")).toHaveCount(0);

  await page.getByTestId("slide-filmstrip").focus();
  await page.keyboard.press("ArrowDown");
  await expect(stage.getByText("Ship presentations MVP")).toBeVisible();
});

test("the preparation milestone's read-only note is gone", async ({ page }) => {
  await openFixture(page);
  await expect(page.getByTestId("presentation-note")).toHaveCount(0);
});
