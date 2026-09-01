import { test, expect, type Page } from "@playwright/test";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

// SPEC-094 — deck themes. Open the themed fixture → Design marks the current
// theme card → applying a built-in recolours the accent-filled shape and the
// themed (dk2) title while the explicitly-coloured run stays put → picking the
// accent1 theme swatch in the style popover stores a REFERENCE, so after
// download → reopen the shape still tracks accent1 (a second apply moves it
// again), the literal run is unmoved, and the panel names the applied theme.
const here = dirname(fileURLToPath(import.meta.url));
const PPTX = join(here, "fixtures", "spec094-themes.pptx");

// Office vs Facet values the assertions pin (core/theme BUILT_IN_THEMES).
const OFFICE_ACCENT1 = "#4472c4";
const FACET_ACCENT1 = "#90c226";
const ION_ACCENT1 = "#b01513";
const OFFICE_DK2 = "rgb(68, 84, 106)"; // #44546a
const FACET_DK2 = "rgb(44, 60, 67)"; // #2c3c43
const LITERAL_RED = "rgb(255, 0, 0)";

async function openFixture(page: Page, file: string): Promise<void> {
  await page.goto("/");
  await page.getByTestId("open-document-input").setInputFiles(file);
  await expect(page).toHaveURL(/\/slides\/.+/);
  await expect(page.getByTestId("presentation-view")).toBeVisible();
}

function stageFrames(page: Page) {
  return page.getByTestId("stage-workspace").getByTestId("slide-frame");
}

/** The centre pixel of the Nth shape canvas on the main stage, as `#rrggbb` —
 * the only way to observe a canvas-painted fill. */
async function shapeFill(page: Page, shapeIndex: number): Promise<string> {
  return page.evaluate((index) => {
    const canvases = document.querySelectorAll(
      '[data-testid="stage-workspace"] [data-testid="slide-frame"] canvas.shape-canvas',
    );
    const canvas = canvases[index] as HTMLCanvasElement;
    const ctx = canvas.getContext("2d")!;
    const d = ctx.getImageData(Math.floor(canvas.width / 2), Math.floor(canvas.height / 2), 1, 1).data;
    return `#${[d[0], d[1], d[2]].map((v) => v.toString(16).padStart(2, "0")).join("")}`;
  }, shapeIndex);
}

async function applyBuiltIn(page: Page, name: string): Promise<void> {
  const panel = page.getByTestId("slides-theme-panel");
  await panel.locator('[data-testid="theme-card"]:not([data-current])', { hasText: name }).click();
  await page.getByTestId("theme-apply").click();
}

async function downloadTo(page: Page, name: string): Promise<string> {
  const downloadPromise = page.waitForEvent("download");
  await page.getByTestId("presentation-download").click();
  const download = await downloadPromise;
  const filePath = join(tmpdir(), name);
  await download.saveAs(filePath);
  return filePath;
}

test("Design panel: the current theme card is marked and applying a built-in recolours refs only", async ({ page }) => {
  await openFixture(page, PPTX);

  // AC-56: Design → the current theme card is first and marked.
  await page.getByTestId("slides-theme-toggle").click();
  const panel = page.getByTestId("slides-theme-panel");
  await expect(panel).toBeVisible();
  const cards = panel.getByTestId("theme-card");
  await expect(cards.first()).toHaveAttribute("data-current", "true");
  await expect(cards.first()).toHaveClass(/theme-current/);
  await expect(cards.first()).toContainText("Office");
  await expect(cards).toHaveCount(7); // current + six built-ins

  // The themed title renders its dk2 reference; the literal run its red.
  const title = stageFrames(page).nth(0).locator("span").last();
  await expect(title).toHaveCSS("color", OFFICE_DK2);
  expect(await shapeFill(page, 0)).toBe(OFFICE_ACCENT1); // the accent1 rect

  // AC-57: apply Facet → the accent shape and the themed title move…
  await applyBuiltIn(page, "Facet");
  await expect(title).toHaveCSS("color", FACET_DK2);
  await expect.poll(() => shapeFill(page, 0)).toBe(FACET_ACCENT1);
  // …while the explicitly-coloured run does not.
  const literalRun = stageFrames(page).nth(4).locator("span").last();
  await expect(literalRun).toHaveCSS("color", LITERAL_RED);
  // The panel names the applied theme.
  await expect(page.getByTestId("theme-current-name")).toContainText("Facet");
});

test("a theme swatch stores the reference and survives download → reopen (AC-58/59)", async ({ page }) => {
  await openFixture(page, PPTX);

  // Select the ellipse (shape canvas #1, frame index 3) and open its style
  // popover; pick the plain accent1 theme swatch for its fill.
  await stageFrames(page).nth(3).click();
  await page.getByTestId("style-button").click();
  await expect(page.getByTestId("frame-style-popover")).toBeVisible();
  await page
    .getByTestId("frame-style-popover")
    .locator('[data-testid="style-theme-swatch"][data-theme-slot="accent1"]')
    .first()
    .click();
  await expect.poll(() => shapeFill(page, 1)).toBe(OFFICE_ACCENT1);

  // Close the popover (its backdrop would swallow the download click).
  await page.locator(".popover-backdrop").click({ position: { x: 5, y: 5 } });
  await expect(page.getByTestId("frame-style-popover")).toBeHidden();

  // Download → reopen.
  const filePath = await downloadTo(page, "spec094-roundtrip.pptx");
  await openFixture(page, filePath);

  // The shape still tracks accent1 under the (unchanged) Office theme…
  await expect.poll(() => shapeFill(page, 1)).toBe(OFFICE_ACCENT1);
  // …and applying a second theme moves it again — it stored the reference.
  await page.getByTestId("slides-theme-toggle").click();
  await applyBuiltIn(page, "Ion");
  await expect.poll(() => shapeFill(page, 1)).toBe(ION_ACCENT1);
  await expect.poll(() => shapeFill(page, 0)).toBe(ION_ACCENT1); // the rect too

  // AC-59: the literal run is unmoved after re-open + re-theme, and the panel
  // names the applied theme.
  const literalRun = stageFrames(page).nth(4).locator("span").last();
  await expect(literalRun).toHaveCSS("color", LITERAL_RED);
  await expect(page.getByTestId("theme-current-name")).toContainText("Ion");
});
