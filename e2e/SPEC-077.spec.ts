import { test, expect, type Page } from "@playwright/test";
import { LINE_END, LINE_START, MOD } from "./keys";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import { BUILT_IN_THEMES, resolveColor } from "../src/lib/core/theme";
import { FIXTURE_THEME, TABLE_SHADING_REF } from "../src/lib/fixtures.theme";

// SPEC-077 — document themes, the AC-54 journey. The unit suites pin the
// model, both round trips and every picker in isolation; this file drives
// the seams only a browser confirms: Format → Theme… with the current card
// marked, one apply recolouring every theme-referencing surface on the page
// (heading run, table shading, shape-canvas fill) while an explicit colour
// stands still, a theme swatch storing a reference that SURVIVES a save →
// reopen — proven by a second theme moving that run again.

const fixturesDir = join(dirname(fileURLToPath(import.meta.url)), "fixtures");

const content = (page: Page) => page.getByTestId("document-content");

const facet = BUILT_IN_THEMES.find((theme) => theme.name === "Facet")!;
const office = BUILT_IN_THEMES.find((theme) => theme.name === "Office")!;

/** "#rrggbb" as the "rgb(r, g, b)" string computed styles report. */
function rgb(hex: string): string {
  const n = Number.parseInt(hex.slice(1), 16);
  return `rgb(${(n >> 16) & 0xff}, ${(n >> 8) & 0xff}, ${n & 0xff})`;
}

async function openFixture(page: Page, name: string): Promise<void> {
  await page.goto("/");
  await page.getByTestId("open-document-input").setInputFiles(join(fixturesDir, name));
  await expect(page).toHaveURL(/\/doc\/.+/);
  await expect(content(page).locator("p, h1, h2").first()).toBeVisible();
  await waitForPagination(page);
}

async function waitForPagination(page: Page, timeout = 30_000): Promise<void> {
  await expect(page.getByTestId("document-pages")).toHaveAttribute("data-pagination-state", "clean", {
    timeout,
  });
}

async function waitForSaved(page: Page): Promise<void> {
  await expect(page.getByTestId("save-state")).toHaveText("All changes saved", { timeout: 15_000 });
}

/** The rendered element carrying the given fixture text. */
function textOf(page: Page, text: string) {
  return content(page).getByText(text, { exact: true }).first();
}

async function colorOf(page: Page, text: string): Promise<string> {
  return textOf(page, text).evaluate((el) => globalThis.getComputedStyle(el).color);
}

/** The shape canvas's centre pixel as [r, g, b] — the fixture shape carries
 * no rotation or shadow, so the canvas centre is solid fill. */
async function shapeCentrePixel(page: Page): Promise<number[]> {
  return content(page)
    .locator("canvas.doc-shape-canvas")
    .first()
    .evaluate((el) => {
      const canvas = el as HTMLCanvasElement;
      const ctx = canvas.getContext("2d")!;
      const data = ctx.getImageData(Math.floor(canvas.width / 2), Math.floor(canvas.height / 2), 1, 1).data;
      return [data[0], data[1], data[2]];
    });
}

/** A pixel of the shape canvas by canvas coordinate — `x` counted from the
 * left edge, so x = 1 lands inside the outline stroke rather than on its
 * antialiased boundary. */
async function shapePixelAt(page: Page, x: number, fromMiddle = true): Promise<number[]> {
  return content(page)
    .locator("canvas.doc-shape-canvas")
    .first()
    .evaluate(
      (el, [px, mid]) => {
        const canvas = el as HTMLCanvasElement;
        const ctx = canvas.getContext("2d")!;
        const y = mid ? Math.floor(canvas.height / 2) : 1;
        const data = ctx.getImageData(px, y, 1, 1).data;
        return [data[0], data[1], data[2]];
      },
      [x, fromMiddle ? 1 : 0] as [number, number],
    );
}

/** Whether any pixel of the shape canvas is (within tolerance) `hex` — the
 * dependable way to look for a dashed outline's colour, whose stroke covers
 * only part of any given row. */
async function shapeCanvasHas(page: Page, hex: string): Promise<boolean> {
  return content(page)
    .locator("canvas.doc-shape-canvas")
    .first()
    .evaluate((el, target) => {
      const canvas = el as HTMLCanvasElement;
      const ctx = canvas.getContext("2d")!;
      const n = Number.parseInt(target.slice(1), 16);
      const want = [(n >> 16) & 0xff, (n >> 8) & 0xff, n & 0xff];
      const { data } = ctx.getImageData(0, 0, canvas.width, canvas.height);
      for (let i = 0; i < data.length; i += 4) {
        if (
          data[i + 3] === 255 &&
          Math.abs(data[i] - want[0]) <= 3 &&
          Math.abs(data[i + 1] - want[1]) <= 3 &&
          Math.abs(data[i + 2] - want[2]) <= 3
        ) {
          return true;
        }
      }
      return false;
    }, hex);
}

/** The outline's colour, found anywhere on the shape canvas.
 *
 * The four preset checks used to read it from `shapePixelAt(page, 0)` — the
 * single leftmost pixel of the row through the middle. That is where the
 * stroke lands only while the shape comes out at the exact size this file's
 * zoom-then-insert sequence usually produces (226x152 at 200%, where x0 is
 * the border, x1 is already an antialiased blend and x2 on is the fill). Under
 * worker oversubscription the 200% zoom can still be settling when the
 * insertion click lands, the shape is a pixel out, and the stroke straddles
 * x0/x1: the sample reads the blend (60,103,180) instead of the border
 * (46,83,149) and misses the ±3 tolerance by exactly 6, which is the failure
 * seen at `--workers=16`.
 *
 * Scanning is what this file already calls "the dependable way" to find a
 * stroke colour, and it is not vacuous here: in every one of these presets the
 * outline is the ONLY thing painted in the colour being looked for — the fill
 * is a different colour, or there is no fill. Wrapped in toPass so a canvas
 * that has not repainted yet is a retry rather than a failure. */
async function expectOutlineColour(page: Page, hex: string): Promise<void> {
  await expect(async () => {
    expect(await shapeCanvasHas(page, hex), `no pixel of the shape is ${hex}`).toBe(true);
  }).toPass({ timeout: 10_000 });
}

function expectPixel(pixel: number[], hex: string): void {
  const n = Number.parseInt(hex.slice(1), 16);
  const expected = [(n >> 16) & 0xff, (n >> 8) & 0xff, n & 0xff];
  for (let channel = 0; channel < 3; channel++) {
    expect(Math.abs(pixel[channel] - expected[channel]), `channel ${channel} of ${hex}`).toBeLessThanOrEqual(3);
  }
}

/** Format → Theme… → click the named built-in's card. */
async function applyTheme(page: Page, name: string): Promise<void> {
  await page.getByTestId("menu-title-format").click();
  await page.getByTestId("format-menu-theme").click();
  await expect(page.getByTestId("theme-gallery-dialog")).toBeVisible();
  await page.locator(`[data-testid="theme-card"][data-theme-name="${name}"]`).click();
  await expect(page.getByTestId("theme-gallery-dialog")).toBeHidden();
}

/** A blank document with one rectangle drawn on it, left selected, zoomed to
 * 200% so a 1pt outline covers a whole canvas pixel and can be sampled
 * without antialiasing against the fill. */
async function newDocumentWithShape(page: Page): Promise<void> {
  await page.goto("/");
  await page.getByTestId("new-document").click();
  await expect(page).toHaveURL(/\/doc\/.+/);
  await content(page).click({ position: { x: 20, y: 20 } });
  await page.keyboard.type("Preset check");
  await page.getByTestId("zoom-percent").click();
  await page.getByTestId("zoom-preset-200").click();
  await waitForPagination(page);
  await page.getByTestId("menu-title-insert").click();
  await page.getByTestId("shape-gallery-rect").click();
  await content(page).click({ position: { x: 400, y: 250 } });
  await expect(page.getByTestId("shape-toolbar")).toBeVisible();
}

/** Applies a style preset from the toolbar's Style popover. */
async function applyPreset(page: Page, preset: string): Promise<void> {
  if (!(await page.getByTestId("shape-style-panel").isVisible())) {
    await page.getByTestId("shape-style-button").click();
    await expect(page.getByTestId("shape-style-panel")).toBeVisible();
  }
  await page.getByTestId(`shape-style-preset-${preset}`).click();
}

/** Selects the first `length` characters of the paragraph holding `text`. */
async function selectStartOf(page: Page, text: string, length: number): Promise<void> {
  await textOf(page, text).click();
  await page.keyboard.press(LINE_START);
  await page.keyboard.down("Shift");
  for (let i = 0; i < length; i++) {
    await page.keyboard.press("ArrowRight");
  }
  await page.keyboard.up("Shift");
}

/** File → Export → the chosen format, returning the download's file path. */
async function download(page: Page, format: "docx" | "odt"): Promise<string> {
  await page.getByTestId("menu-title-file").click();
  await page.getByTestId("file-menu-export").click();
  const file = await Promise.all([
    page.waitForEvent("download"),
    page.getByTestId(`download-${format}`).click(),
  ]).then(([d]) => d);
  const path = await file.path();
  expect(path).toBeTruthy();
  return path!;
}

async function reopen(page: Page, path: string): Promise<void> {
  await page.goto("/");
  await page.getByTestId("open-document-input").setInputFiles(path);
  await expect(page).toHaveURL(/\/doc\/.+/);
  await expect(content(page).locator("p, h1, h2").first()).toBeVisible();
  await waitForPagination(page);
}

const HEADING = "Quarterly review";
const EXPLICIT = "This run is explicitly dark red and stays put.";
const BODY = "Body copy between the themed pieces, unformatted.";

test("the AC-54 journey: gallery, one-apply recolour, an accent swatch that survives save → reopen", async ({ page }) => {
  test.setTimeout(120_000);
  await openFixture(page, "word-theme.docx");

  // The fixture opens resolving against ITS OWN theme, not a built-in.
  await expect(textOf(page, HEADING)).toHaveCSS("color", rgb(FIXTURE_THEME.colors.accent2));
  await expect(textOf(page, EXPLICIT)).toHaveCSS("color", rgb("#c00000"));
  const cell = content(page).locator("td").filter({ hasText: "Region" }).first();
  await expect(cell).toHaveCSS("background-color", rgb(resolveColor(TABLE_SHADING_REF, FIXTURE_THEME)));
  expectPixel(await shapeCentrePixel(page), FIXTURE_THEME.colors.accent4);

  // AC-7: selecting the theme-filled shape shows the RESOLVED hex in the
  // quick toolbar's fill input — not a stringified object rendered black.
  await content(page).locator('[data-testid^="drawing-"]').first().click();
  await expect(page.getByTestId("shape-toolbar")).toBeVisible();
  await expect(page.getByTestId("shape-fill-color")).toHaveValue(FIXTURE_THEME.colors.accent4);
  await textOf(page, BODY).click();

  // Format → Theme…: the document's CURRENT theme leads the gallery, marked.
  await page.getByTestId("menu-title-format").click();
  await page.getByTestId("format-menu-theme").click();
  await expect(page.getByTestId("theme-gallery-dialog")).toBeVisible();
  const current = page.getByTestId("theme-card-current");
  await expect(current).toHaveAttribute("data-theme-name", FIXTURE_THEME.name);
  await expect(current.getByTestId("theme-current-badge")).toBeVisible();

  // Apply Facet: every theme-referencing surface moves at once — heading
  // run, table shading, shape fill — while the explicit run stands still.
  await page.locator('[data-testid="theme-card"][data-theme-name="Facet"]').click();
  await expect(page.getByTestId("theme-gallery-dialog")).toBeHidden();
  await expect(textOf(page, HEADING)).toHaveCSS("color", rgb(facet.colors.accent2));
  await expect(cell).toHaveCSS("background-color", rgb(resolveColor(TABLE_SHADING_REF, facet)));
  await expect(textOf(page, EXPLICIT)).toHaveCSS("color", rgb("#c00000"));
  expectPixel(await shapeCentrePixel(page), facet.colors.accent4);

  // Pick the base Accent 1 swatch for a run of body text: the run shows
  // Facet's accent1 because it stores the REFERENCE, not the hex.
  await selectStartOf(page, BODY, 4);
  await page.getByTestId("text-color-button").click();
  await expect(page.getByTestId("text-color-menu")).toBeVisible();
  await page.getByTestId("text-color-menu").getByLabel("Accent 1", { exact: true }).click();
  await expect(textOf(page, "Body")).toHaveCSS("color", rgb(facet.colors.accent1));
  await waitForSaved(page);

  // Save → reopen: the accent run still TRACKS the accent…
  const saved = await download(page, "docx");
  await reopen(page, saved);
  await expect(textOf(page, "Body")).toHaveCSS("color", rgb(facet.colors.accent1));

  // …proven by a second theme moving it again, and the literal run never moves.
  await applyTheme(page, "Office");
  await expect(textOf(page, "Body")).toHaveCSS("color", rgb(office.colors.accent1));
  await expect(textOf(page, HEADING)).toHaveCSS("color", rgb(office.colors.accent2));
  await expect(textOf(page, EXPLICIT)).toHaveCSS("color", rgb("#c00000"));
  expectPixel(await shapeCentrePixel(page), office.colors.accent4);

  // AC-14: the apply was ONE undo entry — a single undo restores the
  // previous theme and every colour that resolved from it, and redo
  // re-applies it.
  await content(page).click();
  await page.keyboard.press(`${MOD}+z`);
  await expect(textOf(page, "Body")).toHaveCSS("color", rgb(facet.colors.accent1));
  await expect(textOf(page, HEADING)).toHaveCSS("color", rgb(facet.colors.accent2));
  expectPixel(await shapeCentrePixel(page), facet.colors.accent4);
  await page.keyboard.press(`${MOD}+y`);
  await expect(textOf(page, "Body")).toHaveCSS("color", rgb(office.colors.accent1));
  await expect(textOf(page, HEADING)).toHaveCSS("color", rgb(office.colors.accent2));
});

test("a theme page background never freezes the live word count (AC-22)", async ({ page }) => {
  test.setTimeout(60_000);
  await openFixture(page, "word-theme.docx");

  // Store a ThemeColorRef as the section's page background through Page
  // setup → Borders — the object-valued field the word-count projection must
  // keep away from postMessage's structured clone.
  await page.getByTestId("menu-title-file").click();
  await page.getByTestId("page-setup-button").click();
  await expect(page.getByTestId("page-setup-dialog")).toBeVisible();
  await page.getByTestId("page-setup-tab-borders").click();
  await page.getByTestId("page-setup-dialog").getByLabel("Accent 2", { exact: true }).click();
  await page.getByTestId("page-setup-apply").click();
  await expect(page.getByTestId("page-setup-dialog")).toBeHidden();

  // Typing must still advance the live count (it froze with DataCloneError
  // when the reference reached the worker as a reactive Proxy).
  const count = page.getByTestId("word-count-button");
  const before = (await count.textContent()) ?? "";
  await textOf(page, BODY).click();
  await page.keyboard.press(LINE_END);
  await page.keyboard.type(" zeta eta theta");
  await expect(count).not.toHaveText(before, { timeout: 10_000 });
});

// SPEC-073's D3 table, pinned as literals rather than read back out of
// SHAPE_STYLE_PRESETS: AC-44 is the proof that turning those colours into
// accent1 references changed nothing a document already shipped with.
const FILLED_FILL = "#4472c4";
const FILLED_BORDER = "#2e5395";
const SUBTLE_FILL = "#dae3f3";

test("every style preset resolves to SPEC-073's shipped colours under Office, and only the accent-drawn ones follow a theme (AC-43, AC-44)", async ({
  page,
}) => {
  test.setTimeout(120_000);
  await newDocumentWithShape(page);

  // A document with no theme of its own resolves against DEFAULT_THEME —
  // Office — so each preset must paint exactly what SPEC-073 shipped.
  await applyPreset(page, "filled");
  expectPixel(await shapeCentrePixel(page), FILLED_FILL);
  await expectOutlineColour(page, FILLED_BORDER);

  await applyPreset(page, "subtleFilled");
  expectPixel(await shapeCentrePixel(page), SUBTLE_FILL);
  await expectOutlineColour(page, FILLED_FILL);

  await applyPreset(page, "outlined");
  await expectOutlineColour(page, FILLED_FILL);

  // AC-43: the fill is a reference, so a theme change restyles the shape —
  // while the filled preset's border, the one colour that was never the
  // accent nor a lighter variant of it, keeps its literal across the change.
  await applyPreset(page, "filled");
  await page.keyboard.press("Escape");
  await applyTheme(page, "Facet");
  expectPixel(await shapeCentrePixel(page), facet.colors.accent1);
  await expectOutlineColour(page, FILLED_BORDER);
});

test("an outline width or dash edit leaves a theme-referenced outline a reference (AC-18)", async ({ page }) => {
  test.setTimeout(120_000);
  await newDocumentWithShape(page);

  // Give the outline an Accent 3 reference through the style panel's theme
  // row, and make it thick enough to sample away from the fill's edge.
  await page.getByTestId("shape-style-button").click();
  await expect(page.getByTestId("shape-style-panel")).toBeVisible();
  await page.getByTestId("shape-outline-theme-row").getByLabel("Accent 3", { exact: true }).click();
  await page.getByTestId("shape-style-panel").getByTestId("shape-style-outline-width").fill("4");
  await expect(page.getByTestId("shape-outline-color")).toHaveValue(office.colors.accent3);
  await page.getByTestId("shape-style-button").click();

  // Editing only the width and the dash must re-emit the stored reference,
  // never the hex the colour input displays.
  await page.getByTestId("shape-outline-width").fill("5");
  await page.getByTestId("shape-outline-dash").selectOption("dash");
  await expect(page.getByTestId("shape-outline-color")).toHaveValue(office.colors.accent3);

  // The proof the model still holds a reference: a theme change moves the
  // outline. A flattened literal would stay Office's accent3 for ever.
  expect(await shapeCanvasHas(page, office.colors.accent3)).toBe(true);
  await page.keyboard.press("Escape");
  await applyTheme(page, "Facet");
  await expect
    .poll(() => shapeCanvasHas(page, facet.colors.accent3), { timeout: 10_000 })
    .toBe(true);
  expect(await shapeCanvasHas(page, office.colors.accent3)).toBe(false);
});
