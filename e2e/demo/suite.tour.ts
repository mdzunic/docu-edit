import { test, expect, type Locator, type Page } from "@playwright/test";
import { copyFileSync, mkdirSync } from "node:fs";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import { LINE_END, LINE_START } from "../keys";
import { beat, caption, hideTitleCard, shot, showTitleCard, titleCard } from "./tour-helpers";
import { writeDemoPdf } from "./seed-media";

// ---------------------------------------------------------------------------
// docu-edit — the one-minute suite tour: Docs, Sheets, Slides, Library.
//
// A short, headed companion to demo.tour.ts (the editor in depth) and
// library.tour.ts (the home in depth): one quick pass through every product,
// one signature feature each — a table from the Insert grid in Docs, a live
// =SUM pointed at with the arrow keys in Sheets, a slide added from the
// layout gallery in Slides — ending in the library's live previews, Quick
// Look and the built-in media viewer. Runs only through playwright.demo.config.ts:
//
//   npx playwright test --config=playwright.demo.config.ts suite.tour.ts
//
// (DEMO_PORT=<port> runs it on its own dev server when 5173 is taken.)
// Assertions are sync points that keep the pacing honest, nothing more —
// every interaction mirrors a SPEC-NNN e2e (028 tables, 122 cell entry,
// 089 layouts, 107 media). Paced for ~45 s of video.
// ---------------------------------------------------------------------------

const here = dirname(fileURLToPath(import.meta.url));
const TYPE_DELAY = 40; // ms per keystroke — a watchable typewriter pace

async function type(page: Page, text: string): Promise<void> {
  await page.keyboard.type(text, { delay: TYPE_DELAY });
}

/** Renames the open file through the app bar's title control. */
async function rename(page: Page, title: string): Promise<void> {
  await page.getByTestId("document-title").click();
  const input = page.getByTestId("document-title-input");
  await input.fill("");
  await input.pressSequentially(title, { delay: 35 });
  await page.keyboard.press("Enter");
  await expect(page.getByTestId("document-title")).toContainText(title);
}

function sheetCell(page: Page, ref: string): Locator {
  return page.locator(`[data-testid="sheet-cell"][data-ref="${ref}"]`);
}

/** Clicks a cell's centre — empty cells have no element of their own, so
 * the point comes from its column and row headers (as SPEC-122 does). */
async function clickCell(page: Page, ref: string): Promise<void> {
  const [, letters, digits] = /^([A-Z]+)(\d+)$/.exec(ref)!;
  const column = await page.getByTestId("sheet-col-header").filter({ hasText: new RegExp(`^${letters}$`) }).boundingBox();
  const row = await page.getByTestId("sheet-row-header").filter({ hasText: new RegExp(`^${digits}$`) }).boundingBox();
  await page.mouse.click(column!.x + column!.width / 2, row!.y + row!.height / 2, { delay: 40 });
  await expect(page.getByTestId("name-box")).toHaveValue(ref);
}

/** Types one row of the sheet: each value then Tab, the last one Enter. */
async function typeRow(page: Page, values: string[]): Promise<void> {
  for (const [i, value] of values.entries()) {
    await type(page, value);
    await page.keyboard.press(i === values.length - 1 ? "Enter" : "Tab");
  }
}

test("docu-edit — one quick pass through Docs, Sheets and Slides", async ({ page }) => {
  const started = Date.now();

  // A PDF and a photo for the library finale, seeded under the intro card so
  // the viewer never watches an import. Titles come from the file names.
  const seedDir = test.info().outputPath("seed");
  mkdirSync(seedDir, { recursive: true });
  const pdfSeed = join(seedDir, "Contract.pdf");
  writeDemoPdf(pdfSeed);
  const photoSeed = join(seedDir, "Golden Gate.jpg");
  copyFileSync(join(here, "golden-gate.jpg"), photoSeed);

  // Squiggle-free from the first keystroke, and an opaque ground from the
  // first paint (top frame only — init scripts also run inside Chromium's
  // PDF viewer frame, where a cover blanks the document for good).
  await page.addInitScript(() => {
    if (!localStorage.getItem("docu-edit:preferences")) {
      localStorage.setItem("docu-edit:preferences", JSON.stringify({ spellCheck: false }));
    }
    if (window !== window.top) {
      return;
    }
    const cover = document.createElement("div");
    cover.id = "__demo_precover";
    cover.style.cssText = "position:fixed;inset:0;background:#111827;z-index:2147483646;pointer-events:none";
    const attach = () => document.documentElement.appendChild(cover);
    if (document.documentElement) {
      attach();
    } else {
      new MutationObserver((_, observer) => {
        if (document.documentElement) {
          attach();
          observer.disconnect();
        }
      }).observe(document, { childList: true });
    }
  });

  // -- Intro (seeding happens beneath the card) -----------------------------
  await page.goto("/");
  await showTitleCard(page, "docu-edit", "Docs · Sheets · Slides — right in your browser");
  const introShown = Date.now();
  await expect(page.getByTestId("new-document").first()).toBeVisible();
  // Chromium's PDF viewer boots light-grey → dark-grey → page; pre-painting
  // its wrapper in the viewer's own grey collapses that into one appearance.
  await page.addStyleTag({ content: '[data-testid="media-pdf"] { background: #525659 !important; }' });
  for (const seed of [photoSeed, pdfSeed]) {
    await page.getByTestId("open-document-input").setInputFiles(seed);
    await expect(page.getByTestId("media-workspace")).toBeVisible({ timeout: 20_000 });
    await page.getByTestId("close-document").click();
    await expect(page.getByTestId("library")).toBeVisible();
  }
  await page.waitForTimeout(Math.max(0, 3000 - (Date.now() - introShown)));
  await page.evaluate(() => document.getElementById("__demo_precover")?.remove());
  await hideTitleCard(page);

  // -- Docs: a table in two clicks ------------------------------------------
  await caption(page, "Docs — a real word processor, on real pages");
  await beat(page, 900);
  await page.getByTestId("new-document").first().click();
  const content = page.getByTestId("document-content");
  await expect(content).toBeVisible();
  await rename(page, "Launch Plan");
  await content.click();
  await type(page, "Launch plan");
  await page.getByTestId("style-picker-button").click();
  await page.getByTestId("style-option-Title").click();
  await page.keyboard.press(LINE_END);
  await page.keyboard.press("Enter");
  await type(page, "Owners and dates are in the deck.");
  // The table inserts above the caret's line (the AC-14 pattern), so the
  // sentence ends up below it.
  await page.keyboard.press(LINE_START);
  await caption(page, "Insert a table straight from the grid");
  await page.getByTestId("menu-title-insert").click();
  await page.getByTestId("table-grid-cell-1-1").hover();
  await beat(page, 250);
  await page.getByTestId("table-grid-cell-2-2").hover();
  await beat(page, 250);
  await page.getByTestId("table-grid-cell-3-3").hover();
  await beat(page, 450);
  await page.getByTestId("table-grid-cell-3-3").click();
  await expect(content.locator("table.doc-table")).toHaveCount(1);
  await page.getByTestId("table-cell-0-0").click();
  await type(page, "Team");
  await page.keyboard.press("Tab");
  await type(page, "Owner");
  await page.keyboard.press("Tab");
  await type(page, "Ships");
  await page.keyboard.press("Tab");
  await type(page, "Design");
  await page.keyboard.press("Tab");
  await type(page, "Ana");
  await page.keyboard.press("Tab");
  await type(page, "May 4");
  await page.mouse.move(90, 500, { steps: 3 }); // no hover handles in the frame
  await shot(page, "docs-table");
  await beat(page, 1800);

  // -- Sheets: a live formula -----------------------------------------------
  await caption(page, "Sheets — Excel-style cells and live formulas");
  await page.getByTestId("close-document").click();
  await expect(page.getByTestId("library")).toBeVisible();
  await beat(page, 600);
  await page.getByTestId("new-spreadsheet").first().click();
  await expect(page).toHaveURL(/\/sheets\/.+/);
  await clickCell(page, "A1");
  await typeRow(page, ["Team", "Cost"]);
  await typeRow(page, ["Design", "1200"]);
  await typeRow(page, ["Build", "3400"]);
  await typeRow(page, ["Launch", "900"]);
  await type(page, "Total");
  await page.keyboard.press("Tab");
  await caption(page, "Point at a range with the arrow keys");
  await type(page, "=SUM(");
  for (let i = 0; i < 3; i++) {
    await page.keyboard.press("ArrowUp");
    await beat(page, 220);
  }
  await page.keyboard.press("Shift+ArrowDown");
  await beat(page, 220);
  await page.keyboard.press("Shift+ArrowDown");
  await expect(page.getByTestId("cell-editor")).toHaveText("=SUM(B2:B4");
  await beat(page, 500);
  await type(page, ")");
  await page.keyboard.press("Enter");
  await expect(sheetCell(page, "B5")).toHaveText("5500");
  await beat(page, 900);
  await caption(page, "Change a number — the total follows");
  await clickCell(page, "B2");
  await beat(page, 300);
  await type(page, "2000");
  await page.keyboard.press("Enter");
  await expect(sheetCell(page, "B5")).toHaveText("6300");
  await clickCell(page, "B5"); // the formula bar shows the formula
  await shot(page, "sheets-formula");
  await beat(page, 2000);

  // -- Slides: a slide from a layout ----------------------------------------
  await caption(page, "Slides — start every slide from a layout");
  await page.getByTestId("close-document").click();
  await expect(page.getByTestId("library")).toBeVisible();
  await beat(page, 600);
  await page.getByTestId("new-presentation").first().click();
  await expect(page.getByTestId("presentation-view")).toBeVisible();
  const stageFrames = page.getByTestId("stage-workspace").getByTestId("slide-frame");
  await stageFrames.first().dblclick();
  await type(page, "Spring launch");
  await page.keyboard.press("Escape");
  await beat(page, 500);
  await page.getByTestId("presentation-slide-menu").first().click();
  await page.getByTestId("slide-menu-new-with-layout").click();
  const gallery = page.getByTestId("slides-layout-gallery");
  await expect(gallery).toBeVisible();
  await beat(page, 1300);
  await gallery.getByRole("option", { name: "Section Header" }).hover();
  await beat(page, 350);
  await gallery.getByRole("option", { name: "Two Content" }).click();
  await shot(page, "layout-gallery");
  await beat(page, 700);
  await page.getByTestId("layout-apply").click();
  await expect(gallery).toBeHidden();
  await expect(page.getByTestId("filmstrip-slide")).toHaveCount(2);
  await stageFrames.first().dblclick();
  await type(page, "Design vs. Build");
  await page.keyboard.press("Escape");
  await page.mouse.move(90, 500, { steps: 3 });
  await shot(page, "slides-layout");
  await beat(page, 1900);

  // -- Library and viewer ---------------------------------------------------
  await caption(page, "One private library — live previews of every file");
  await page.getByTestId("presentation-close").click();
  await expect(page.getByTestId("library")).toBeVisible();
  await expect(page.getByTestId("library-row")).toHaveCount(5);
  await page.getByTestId("view-grid").click();
  await beat(page, 3000); // thumbnails render at idle
  await shot(page, "library-grid");
  const card = (title: string) => page.getByTestId("library-row").filter({ hasText: title });
  await caption(page, "Click to peek — Quick Look previews before anything opens");
  await card("Contract").click();
  await expect(page.getByTestId("quick-look-title")).toHaveText("Contract");
  await shot(page, "quick-look-pdf");
  await beat(page, 2400);
  await caption(page, "…and photos open in the built-in viewer");
  await card("Golden Gate").dblclick();
  await expect(page.getByTestId("media-image")).toBeVisible();
  await beat(page, 2600);
  await shot(page, "photo-viewer");

  // -- Finale ---------------------------------------------------------------
  await titleCard(page, "docu-edit", "Docs · Sheets · Slides · PDFs · photos — local, private, yours", 3400);

  console.log(`suite tour duration: ${((Date.now() - started) / 1000).toFixed(1)}s`);
});
