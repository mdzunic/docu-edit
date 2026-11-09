import { test, expect, type Download, type Page } from "@playwright/test";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import { MOD, LINE_END, LINE_START, extend } from "../keys";
import { beat, caption, shot, titleCard } from "./tour-helpers";

// ---------------------------------------------------------------------------
// docu-edit — the two-minute product tour.
//
// A headed, watchable pitch for the product: create a document, type on real
// pages, format text, styles and lists, tables, charts with editable data,
// shapes, the command palette, footnotes, inline comments, track changes
// with one-click accept, headers and self-updating page numbers, a second
// page with a rotated callout, zoom, dark mode, the phone-sized compact
// tier, export to .docx / .odt / .md (with the honest conversion report),
// and the presentation product. Spell-check is switched off up front: a
// pitch needs a squiggle-free page. Runs only through
// playwright.demo.config.ts (the regular suite never matches *.tour.ts):
//
//   npx playwright test --config=playwright.demo.config.ts
//
// Every interaction mirrors a pattern from a real SPEC-NNN e2e spec, so the
// tour doubles as a smoke pass over the demoed features — but its assertions
// are deliberately light; the regression suite owns correctness. Pacing
// lands around 110–120 s on a warm dev server; adjust the beat() values /
// TYPE_DELAY to taste, keeping the whole run at two minutes or less.
// ---------------------------------------------------------------------------

const fixturesDir = join(dirname(fileURLToPath(import.meta.url)), "..", "fixtures");
const TYPE_DELAY = 42; // ms per keystroke — a watchable typewriter pace

async function type(page: Page, text: string): Promise<void> {
  await page.keyboard.type(text, { delay: TYPE_DELAY });
}

/** Extends the selection right by `chars` caret stops. */
async function selectRight(page: Page, chars: number): Promise<void> {
  await page.keyboard.down("Shift");
  for (let i = 0; i < chars; i++) {
    await page.keyboard.press("ArrowRight");
  }
  await page.keyboard.up("Shift");
}

/** Same barrier the SPEC-028+ specs use after anything that reflows pages. */
async function waitForPagination(page: Page, timeout = 30_000): Promise<void> {
  await expect(page.getByTestId("document-pages")).toHaveAttribute(
    "data-pagination-state",
    "clean",
    { timeout },
  );
}

test("docu-edit — two-minute product tour", async ({ page }) => {
  const started = Date.now();
  const content = page.getByTestId("document-content");

  // -- Intro ----------------------------------------------------------------
  await page.goto("/");
  await expect(page.getByTestId("new-document").first()).toBeVisible();
  await titleCard(page, "docu-edit", "Word-class editing in a browser tab — nothing to install", 3400);

  // -- Create and name a document -------------------------------------------
  await caption(page, "No install. No account. Files stay on your machine.");
  await beat(page, 1500);
  await page.getByTestId("new-document").first().click();
  await expect(page).toHaveURL(/\/doc\/.+/);
  await expect(content).toBeVisible();
  await page.getByTestId("document-title").click();
  await page.getByTestId("document-title-input").pressSequentially("Product Tour", { delay: 60 });
  await page.keyboard.press("Enter");
  await beat(page, 800);

  // Keep the tour squiggle-free: red underlines under the demo's product
  // names would read as sloppiness, so turn "Check spelling as you type"
  // off before any text exists.
  await page.getByTestId("menu-title-file").click();
  await page.getByTestId("file-menu-preferences").click();
  await expect(page.getByTestId("preferences-dialog")).toBeVisible();
  await page.getByTestId("preferences-spell-check").uncheck();
  await page.getByTestId("preferences-close").click();
  await beat(page, 400);

  // -- Type on a real page --------------------------------------------------
  await caption(page, "Type on a real page — layout happens live");
  await content.click();
  await type(page, "docu-edit");
  await page.keyboard.press("Enter");
  await type(page, "Every document format, right in your browser.");
  await content.getByText("docu-edit").click();
  await page.getByTestId("style-picker-button").click();
  await expect(page.getByTestId("style-picker-menu")).toBeVisible();
  await beat(page, 400);
  await page.getByTestId("style-option-Title").click();
  await shot(page, "title-style");
  await beat(page, 1600);

  // -- Character formatting -------------------------------------------------
  await caption(page, "Rich formatting, exactly where you expect it");
  await content.getByText("right in your browser").click();
  await page.keyboard.press(LINE_START);
  await page.keyboard.press(extend(LINE_END)); // select the whole sentence
  await beat(page, 800); // the floating selection toolbar appears
  await page.keyboard.press(`${MOD}+b`);
  await beat(page, 600);
  await page.getByTestId("format-italic").click();
  await beat(page, 600);
  await page.getByTestId("text-color-button").click();
  await beat(page, 400);
  await page.locator('[data-testid="text-color-swatch"][data-color="#ff0000"]').click();
  await page.keyboard.press("ArrowRight"); // collapse the selection
  await shot(page, "formatted-sentence");
  await beat(page, 1200);
  await caption(page, "…with unlimited undo");
  await page.keyboard.press(`${MOD}+z`);
  await beat(page, 500);
  await page.keyboard.press(`${MOD}+z`);
  await beat(page, 700);
  await page.keyboard.press(`${MOD}+Shift+z`);
  await beat(page, 500);
  await page.keyboard.press(`${MOD}+Shift+z`);
  await beat(page, 800);

  // -- Styles and lists -----------------------------------------------------
  await caption(page, "Real paragraph styles, real lists");
  await page.keyboard.press(LINE_END);
  await page.keyboard.press("Enter");
  await type(page, "Highlights");
  await page.getByTestId("style-picker-button").click();
  await beat(page, 400);
  await page.getByTestId("style-option-Heading1").click();
  await page.keyboard.press(LINE_END);
  await page.keyboard.press("Enter");
  await page.getByTestId("list-bulleted").click();
  await type(page, "Opens .docx, .odt and Markdown");
  await page.keyboard.press("Enter");
  await type(page, "Tables, charts and shapes");
  await page.keyboard.press("Enter");
  await page.keyboard.press("Enter"); // empty item exits the list
  await type(page, "The numbers:");
  await page.keyboard.press("Enter");
  await type(page, "Revenue by quarter:");
  // The table inserts above this line, and the chart anchors to it — keeping
  // both at the document's end, where the chart's floating wrap rectangle
  // cannot sit over text later beats click.
  await page.keyboard.press(LINE_START);
  await shot(page, "lists");
  await beat(page, 900);

  // -- Tables ---------------------------------------------------------------
  await caption(page, "A table in two clicks");
  await page.getByTestId("menu-title-insert").click();
  await page.getByTestId("table-grid-cell-1-1").hover();
  await beat(page, 400);
  await page.getByTestId("table-grid-cell-2-2").hover();
  await beat(page, 400);
  await page.getByTestId("table-grid-cell-3-4").hover();
  await beat(page, 600);
  await page.getByTestId("table-grid-cell-3-4").click();
  await expect(content.locator("table.doc-table")).toHaveCount(1);
  await page.getByTestId("table-cell-0-0").click();
  await type(page, "Region");
  await page.keyboard.press("Tab");
  await type(page, "Q1");
  await page.keyboard.press("Tab");
  await type(page, "Q2");
  await shot(page, "table");
  await beat(page, 1600);

  // -- Charts with editable data --------------------------------------------
  await caption(page, "Charts are built in — double-click to edit the data");
  await content.getByText("Revenue by quarter:").click();
  await page.keyboard.press(LINE_END);
  await page.getByTestId("menu-title-insert").click();
  await page.getByTestId("chart-gallery-columnClustered").click();
  const chart = content.locator('[data-testid^="drawing-"]').first();
  await expect(chart).toBeVisible();
  await beat(page, 1100);
  await chart.dblclick();
  await expect(page.getByTestId("chart-data-sheet-dialog")).toBeVisible();
  await beat(page, 700);
  // Sample data is [4.3, 2.5, 3.5] — 9 makes the first column visibly jump
  // without flattening the rest of the chart.
  await page.getByTestId("chart-data-cell-1-1").fill("9");
  await page.getByTestId("chart-data-cell-1-1").blur();
  await beat(page, 600);
  await page.getByTestId("chart-data-apply").click();
  await shot(page, "chart-redrawn");
  await beat(page, 2000); // watch the chart redraw

  // -- Shapes ---------------------------------------------------------------
  await caption(page, "Shapes anchor to your text and follow it");
  await content.click({ position: { x: 20, y: 20 } }); // anchor paragraph
  await page.getByTestId("menu-title-insert").click();
  await page.getByTestId("shape-gallery-star5").click();
  await content.click({ position: { x: 855, y: 270 } }); // right margin, beside the bullets
  await expect(page.getByTestId("shape-toolbar")).toBeVisible();
  await shot(page, "star-shape");
  await beat(page, 2000);
  await page.keyboard.press("Escape"); // deselect the shape

  // -- Command palette + footnote -------------------------------------------
  await caption(page, "Every command is one keystroke away");
  await content.getByText("right in your browser").click();
  await page.keyboard.press(LINE_END); // the note reference belongs after "browser.", not mid-word
  await page.keyboard.press(`${MOD}+k`);
  await expect(page.getByTestId("command-palette")).toBeVisible();
  await page.getByTestId("command-palette-input").pressSequentially("footnote", { delay: 65 });
  await beat(page, 800);
  await page.keyboard.press("Enter"); // runs "Insert footnote"; caret lands in the note
  await expect(page.getByTestId("footnote-area")).toBeVisible();
  await type(page, "Made with docu-edit — no install required.");
  await shot(page, "footnote");
  await beat(page, 1200);

  // -- Inline comments ------------------------------------------------------
  await caption(page, "Built for review — inline comments");
  await page.getByTestId("menu-title-view").click();
  await page.getByTestId("view-menu-comments-pane").click();
  const commentsPane = page.getByTestId("comments-pane");
  await expect(commentsPane).toBeVisible();
  await commentsPane.getByTestId("comment-author-name-input").fill("Sam");
  await content.getByText("right in your browser").click();
  await page.keyboard.press(LINE_START);
  // "browser" starts 37 caret stops into the sentence.
  for (let i = 0; i < 37; i++) {
    await page.keyboard.press("ArrowRight");
  }
  await selectRight(page, "browser".length);
  await page.keyboard.press(`${MOD}+Alt+m`);
  await expect(commentsPane.getByTestId("comment-composer")).toBeVisible();
  await commentsPane
    .getByTestId("comment-composer-input")
    .pressSequentially("Great line — ship it.", { delay: 30 });
  await shot(page, "comment");
  await beat(page, 1500); // the filled composer over the highlighted range is the shot
  await commentsPane.getByTestId("comment-composer-submit").click();
  await expect(page.locator('[data-testid^="comment-highlight-"]')).toHaveCount(1);
  await expect(commentsPane.locator('[data-testid^="comment-thread-"]')).toHaveCount(1);
  await beat(page, 800);
  await page.getByTestId("menu-title-view").click();
  await page.getByTestId("view-menu-comments-pane").click(); // close the pane
  await beat(page, 400);

  // -- Track changes --------------------------------------------------------
  await caption(page, "Track changes — every edit becomes a redline");
  await page.getByTestId("status-bar-track-changes").click();
  await expect(page.getByTestId("status-bar-track-changes")).toContainText("On");
  await beat(page, 700);
  await content.getByText("Opens .docx").click();
  await page.keyboard.press(LINE_START);
  // Insert first, then strike the old word — typing over a live selection
  // with tracking on interleaves the per-keystroke insertions into the
  // deletion mark and accept-all then commits the scrambled text.
  await type(page, "Round-trips "); // tracked insertion
  await selectRight(page, "Opens ".length);
  await page.keyboard.press("Backspace"); // tracked deletion — struck out, not removed
  await shot(page, "redline");
  await beat(page, 2000);
  await caption(page, "…one click accepts them all");
  await page.getByTestId("menu-title-view").click();
  await page.getByTestId("view-menu-track-changes-pane").click();
  const trackPane = page.getByTestId("track-changes-pane");
  await expect(trackPane).toBeVisible();
  await beat(page, 900);
  await trackPane.getByTestId("track-changes-accept-all").click();
  await beat(page, 1300);
  await page.getByTestId("menu-title-view").click();
  await page.getByTestId("view-menu-track-changes-pane").click(); // close the pane
  await page.getByTestId("status-bar-track-changes").click(); // tracking back off
  await beat(page, 500);

  // -- Headers and page numbers ---------------------------------------------
  await caption(page, "Headers, footers, self-updating page numbers");
  await page.getByTestId("menu-title-insert").click();
  await page.getByTestId("hf-menu-header").click(); // body dims, caret in the header band
  await type(page, "Product Tour · docu-edit");
  await beat(page, 900);
  await page.keyboard.press("Escape"); // back to body editing
  await page.getByTestId("menu-title-insert").click();
  await page.getByTestId("hf-menu-page-number").click();
  await expect(page.getByTestId("page-number-dialog")).toBeVisible();
  await page.getByTestId("page-number-part").selectOption("footer");
  await page.getByTestId("page-number-position").selectOption("right");
  await page.getByTestId("page-number-of-total").check(); // "Page X of Y"
  await beat(page, 700);
  await page.getByTestId("page-number-apply").click();
  await waitForPagination(page);
  await shot(page, "header-pagenum");
  await beat(page, 800);

  // -- Second page + a rotated callout --------------------------------------
  await caption(page, "More pages? The page numbers keep up.");
  await content.getByText("Revenue by quarter:").click();
  // Not LINE_END here: this paragraph wraps beside the chart, and LINE_END
  // stops at the visual line's end — breaking there splits the sentence and
  // drags the anchored chart to page 2. The editor's own Cmd/Ctrl+End goes
  // to the true document end.
  await page.keyboard.press(`${MOD}+End`);
  await page.getByTestId("menu-title-insert").click();
  await page.getByTestId("break-page").click();
  await waitForPagination(page);
  await type(page, "Try it today — it runs wherever your browser runs.");
  await beat(page, 900);
  await caption(page, "Callouts from the shape gallery — place, rotate, done");
  await page.getByTestId("menu-title-insert").click();
  await page.getByTestId("shape-gallery-wedgeRectCallout").click(); // arms draw mode
  const pageTwo = page.getByTestId("document-page").nth(1);
  const pageTwoBox = (await pageTwo.boundingBox())!;
  await page.mouse.click(pageTwoBox.x + pageTwoBox.width * 0.55, pageTwoBox.y + 280);
  await expect(page.getByTestId("shape-toolbar")).toBeVisible();
  await beat(page, 900);
  await page.getByTestId("shape-rotation").fill("20");
  await page.getByTestId("shape-rotation").press("Enter");
  await shot(page, "callout-rotated");
  await beat(page, 1600);
  await page.keyboard.press("Escape"); // deselect the callout

  // -- Zoom and dark mode ---------------------------------------------------
  await caption(page, "Step back — both pages, header to footer");
  // Rulers off for the wide shot — the framed two-page spread reads cleaner
  // without measurement chrome.
  await page.getByTestId("menu-title-view").click();
  await page.getByTestId("toggle-ruler-button").click();
  await page.getByTestId("zoom-percent").click();
  await expect(page.getByTestId("zoom-menu")).toBeVisible();
  await beat(page, 400);
  await page.getByTestId("zoom-fit-page").click();
  // Depending on where earlier beats left the workspace scrolled, the fitted
  // page can sit partly below the fold — bring all of it (header, footer,
  // page number) on screen for the wide shot, and park the pointer off the
  // pages so no hover overlays (table handles) sit in the frame.
  await page.getByTestId("document-page").first().scrollIntoViewIfNeeded();
  // Park the pointer off the pages so no hover affordance sits in the frame.
  await page.mouse.move(90, 500, { steps: 3 });
  await beat(page, 2200);
  await caption(page, "Easy on the eyes, too — full dark mode");
  await page.getByTestId("menu-title-view").click();
  await page.getByTestId("view-menu-appearance").click();
  await expect(page.getByTestId("view-menu-appearance-flyout")).toBeVisible();
  await beat(page, 400);
  await page.getByTestId("appearance-dark").click();
  await page.keyboard.press("Escape");
  await page.keyboard.press("Escape");
  // The closing menu leaves the pointer over the page (the browser re-hit-
  // tests under a stationary pointer), which can legitimately arm the table
  // hover handles — park it off the pages before the wide shot.
  await page.mouse.move(90, 500, { steps: 3 });
  await shot(page, "dark-fit-page");
  await beat(page, 2800);

  // -- Compact tier: phone-sized viewport -----------------------------------
  await caption(page, "Fits your phone, too");
  await page.getByTestId("zoom-percent").click();
  await page.getByTestId("zoom-fit-width").click(); // natural reading zoom for a narrow screen
  await page.setViewportSize({ width: 390, height: 844 });
  await beat(page, 900); // the compact layout tier kicks in
  await shot(page, "mobile");
  await beat(page, 2400);
  await page.setViewportSize({ width: 1440, height: 900 });
  await beat(page, 1400); // let the toolbar re-measure back to the desktop tier

  // -- Export: .docx, .odt, .md ---------------------------------------------
  const exportAs = async (format: "docx" | "odt" | "md"): Promise<Download> => {
    // The .docx and .odt exports can raise the conversion report too (the
    // chart and the rotated callout are the usual items). An open report
    // covers the File menu, so close it before the next export.
    const report = page.getByTestId("conversion-report");
    if (await report.isVisible()) {
      await page.getByTestId("conversion-report-close").click();
      await expect(report).toBeHidden();
      await beat(page, 300);
    }
    await page.getByTestId("menu-title-file").click();
    await beat(page, 300);
    await page.getByTestId("file-menu-export").click();
    await beat(page, 400);
    const [download] = await Promise.all([
      page.waitForEvent("download"),
      page.getByTestId(`download-${format}`).click(),
    ]);
    return download;
  };

  await caption(page, "Exports a real Word .docx — fidelity first");
  const docx = await exportAs("docx");
  expect(docx.suggestedFilename()).toMatch(/\.docx$/);
  await caption(page, `Saved ${docx.suggestedFilename()} ✓ — opens anywhere Word does`);
  await beat(page, 1800);
  const odt = await exportAs("odt");
  expect(odt.suggestedFilename()).toMatch(/\.odt$/);
  await caption(page, "Native OpenDocument .odt too — at home in LibreOffice");
  await beat(page, 1800);
  const md = await exportAs("md");
  expect(md.suggestedFilename()).toMatch(/\.md$/);
  await caption(page, "Even Markdown — with an honest report of what won't carry");
  await expect(page.getByTestId("conversion-report")).toBeVisible();
  await shot(page, "conversion-report");
  await beat(page, 2400); // let the report's item list be read
  await page.getByTestId("conversion-report-close").click();
  await beat(page, 400);

  // -- Presentations --------------------------------------------------------
  await caption(page, "Presentations? Built in — this is a real .pptx");
  await page.getByTestId("close-document").click();
  await expect(page.getByTestId("library-row")).toHaveCount(1);
  await page.getByTestId("view-grid").click(); // library grid with thumbnails
  await shot(page, "library-grid");
  await beat(page, 1500);
  await page.getByTestId("open-document-input").setInputFiles(join(fixturesDir, "mini-deck.pptx"));
  await expect(page).toHaveURL(/\/slides\/.+/);
  await expect(page.getByTestId("presentation-view")).toBeVisible();
  await expect(page.getByTestId("presentation-slide-count")).toHaveText("4 slides");
  await shot(page, "slides-open");
  await beat(page, 1800);
  const thumbs = page.getByTestId("filmstrip-slide");
  await thumbs.nth(1).click();
  await beat(page, 1000);
  await thumbs.nth(2).click();
  await beat(page, 1000);
  await thumbs.nth(3).click();
  await beat(page, 1000);
  await caption(page, "Slides add, reorder, duplicate — with undo");
  await page.getByTestId("presentation-add-slide").click();
  await expect(page.getByTestId("presentation-slide-count")).toHaveText("5 slides");
  await beat(page, 1000);
  await page.getByTestId("presentation-undo").click();
  await expect(page.getByTestId("presentation-slide-count")).toHaveText("4 slides");
  await beat(page, 800);

  // -- Finale: the library --------------------------------------------------
  await caption(page, "One private library — local, fast, yours");
  await page.getByTestId("presentation-close").click();
  await expect(page.getByTestId("library-row")).toHaveCount(2);
  await shot(page, "library-final");
  await beat(page, 2600);
  await titleCard(page, "That's docu-edit", "Opens, edits and preserves .docx · .odt · .md · .pptx", 3600);

  console.log(`tour duration: ${((Date.now() - started) / 1000).toFixed(1)}s`);
});
