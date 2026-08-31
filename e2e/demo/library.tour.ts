import { test, expect } from "@playwright/test";
import { copyFileSync, mkdirSync, writeFileSync } from "node:fs";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import { beat, caption, shot, titleCard } from "./tour-helpers";
import { writeDemoPdf, writeDemoPoster } from "./seed-media";

// ---------------------------------------------------------------------------
// docu-edit — round two: the library, the themes, the local-first pitch.
//
// A headed, watchable companion to demo.tour.ts. Where the first tour sells
// the editor, this one sells the home around it: six file formats imported
// into one library, the redesigned full-tier layout (rail, live previews,
// Quick Look), starring and instant search, a measured stopwatch open
// straight from local storage, the ten-palette design-theme system with its
// light and dark schemes, and the private on-device storage story — down to
// the friendly trash-with-undo. Runs only through playwright.demo.config.ts
// (the regular suite never matches *.tour.ts):
//
//   npx playwright test --config=playwright.demo.config.ts library.tour.ts
//
// (Without the file argument the config runs every tour back to back.)
// Assertions are deliberately light — the regression suite owns correctness;
// here they are sync points that keep the pacing honest.
// ---------------------------------------------------------------------------

const fixturesDir = join(dirname(fileURLToPath(import.meta.url)), "..", "fixtures");

test("docu-edit — library, themes and the local-first story", async ({ page }) => {
  const started = Date.now();

  // The library entries take their titles from the file names, so the seed
  // files are fixture copies under presentable names — plus a generated
  // poster PNG and text PDF (seed-media.ts), because the SPEC-107 media
  // fixtures are parser-minimal (a 1×1 pixel, an empty-page PDF) and would
  // put blank previews on the demo's cards.
  const seedDir = test.info().outputPath("seed");
  mkdirSync(seedDir, { recursive: true });
  const seeded = (fixture: string, niceName: string): string => {
    const dest = join(seedDir, niceName);
    copyFileSync(join(fixturesDir, fixture), dest);
    return dest;
  };
  const markdownSeed = join(seedDir, "Meeting Notes.md");
  writeFileSync(
    markdownSeed,
    [
      "# Meeting notes — launch week",
      "",
      "- Library redesign shipped",
      "- Ten colour themes, each with a dark scheme",
      "- Everything stays on-device",
      "",
      "## Next",
      "",
      "1. Folders and tags",
      "2. Shared templates",
      "",
    ].join("\n"),
  );
  const pdfSeed = join(seedDir, "Contract.pdf");
  writeDemoPdf(pdfSeed);
  const posterSeed = join(seedDir, "Launch Poster.png");
  writeDemoPoster(posterSeed);

  // Keep the editor beats squiggle-free, as in the first tour: the montage
  // opens fixture documents whose text trips the spell checker. Seeding the
  // preference before the app boots (the loader tolerates a partial object)
  // spares the audience an unexplained settings visit mid-intro.
  await page.addInitScript(() => {
    if (!localStorage.getItem("docu-edit:preferences")) {
      localStorage.setItem("docu-edit:preferences", JSON.stringify({ spellCheck: false }));
    }
  });

  /** Imports a file through the library's Open flow and waits for its
   * viewer — the same kind-aware routing a real drop goes through. */
  const importFile = async (path: string, readyTestId: string): Promise<void> => {
    await page.getByTestId("open-document-input").setInputFiles(path);
    await expect(page.getByTestId(readyTestId)).toBeVisible({ timeout: 20_000 });
  };

  /** A library card by its title (titles here are unique by construction). */
  const card = (title: string) => page.getByTestId("library-row").filter({ hasText: title });

  // -- Intro ----------------------------------------------------------------
  await page.goto("/");
  await expect(page.getByTestId("new-document").first()).toBeVisible();
  await titleCard(page, "docu-edit — the library", "Six formats. Ten design themes. All of it local.", 3400);

  // -- Import montage: one home for every format ----------------------------
  await caption(page, "Bring your files — each one opens in its own editor");
  await beat(page, 1800);

  await importFile(seeded("word-toc.docx", "Quarterly Report.docx"), "document-content");
  await caption(page, "Word .docx — real pages, real styles");
  await beat(page, 1900);
  await page.getByTestId("close-document").click();
  await expect(page.getByTestId("library")).toBeVisible();

  await importFile(seeded("lo-shapes.odt", "Design Notes.odt"), "document-content");
  await caption(page, "OpenDocument .odt — first-class, not an afterthought");
  await beat(page, 1700);
  await page.getByTestId("close-document").click();
  await expect(page.getByTestId("library")).toBeVisible();

  await importFile(markdownSeed, "document-content");
  await caption(page, "Markdown, straight onto a page");
  await beat(page, 1500);
  await page.getByTestId("close-document").click();
  await expect(page.getByTestId("library")).toBeVisible();

  await importFile(seeded("mini-deck.pptx", "Pitch Deck.pptx"), "presentation-view");
  await caption(page, "Slide decks — a real .pptx editor is built in");
  await beat(page, 1900);
  await page.getByTestId("presentation-close").click();
  await expect(page.getByTestId("library")).toBeVisible();

  await importFile(pdfSeed, "media-workspace");
  await caption(page, "PDFs live here too — viewed in place");
  await beat(page, 1800);
  await page.getByTestId("close-document").click();
  await expect(page.getByTestId("library")).toBeVisible();

  await importFile(posterSeed, "media-workspace");
  await caption(page, "…and your images");
  await beat(page, 1600);
  await page.getByTestId("close-document").click();
  await expect(page.getByTestId("library")).toBeVisible();

  // -- The redesigned library -----------------------------------------------
  await expect(page.getByTestId("library-row")).toHaveCount(6);
  await page.getByTestId("view-grid").click();
  await caption(page, "The new library: sections on the left, live previews of every file");
  // Thumbnails render off the critical path (idle) — give the grid a moment
  // to fill in the page miniatures, the slide preview and the PDF raster.
  await beat(page, 3800);
  await shot(page, "library-grid-seeded");

  await caption(page, "Click to peek — Quick Look previews it before anything opens");
  await card("Pitch Deck").click();
  await expect(page.getByTestId("quick-look-title")).toHaveText("Pitch Deck");
  await beat(page, 2400);
  await card("Contract").click();
  await expect(page.getByTestId("quick-look-title")).toHaveText("Contract");
  await shot(page, "quick-look-pdf");
  await beat(page, 2200);

  await caption(page, "Star what matters — the rail keeps it one click away");
  await card("Contract").getByTestId("row-star").click();
  await beat(page, 600);
  await card("Quarterly Report").getByTestId("row-star").click();
  await beat(page, 800);
  await page.getByTestId("rail-starred").click();
  await expect(page.getByTestId("library-row")).toHaveCount(2);
  await shot(page, "starred-section");
  await beat(page, 2400);
  await page.getByTestId("rail-all").click();
  await expect(page.getByTestId("library-row")).toHaveCount(6);
  await beat(page, 700);

  await caption(page, "Search is instant — your files never leave your machine");
  await page.getByTestId("library-search").click();
  await page.keyboard.type("pitch", { delay: 110 });
  await expect(page.getByTestId("library-row")).toHaveCount(1);
  await beat(page, 2200);
  await page.getByTestId("library-search").fill("");
  await expect(page.getByTestId("library-row")).toHaveCount(6);
  await beat(page, 500);

  // -- The stopwatch open ---------------------------------------------------
  await caption(page, "No cloud between you and your work — watch the open");
  await beat(page, 1600);
  const openStarted = Date.now();
  await card("Quarterly Report").dblclick();
  await expect(page.getByTestId("document-content")).toBeVisible();
  const openMs = Date.now() - openStarted;
  await caption(page, `Opened in ${openMs} ms — straight from local storage, no spinner`);
  await shot(page, "stopwatch-open");
  await beat(page, 3400);
  await page.getByTestId("close-document").click();
  await expect(page.getByTestId("library")).toBeVisible();

  // -- Design themes --------------------------------------------------------
  await caption(page, "Now make it yours — ten designer palettes, applied live");
  await page.getByTestId("library-preferences").click();
  await expect(page.getByTestId("preferences-dialog")).toBeVisible();
  await beat(page, 1000);
  await page.getByTestId("preferences-palette-nocturne").click();
  await beat(page, 1900);
  await page.getByTestId("preferences-palette-classic").click();
  await beat(page, 1900);
  await page.getByTestId("preferences-palette-pine").click();
  await beat(page, 1900);
  await page.getByTestId("preferences-palette-ochre").click();
  await shot(page, "palette-ochre-light");
  await beat(page, 2000);

  await caption(page, "Every palette is contrast-checked — in light and in dark");
  await page.getByTestId("preferences-theme-dark").click();
  await beat(page, 2400);
  await page.getByTestId("preferences-palette-petrol").click();
  await beat(page, 1900);
  await page.getByTestId("preferences-palette-nocturne").click();
  await beat(page, 1900);
  await page.getByTestId("preferences-close").click();
  await caption(page, "Your look, remembered on this device");
  await shot(page, "dark-nocturne-library");
  await beat(page, 3400);

  // -- Local-first ----------------------------------------------------------
  await caption(page, "It all lives right here — private browser storage, no account, ever");
  await page.getByTestId("storage-panel").hover();
  await beat(page, 3200);

  await caption(page, "Friendly to mistakes, too");
  await card("Launch Poster").click();
  await expect(page.getByTestId("quick-look-title")).toHaveText("Launch Poster");
  await beat(page, 800);
  await page.getByTestId("quick-look-trash").click();
  await expect(page.getByTestId("library-row")).toHaveCount(5);
  await expect(page.getByTestId("toast-action")).toBeVisible();
  await beat(page, 2000);
  await page.getByTestId("toast-action").click(); // Undo
  await expect(page.getByTestId("library-row")).toHaveCount(6);
  await caption(page, "Undo brings it straight back");
  await shot(page, "undo-restore");
  await beat(page, 2400);

  // -- Finale ---------------------------------------------------------------
  await titleCard(page, "docu-edit", "Well-rounded · fast · entirely yours — .docx · .odt · .md · .pptx · .pdf · images", 4000);

  console.log(`library tour duration: ${((Date.now() - started) / 1000).toFixed(1)}s`);
});
