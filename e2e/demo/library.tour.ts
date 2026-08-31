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
//
// Pacing is tuned for the recorded video (~90 s): every full-screen swap —
// an import opening its editor, a close landing back in the library — gets
// a dwell beat on both sides so no state flashes past the viewer, captions
// announce each swap before it happens, and the SMOOTHING_CSS injected
// after the intro fades the preferences dialog in, pre-paints the PDF
// embed's ground, and lets theme/palette switches glide instead of snap.
// ---------------------------------------------------------------------------

const fixturesDir = join(dirname(fileURLToPath(import.meta.url)), "..", "fixtures");

// Demo-only presentation polish, injected into the page (like the captions):
// a mount fade for the preferences dialog, and a colour transition on
// everything so palette and light/dark switches read as a deliberate glide
// on video rather than a single-frame flash. Colour transitions never run on
// freshly mounted DOM, so route swaps and list re-renders are unaffected.
// The pdf rule pre-paints the embed area in Chromium's own PDF-viewer grey
// (#525659): the plugin boots light-grey → dark-grey → page within ~400 ms,
// and matching the ground up front collapses that visible double-flash into
// one content appearance.
const SMOOTHING_CSS = `
  /* Plugin elements are left untouched — no reason to force paint-affecting
     properties onto the <object>/<embed> hosting Chromium's PDF viewer. */
  *:not(object):not(embed):not(iframe),
  *:not(object):not(embed):not(iframe)::before,
  *:not(object):not(embed):not(iframe)::after {
    transition: background-color 260ms ease, color 260ms ease, border-color 260ms ease !important;
  }
  [data-testid="preferences-dialog"] { animation: __demoFade 280ms ease-out both; }
  [data-testid="preferences-dialog"] .dialog { animation: __demoRise 280ms ease-out both; }
  @keyframes __demoFade { from { opacity: 0; } }
  @keyframes __demoRise { from { transform: translateY(10px) scale(0.985); } }
  /* The wrapper (not the plugin element itself) pre-painted in the PDF
     viewer's own grey, so its light-grey → dark-grey boot flash collapses
     into one content appearance. */
  [data-testid="media-pdf"] { background: #525659 !important; }
  /* The global rule above would cancel the overlay helpers' own opacity
     fades (an id selector out-specifies it): keep the title card and the
     caption pill fading, not snapping. */
  #__demo_title_card, #__demo_caption { transition: opacity 350ms ease !important; }
`;

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

  // The recording starts with the navigation, a beat before the title card's
  // evaluate can run — without a cover, the first frames are a white flash of
  // the empty library. This paints the title card's ground from the very
  // first paint; the tour removes it once the card itself is up. Top frame
  // ONLY: init scripts run in every frame, and a cover inside Chromium's PDF
  // viewer frame blanks the document for good (the main-frame removal can't
  // reach it) — that bug shipped once.
  await page.addInitScript(() => {
    if (window !== window.top) {
      return;
    }
    const cover = document.createElement("div");
    cover.id = "__demo_precover";
    cover.style.cssText = "position:fixed;inset:0;background:#111827;z-index:2147483646";
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

  /** Imports a file through the library's Open flow and waits for its
   * viewer — the same kind-aware routing a real drop goes through. */
  const importFile = async (path: string, readyTestId: string): Promise<void> => {
    await page.getByTestId("open-document-input").setInputFiles(path);
    await expect(page.getByTestId(readyTestId)).toBeVisible({ timeout: 20_000 });
  };

  /** Closes the current viewer and dwells on the library, so the return
   * never reads as a flash between two full-screen swaps. */
  const backToLibrary = async (closeTestId = "close-document"): Promise<void> => {
    await page.getByTestId(closeTestId).click();
    await expect(page.getByTestId("library")).toBeVisible();
    await beat(page, 1000);
  };

  /** A library card by its title (titles here are unique by construction). */
  const card = (title: string) => page.getByTestId("library-row").filter({ hasText: title });

  // -- Intro ----------------------------------------------------------------
  // The title card goes up before anything settles (the pre-cover holds its
  // ground from the first paint), so the viewer never sees a blink of the
  // empty library.
  await page.goto("/");
  // The card must be painted before the pre-cover comes down — the other way
  // round leaves a one-frame white gap between them on video.
  await titleCard(page, "docu-edit — the library", "Six formats. Ten design themes. All of it local.", 3400);
  await page.evaluate(() => document.getElementById("__demo_precover")?.remove());
  await expect(page.getByTestId("new-document").first()).toBeVisible();
  await page.addStyleTag({ content: SMOOTHING_CSS });

  // -- Import montage: one home for every format ----------------------------
  // Each caption goes up while the library is still on screen, then the file
  // opens under it — the viewer is told what is coming before the screen
  // changes, and no swap ever lands under a stale caption.
  await caption(page, "Bring your files — each one opens in its own editor");
  await beat(page, 1900);

  await caption(page, "Word .docx — real pages, real styles");
  await beat(page, 700);
  await importFile(seeded("word-toc.docx", "Quarterly Report.docx"), "document-content");
  await beat(page, 2100);
  await backToLibrary();

  await caption(page, "OpenDocument .odt — first-class, not an afterthought");
  await beat(page, 700);
  await importFile(seeded("lo-shapes.odt", "Design Notes.odt"), "document-content");
  await beat(page, 1900);
  await backToLibrary();

  await caption(page, "Markdown, straight onto a page");
  await beat(page, 700);
  await importFile(markdownSeed, "document-content");
  await beat(page, 1700);
  await backToLibrary();

  await caption(page, "Slide decks — a real .pptx editor is built in");
  await beat(page, 700);
  await importFile(seeded("mini-deck.pptx", "Pitch Deck.pptx"), "presentation-view");
  await beat(page, 2100);
  await backToLibrary("presentation-close");

  await caption(page, "PDFs live here too — viewed in place");
  await beat(page, 700);
  await importFile(pdfSeed, "media-workspace");
  await beat(page, 2000);
  await backToLibrary();

  await caption(page, "…and your images");
  await beat(page, 700);
  await importFile(posterSeed, "media-workspace");
  await beat(page, 1800);
  await backToLibrary();

  // -- The redesigned library -----------------------------------------------
  await expect(page.getByTestId("library-row")).toHaveCount(6);
  await page.getByTestId("view-grid").click();
  await caption(page, "The new library: sections on the left, live previews of every file");
  // Thumbnails render off the critical path (idle) — give the grid a moment
  // to fill in the page miniatures, the slide preview and the PDF raster.
  await beat(page, 4200);
  await shot(page, "library-grid-seeded");

  await caption(page, "Click to peek — Quick Look previews it before anything opens");
  await card("Pitch Deck").click();
  await expect(page.getByTestId("quick-look-title")).toHaveText("Pitch Deck");
  await beat(page, 2600);
  await card("Contract").click();
  await expect(page.getByTestId("quick-look-title")).toHaveText("Contract");
  await shot(page, "quick-look-pdf");
  await beat(page, 2400);

  await caption(page, "Star what matters — the rail keeps it one click away");
  await card("Contract").getByTestId("row-star").click();
  await beat(page, 700);
  await card("Quarterly Report").getByTestId("row-star").click();
  await beat(page, 900);
  await page.getByTestId("rail-starred").click();
  await expect(page.getByTestId("library-row")).toHaveCount(2);
  await shot(page, "starred-section");
  await beat(page, 2600);
  await page.getByTestId("rail-all").click();
  await expect(page.getByTestId("library-row")).toHaveCount(6);
  await beat(page, 800);

  await caption(page, "Search is instant — your files never leave your machine");
  await page.getByTestId("library-search").click();
  await page.keyboard.type("pitch", { delay: 150 });
  await expect(page.getByTestId("library-row")).toHaveCount(1);
  await beat(page, 2400);
  await page.getByTestId("library-search").fill("");
  await expect(page.getByTestId("library-row")).toHaveCount(6);
  await beat(page, 700);

  // -- The stopwatch open ---------------------------------------------------
  await caption(page, "No cloud between you and your work — watch the open");
  await beat(page, 1800);
  const openStarted = Date.now();
  await card("Quarterly Report").dblclick();
  await expect(page.getByTestId("document-content")).toBeVisible();
  const openMs = Date.now() - openStarted;
  await caption(page, `Opened in ${openMs} ms — straight from local storage, no spinner`);
  await shot(page, "stopwatch-open");
  await beat(page, 3600);
  await backToLibrary();

  // -- Design themes --------------------------------------------------------
  await caption(page, "Now make it yours — ten designer palettes, applied live");
  await page.getByTestId("library-preferences").click();
  await expect(page.getByTestId("preferences-dialog")).toBeVisible();
  await beat(page, 1400);
  await page.getByTestId("preferences-palette-nocturne").click();
  await beat(page, 2100);
  await page.getByTestId("preferences-palette-classic").click();
  await beat(page, 2100);
  await page.getByTestId("preferences-palette-pine").click();
  await beat(page, 2100);
  await page.getByTestId("preferences-palette-ochre").click();
  await shot(page, "palette-ochre-light");
  await beat(page, 2200);

  await caption(page, "Every palette is contrast-checked — in light and in dark");
  await page.getByTestId("preferences-theme-dark").click();
  await beat(page, 2800);
  await page.getByTestId("preferences-palette-petrol").click();
  await beat(page, 2100);
  await page.getByTestId("preferences-palette-nocturne").click();
  await beat(page, 2100);
  await page.getByTestId("preferences-close").click();
  await caption(page, "Your look, remembered on this device");
  await shot(page, "dark-nocturne-library");
  await beat(page, 3800);

  // -- Local-first ----------------------------------------------------------
  await caption(page, "It all lives right here — private browser storage, no account, ever");
  await page.getByTestId("storage-panel").hover();
  await beat(page, 3600);

  await caption(page, "Friendly to mistakes, too");
  await card("Launch Poster").click();
  await expect(page.getByTestId("quick-look-title")).toHaveText("Launch Poster");
  await beat(page, 900);
  await page.getByTestId("quick-look-trash").click();
  await expect(page.getByTestId("library-row")).toHaveCount(5);
  await expect(page.getByTestId("toast-action")).toBeVisible();
  await beat(page, 2200);
  await page.getByTestId("toast-action").click(); // Undo
  await expect(page.getByTestId("library-row")).toHaveCount(6);
  await caption(page, "Undo brings it straight back");
  await shot(page, "undo-restore");
  await beat(page, 2600);

  // -- Finale ---------------------------------------------------------------
  await titleCard(page, "docu-edit", "Well-rounded · fast · entirely yours — .docx · .odt · .md · .pptx · .pdf · images", 4000);

  console.log(`library tour duration: ${((Date.now() - started) / 1000).toFixed(1)}s`);
});
