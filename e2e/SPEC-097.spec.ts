import { test, expect, type Locator, type Page } from "@playwright/test";
import { readFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { MOD } from "./keys";
import { unzip } from "../src/lib/core/zip";
import { EXPECTED_TEXT_DEPTH } from "../src/lib/presentation/fixtures/build";

// SPEC-097 — slide text depth. Open a deck whose first slide carries a
// numbered body, a start-at-5 list, a sentence to link, an overflowing title
// and one deliberately out-of-subset vertical-text frame; demote a numbered
// line with Tab, set exact line spacing and space-after in the Text panel,
// shrink the overflowing title, link a word to slide 3, follow it in
// playback without also advancing slide 3's click build, then download →
// reopen and confirm markers, spacing, the autofit scale and the link all
// come back while the vertical-text frame's source survives verbatim.
const here = dirname(fileURLToPath(import.meta.url));
const PPTX = join(here, "fixtures", "spec097-text-depth.pptx");

async function openFixture(page: Page, file: string): Promise<string> {
  await page.goto("/");
  await page.getByTestId("open-document-input").setInputFiles(file);
  await expect(page).toHaveURL(/\/slides\/.+/);
  await expect(page.getByTestId("presentation-view")).toBeVisible();
  return new URL(page.url()).pathname.split("/")[2];
}

async function downloadTo(page: Page, name: string): Promise<string> {
  const downloadPromise = page.waitForEvent("download");
  await page.getByTestId("presentation-download").click();
  const download = await downloadPromise;
  const filePath = join(tmpdir(), name);
  await download.saveAs(filePath);
  return filePath;
}

/** A top-level slide frame on the editor stage whose text contains `text`. */
function frameWith(page: Page, text: string): Locator {
  return page.getByTestId("stage-workspace").getByTestId("slide-frame").filter({ hasText: text });
}

/** The rendered list markers of a frame, in order. */
async function markersOf(frame: Locator): Promise<string[]> {
  return (await frame.locator(".bullet-marker").allTextContents()).map((t) => t.trim());
}

/** Whether a frame's rendered text fits inside its own box (1px tolerance). */
async function textFits(frame: Locator): Promise<boolean> {
  return frame.evaluate((el) => {
    const box = el.getBoundingClientRect();
    const paragraphs = Array.from(el.querySelectorAll("[data-testid='slide-paragraph']"));
    const bottom = Math.max(...paragraphs.map((p) => p.getBoundingClientRect().bottom));
    return bottom <= box.bottom + 1;
  });
}

/** Selects `word` inside the open contenteditable host (the session's own
 * selectionchange listener records it for the next command). */
async function selectWordInSession(page: Page, word: string): Promise<void> {
  await page.getByTestId("slide-text-edit").evaluate((host, target) => {
    (host as HTMLElement).focus();
    const walker = document.createTreeWalker(host, NodeFilter.SHOW_TEXT);
    for (let node = walker.nextNode(); node; node = walker.nextNode()) {
      const at = (node.textContent ?? "").indexOf(target);
      if (at >= 0) {
        const range = document.createRange();
        range.setStart(node, at);
        range.setEnd(node, at + target.length);
        const selection = document.getSelection()!;
        selection.removeAllRanges();
        selection.addRange(range);
        return;
      }
    }
    throw new Error(`"${target}" not found in the editing host`);
  }, word);
}

test("pptx: numbered lists, Tab demote, spacing, shrink autofit and slide links round-trip", async ({ page }) => {
  const deckId = await openFixture(page, PPTX);
  const body = frameWith(page, EXPECTED_TEXT_DEPTH.bodyLines[0]);
  const continued = frameWith(page, EXPECTED_TEXT_DEPTH.continuedLines[0]);

  // AC-60: the numbered body renders 1.–3.; the second list honours its
  // start-at of 5.
  await expect(body.locator(".bullet-marker")).toHaveCount(3);
  expect(await markersOf(body)).toEqual(EXPECTED_TEXT_DEPTH.bodyMarkers);
  expect(await markersOf(continued)).toEqual(EXPECTED_TEXT_DEPTH.continuedMarkers);

  // AC-61: in a session on line 3, Tab demotes it and its marker becomes "a.".
  await body.getByText(EXPECTED_TEXT_DEPTH.bodyLines[2]).dblclick();
  await expect(page.getByTestId("slide-text-edit")).toBeVisible();
  await page.keyboard.press("Tab");
  await expect(page.getByTestId("slide-text-edit")).toBeFocused();
  await page.keyboard.press("Escape");
  await expect(page.getByTestId("slide-text-edit")).toHaveCount(0);
  await expect.poll(() => markersOf(body)).toEqual(["1.", "2.", "a."]);

  // AC-62: Text panel → line spacing "exactly 24pt", space-after 6pt on the
  // selected body.
  await body.click();
  await page.getByTestId("slides-text-toggle").click();
  const panel = page.getByTestId("slides-text-panel");
  await expect(panel).toBeVisible();
  await expect(panel).toHaveAttribute("aria-label", "Text for slide 1");
  await panel.getByTestId("text-line-spacing-mode").selectOption("exact");
  await panel.getByTestId("text-line-spacing").fill("24");
  await panel.getByTestId("text-line-spacing").press("Enter");
  await panel.getByTestId("text-space-after").fill("6");
  await panel.getByTestId("text-space-after").press("Enter");
  const firstLine = body.getByTestId("slide-paragraph").first();
  await expect(firstLine).toHaveCSS("line-height", "32px"); // 24pt
  await expect(firstLine).toHaveCSS("margin-bottom", "8px"); // 6pt

  // AC-63: the title overflows its frame; "Shrink text on overflow" makes it fit.
  const title = frameWith(page, EXPECTED_TEXT_DEPTH.title);
  expect(await textFits(title)).toBe(false);
  await title.click();
  await panel.getByTestId("text-autofit").selectOption({ label: "Shrink text on overflow" });
  await expect.poll(() => textFits(title)).toBe(true);

  // AC-64: select the word "results", Cmd/Ctrl+K, link it to slide 3.
  const sentence = frameWith(page, EXPECTED_TEXT_DEPTH.linkSentence);
  await sentence.getByText(EXPECTED_TEXT_DEPTH.linkSentence).dblclick();
  await expect(page.getByTestId("slide-text-edit")).toBeVisible();
  await selectWordInSession(page, EXPECTED_TEXT_DEPTH.linkWord);
  await page.keyboard.press(`${MOD}+k`);
  const dialog = page.getByTestId("slide-link-dialog");
  await expect(dialog).toBeVisible();
  await dialog.getByTestId("link-slide").selectOption(String(EXPECTED_TEXT_DEPTH.linkSlide));
  await dialog.getByTestId("link-apply").click();
  await expect(dialog).toHaveCount(0);
  await page.keyboard.press("Escape");
  await expect(sentence.getByTestId("slide-link")).toHaveText(EXPECTED_TEXT_DEPTH.linkWord);

  // AC-65: Present → clicking the link jumps to slide 3 without also
  // advancing that slide's click build.
  await page.getByTestId("slides-present").click();
  await expect(page).toHaveURL(new RegExp(`/slides/${deckId}/present/1$`));
  const play = page.getByTestId("presentation-play");
  await play.getByTestId("slide-link").filter({ hasText: EXPECTED_TEXT_DEPTH.linkWord }).click();
  await expect(page).toHaveURL(new RegExp(`/slides/${deckId}/present/3$`));
  const playStage = play.getByTestId("slide-stage");
  await expect(playStage.getByText(EXPECTED_TEXT_DEPTH.buildSlideTitle)).toBeVisible();
  await expect(playStage.getByText(EXPECTED_TEXT_DEPTH.buildShapeText)).toBeHidden();
  await expect(page).toHaveURL(new RegExp(`/slides/${deckId}/present/3$`));

  // AC-66: exit, download, reopen → markers, spacing, autofit scale and the
  // link all come back; the vertical-text frame is preserved verbatim.
  await page.keyboard.press("Escape");
  await expect(page.getByTestId("presentation-view")).toBeVisible();
  await page.getByTestId("filmstrip-slide").first().click();
  const file = await downloadTo(page, "spec097-text-depth-roundtrip.pptx");

  const parts = await unzip(new Uint8Array(readFileSync(file)));
  const slideXml = new TextDecoder().decode(parts.get("ppt/slides/slide1.xml")!);
  expect(slideXml).toContain('vert="vert270"');
  expect(slideXml).toContain('<a:tab pos="914400" algn="l"/>');
  expect(slideXml).toContain(EXPECTED_TEXT_DEPTH.verticalText);
  expect(slideXml).toMatch(/<a:normAutofit fontScale="\d+"/);
  expect(slideXml).toContain('action="ppaction://hlinksldjump"');

  await openFixture(page, file);
  const body2 = frameWith(page, EXPECTED_TEXT_DEPTH.bodyLines[0]);
  await expect(body2.locator(".bullet-marker")).toHaveCount(3);
  expect(await markersOf(body2)).toEqual(["1.", "2.", "a."]);
  expect(await markersOf(frameWith(page, EXPECTED_TEXT_DEPTH.continuedLines[0]))).toEqual(EXPECTED_TEXT_DEPTH.continuedMarkers);
  const firstLine2 = body2.getByTestId("slide-paragraph").first();
  await expect(firstLine2).toHaveCSS("line-height", "32px");
  await expect(firstLine2).toHaveCSS("margin-bottom", "8px");
  const title2 = frameWith(page, EXPECTED_TEXT_DEPTH.title);
  await expect.poll(() => textFits(title2)).toBe(true);
  const scale = await title2.locator(".text-body").evaluate((el) => getComputedStyle(el).getPropertyValue("--fit-scale").trim());
  expect(Number(scale)).toBeLessThan(1);
  await expect(frameWith(page, EXPECTED_TEXT_DEPTH.linkSentence).getByTestId("slide-link")).toHaveText(EXPECTED_TEXT_DEPTH.linkWord);
  await expect(frameWith(page, EXPECTED_TEXT_DEPTH.verticalText)).toBeVisible();
});

test("Enter in a Text panel field during a text session commits the field without splitting the paragraph", async ({ page }) => {
  await openFixture(page, PPTX);
  const body = frameWith(page, EXPECTED_TEXT_DEPTH.bodyLines[0]);
  await page.getByTestId("slides-text-toggle").click();
  const panel = page.getByTestId("slides-text-panel");
  await expect(panel).toBeVisible();

  // A session on line 3 with the caret inside the line, not at its end.
  await body.getByText(EXPECTED_TEXT_DEPTH.bodyLines[2]).dblclick();
  await expect(page.getByTestId("slide-text-edit")).toBeVisible();
  await page.keyboard.press("ArrowLeft");

  const spaceAfter = panel.getByTestId("text-space-after");
  await spaceAfter.click();
  await spaceAfter.fill("12");
  await spaceAfter.press("Enter");

  const lastLine = body.getByTestId("slide-paragraph").nth(2);
  await expect(lastLine).toHaveCSS("margin-bottom", "16px"); // 12pt
  await page.keyboard.press("Escape");
  await expect(page.getByTestId("slide-text-edit")).toHaveCount(0);
  await expect(body.getByTestId("slide-paragraph")).toHaveCount(3);
  await expect(lastLine).toHaveText(new RegExp(EXPECTED_TEXT_DEPTH.bodyLines[2]));
  expect(await markersOf(body)).toEqual(EXPECTED_TEXT_DEPTH.bodyMarkers);
});
