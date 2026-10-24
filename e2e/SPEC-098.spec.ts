import { test, expect, type Page } from "@playwright/test";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

const here = dirname(fileURLToPath(import.meta.url));
const FIXTURE_PPTX = join(here, "fixtures", "mini-deck.pptx");

// SPEC-098 — deck find and replace: open the pptx fixture, Cmd/Ctrl+F, search
// "Open" (slide 1's notes and slide 2's body), activate the slide-2 hit (slide
// 2 shows with that frame selected and the word selected), see "Grouped" come
// back empty with the scope note (slide 3's grouped text is deliberately not
// searched), then Replace all "Open" → "Launch" as one undo entry: undo
// restores both, redo, download and reopen — both strings changed.

async function openFixture(page: Page, file: string): Promise<void> {
  await page.goto("/");
  await page.getByTestId("open-document-input").setInputFiles(file);
  await expect(page).toHaveURL(/\/slides\/.+/);
  await expect(page.getByTestId("presentation-view")).toBeVisible();
  await expect(page.getByTestId("slide-frame").first()).toBeVisible();
}

/** The main stage (the filmstrip thumbnails are slide-stages too). */
function stage(page: Page) {
  return page.getByTestId("stage-workspace").getByTestId("slide-stage");
}

async function downloadTo(page: Page, name: string): Promise<string> {
  const downloadPromise = page.waitForEvent("download");
  await page.getByTestId("presentation-download").click();
  const download = await downloadPromise;
  const filePath = join(tmpdir(), name);
  await download.saveAs(filePath);
  return filePath;
}

test("find across the deck, scope note, and Replace all as one undo entry that round-trips", async ({ page }) => {
  await openFixture(page, FIXTURE_PPTX);

  // Cmd/Ctrl+F opens the shared find bar with the deck's scope.
  await page.keyboard.press("ControlOrMeta+f");
  await expect(page.getByTestId("find-bar")).toBeVisible();
  await expect(page.getByTestId("find-query")).toBeFocused();
  await expect(page.getByTestId("find-scope")).toHaveText("Slides, shapes, tables and notes");
  await expect(page.getByTestId("find-scope")).toHaveAttribute("title", /Not searched: text inside 2 groups/);

  // "Open": two groups — slide 1 (a Notes hit) and slide 2.
  await page.getByTestId("find-query").fill("Open");
  const groups = page.getByTestId("deck-search-group");
  await expect(groups).toHaveCount(2);
  await expect(groups.nth(0).getByTestId("deck-search-slide-chip")).toHaveText("1");
  await expect(groups.nth(0).getByTestId("deck-search-hit")).toHaveCount(1);
  await expect(groups.nth(0).getByTestId("deck-search-hit")).toContainText("Open with the Q3 highlight");
  await expect(groups.nth(0).getByTestId("deck-search-hit-scope")).toHaveText("Notes");
  await expect(groups.nth(1).getByTestId("deck-search-slide-chip")).toHaveText("2");
  await expect(groups.nth(1).getByTestId("deck-search-hit")).toContainText("Open Berlin office");
  await expect(groups.nth(1).locator("mark")).toHaveText("Open");
  await expect(page.getByTestId("find-count")).toHaveText("0 of 2");

  // Activating the slide-2 hit: slide 2 shows, that frame is selected and the
  // word is selected inside a text session on it.
  await groups.nth(1).getByTestId("deck-search-hit").click();
  await expect(page.getByTestId("filmstrip-slide").nth(1)).toHaveAttribute("aria-selected", "true");
  const selected = page.locator('[data-testid="slide-frame"].is-selected');
  await expect(selected).toHaveCount(1);
  await expect(selected.getByTestId("slide-text-edit")).toContainText("Open Berlin office");
  await expect.poll(() => page.evaluate(() => window.getSelection()?.toString())).toBe("Open");
  await expect(page.getByTestId("find-count")).toHaveText("2 of 2");

  // "Grouped": nothing searchable matches; the scope note says why.
  await page.getByTestId("find-query").fill("Grouped");
  await expect(page.getByTestId("deck-search-empty")).toBeVisible();
  await expect(page.getByTestId("deck-search-group")).toHaveCount(0);
  await expect(page.getByTestId("deck-search-scope")).toContainText("Not searched: text inside 2 groups");

  // Expand, Replace all "Open" → "Launch".
  await page.getByTestId("find-query").fill("Open");
  await expect(groups).toHaveCount(2);
  await page.getByTestId("find-expand").click();
  await page.getByTestId("find-replace-input").fill("Launch");
  await page.getByTestId("find-replace-all").click();
  await expect(page.getByTestId("find-status")).toHaveText("Replaced 2 occurrences");
  await expect(stage(page)).toContainText("Launch Berlin office");
  await expect(page.getByTestId("deck-search-empty")).toBeVisible();

  // ONE undo restores both — the stage text and the notes hit come back.
  await page.getByTestId("presentation-undo").click();
  await expect(stage(page)).toContainText("Open Berlin office");
  await expect(groups).toHaveCount(2);
  await expect(groups.nth(0).getByTestId("deck-search-hit")).toContainText("Open with the Q3 highlight");

  // Redo, download, reopen: both strings changed.
  await page.getByTestId("presentation-redo").click();
  await expect(stage(page)).toContainText("Launch Berlin office");
  await expect(groups).toHaveCount(0);
  const file = await downloadTo(page, "spec098-deck.pptx");

  await openFixture(page, file);
  await page.getByTestId("filmstrip-slide").nth(1).click();
  await expect(stage(page)).toContainText("Launch Berlin office");
  await page.getByTestId("filmstrip-slide").nth(0).click();
  await page.getByTestId("slides-notes-toggle").click();
  await expect(page.getByTestId("slides-notes-editor")).toHaveValue("Launch with the Q3 highlight\nPause before the revenue chart");
});

test("Find toggle, Cmd/Ctrl+H, find-next with the wrap hint, a notes hit and Escape back to the stage", async ({ page }) => {
  await openFixture(page, FIXTURE_PPTX);

  // The Find toggle sits with the panel toggles beside Transitions/Animate.
  const toggle = page.getByTestId("slides-find-toggle");
  await expect(page.getByTestId("slides-ribbon-panels").getByTestId("slides-find-toggle")).toBeVisible();
  await toggle.click();
  await expect(page.getByTestId("find-bar")).toBeVisible();
  await expect(toggle).toHaveAttribute("aria-pressed", "true");
  await page.getByTestId("find-close").click();
  await expect(page.getByTestId("find-bar")).toHaveCount(0);

  // Cmd/Ctrl+H opens it with the replace row revealed.
  await page.keyboard.press("ControlOrMeta+h");
  await expect(page.getByTestId("find-replace-row")).toBeVisible();

  // Enter steps from the bar: the notes hit first (slide 1's notes panel
  // opens with the range selected), then slide 2; past the last hit wraps.
  await page.getByTestId("find-query").fill("Open");
  await expect(page.getByTestId("deck-search-group")).toHaveCount(2);
  await page.getByTestId("find-query").press("Enter");
  await expect(page.getByTestId("find-count")).toHaveText("1 of 2");
  await expect(page.getByTestId("slides-notes-editor")).toBeVisible();
  await expect
    .poll(() => page.getByTestId("slides-notes-editor").evaluate((el) => {
      const area = el as HTMLTextAreaElement;
      return area.value.slice(area.selectionStart, area.selectionEnd);
    }))
    .toBe("Open");
  await expect(page.getByTestId("find-query")).toBeFocused();
  await page.getByTestId("find-query").press("Enter");
  await expect(page.getByTestId("find-count")).toHaveText("2 of 2");
  await expect(page.getByTestId("filmstrip-slide").nth(1)).toHaveAttribute("aria-selected", "true");
  // Off the session's focus the current hit paints as a highlight rect.
  await expect(page.locator('[data-testid="find-highlight"].current')).toHaveCount(1);
  await page.getByTestId("find-query").press("Enter");
  await expect(page.getByTestId("find-count")).toHaveText("1 of 2");
  await expect(page.getByTestId("find-status")).toHaveText("Wrapped to top");

  // An invalid expression reports inline and runs nothing.
  await page.getByTestId("find-regex").click();
  await page.getByTestId("find-query").fill("(");
  await expect(page.getByTestId("find-status")).toHaveText("Invalid regular expression");
  await expect(page.getByTestId("deck-search-results")).toHaveCount(0);

  // Escape closes the bar and hands focus back to the stage.
  await page.getByTestId("find-query").press("Escape");
  await expect(page.getByTestId("find-bar")).toHaveCount(0);
  await expect.poll(() => page.evaluate(() => !!document.activeElement?.closest('[data-testid="stage-workspace"]'))).toBe(true);
});
