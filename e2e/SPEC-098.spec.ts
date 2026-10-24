import { test, expect, type Page } from "@playwright/test";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

const here = dirname(fileURLToPath(import.meta.url));
const FIXTURE_PPTX = join(here, "fixtures", "mini-deck.pptx");
// Each slide carries a rounded-rectangle ShapeFrame holding "Press me".
const SHAPE_TEXT_PPTX = join(here, "fixtures", "spec105-animation-depth.pptx");

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
  await expect(page.getByTestId("presentation-empty")).toHaveCount(0);
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

/** The squiggle for `word` on the stage. */
function underline(page: Page, word: string) {
  return page.locator(`[data-testid="spell-underline"][data-word="${word}"]`);
}

async function rightClick(page: Page, word: string): Promise<void> {
  const box = await underline(page, word).boundingBox();
  if (!box) throw new Error(`no underline for ${word}`);
  await page.mouse.click(box.x + box.width / 2, box.y + box.height / 2, { button: "right" });
}

test("spell check on the stage: squiggles, the suggestions menu, a correction that keeps its case, and noProof", async ({ page }) => {
  await openFixture(page, FIXTURE_PPTX);
  await page.getByTestId("filmstrip-slide").nth(1).click();
  await expect(page.getByTestId("spell-underlines")).toBeVisible();
  await expect(underline(page, "Berlin")).toHaveCount(1);

  // Right-click opens SPEC-035's menu; Ignore all clears the word.
  await rightClick(page, "Berlin");
  const menu = page.getByTestId("spell-context-menu");
  await expect(menu).toBeVisible();
  await expect(menu.getByTestId("spell-ignore")).toBeVisible();
  await expect(menu.getByTestId("spell-add-to-dictionary")).toBeVisible();
  expect(await menu.getByTestId("spell-suggestion").count()).toBeLessThanOrEqual(5);
  await menu.getByTestId("spell-ignore-all").click();
  await expect(underline(page, "Berlin")).toHaveCount(0);

  // A misspelling typed into the title: the suggestion keeps its capital, and
  // the correction is one undo entry.
  const title = page.locator('[data-testid="slide-frame"]').first();
  await title.dblclick();
  await page.keyboard.press("End");
  await page.keyboard.type(" Teh");
  await expect(underline(page, "Teh")).toHaveCount(1);
  await rightClick(page, "Teh");
  await menu.getByTestId("spell-suggestion").filter({ hasText: /^\s*the\s*$/ }).click();
  await expect(page.getByTestId("slide-text-edit")).toHaveText("Roadmap The");
  await expect(underline(page, "Teh")).toHaveCount(0);
  await page.keyboard.press("Escape");
  await expect(page.getByTestId("slide-text-edit")).toHaveCount(0);
  await page.getByTestId("presentation-undo").click();
  await expect(stage(page)).toContainText("Roadmap Teh");
  await page.getByTestId("presentation-redo").click();
  await expect(stage(page)).toContainText("Roadmap The");

  // "Do not check spelling" on a selected frame clears its squiggles.
  const body = page.locator('[data-testid="slide-frame"]').nth(1);
  await body.click();
  await expect(underline(page, "MVP")).toHaveCount(1);
  await page.getByTestId("language-status-button").click();
  await page.getByTestId("language-no-proof").check();
  await expect(underline(page, "MVP")).toHaveCount(0);
});

test("a table-cell hit carries its Table rNcN tag, opens a session in that cell, and Replace rewrites it", async ({ page }) => {
  await openFixture(page, FIXTURE_PPTX);
  await page.keyboard.press("ControlOrMeta+h");
  await page.getByTestId("find-query").fill("Cell");
  const group = page.getByTestId("deck-search-group");
  await expect(group).toHaveCount(1);
  await expect(group.getByTestId("deck-search-slide-chip")).toHaveText("4");
  await expect(group.getByTestId("deck-search-hit-scope")).toHaveText("Table r1c1");
  await group.getByTestId("deck-search-hit").click();
  await expect(page.getByTestId("filmstrip-slide").nth(3)).toHaveAttribute("aria-selected", "true");
  await expect(page.getByTestId("slide-table-cell-edit")).toBeVisible();
  await expect.poll(() => page.evaluate(() => window.getSelection()?.toString())).toBe("Cell");

  await page.getByTestId("find-replace-input").fill("Box");
  await page.getByTestId("find-replace").click();
  await expect(page.getByTestId("find-count")).toHaveText("0 of 0");
  await expect(stage(page).getByTestId("slide-table-cell-0-0")).toHaveText("Box");
});

test("a shape's hit opens a text session on that shape with the word selected", async ({ page }) => {
  await openFixture(page, SHAPE_TEXT_PPTX);
  await page.keyboard.press("ControlOrMeta+h");
  await page.getByTestId("find-query").fill("Press me");
  const groups = page.getByTestId("deck-search-group");
  await expect(groups).toHaveCount(4);
  await page.getByTestId("find-query").fill("Press");
  await expect(groups).toHaveCount(4);
  await groups.nth(0).getByTestId("deck-search-hit").click();

  await expect(page.getByTestId("filmstrip-slide").nth(0)).toHaveAttribute("aria-selected", "true");
  const selected = page.locator('[data-testid="slide-frame"].is-selected');
  await expect(selected).toHaveCount(1);
  await expect(selected).toHaveClass(/shape-frame/);
  const host = selected.getByTestId("slide-text-edit");
  await expect(host).toHaveText("Press me");
  await expect.poll(() => page.evaluate(() => window.getSelection()?.toString())).toBe("Press");
  // The shape's text keeps its centring inside the session.
  await expect(host.locator(".ce-p")).toHaveCSS("text-align", "center");

  // Replace rewrites the shape's text and steps on to slide 2's shape.
  await page.getByTestId("find-replace-input").fill("Push");
  await page.getByTestId("find-replace").click();
  await expect(page.getByTestId("find-count")).toHaveText("1 of 3");
  await expect(page.getByTestId("filmstrip-slide").nth(1)).toHaveAttribute("aria-selected", "true");
  await page.getByTestId("find-query").press("Escape");
  await expect(page.getByTestId("slide-text-edit")).toHaveCount(0);
  await page.getByTestId("filmstrip-slide").nth(0).click();
  await expect(stage(page).locator(".shape-frame")).toHaveText("Push me");
});

test("Escape from the bar after stepping into a text frame ends that session and leaves the keyboard on the stage", async ({ page }) => {
  await openFixture(page, FIXTURE_PPTX);
  await page.keyboard.press("ControlOrMeta+f");
  await page.getByTestId("find-query").fill("Open");
  await expect(page.getByTestId("deck-search-group")).toHaveCount(2);
  await page.getByTestId("find-query").press("Enter");
  await page.getByTestId("find-query").press("Enter");
  await expect(page.getByTestId("find-count")).toHaveText("2 of 2");
  await expect(page.getByTestId("slide-text-edit")).toBeVisible();
  await expect(page.getByTestId("find-query")).toBeFocused();

  await page.getByTestId("find-query").press("Escape");
  await expect(page.getByTestId("find-bar")).toHaveCount(0);
  await expect(page.getByTestId("slide-text-edit")).toHaveCount(0);
  await expect(page.locator('[data-testid="slide-frame"].is-selected')).toBeFocused();
  // The stage's keys work again: Enter opens a session on the selected frame.
  await page.keyboard.press("Enter");
  await expect(page.getByTestId("slide-text-edit")).toBeVisible();
});
