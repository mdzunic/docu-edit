import { test, expect, type Locator, type Page } from "@playwright/test";
import { readFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

// SPEC-099 — the slide sorter, deck sections and slide reuse. Open mini-deck
// (so the library holds a second deck), go back, open the sections fixture,
// switch to the Sorter: six cards under two section headers. Shift+click
// selects slides 2–3, Copy, click slide 5, Paste — the two copies land after
// it inside that section. Collapse the first section, drag a slide within the
// second, then Reuse slides from "mini-deck" (two slides, destination theme)
// and read the report. Download, reopen: sections, names and order survive,
// and the reused slide's image renders.
const here = dirname(fileURLToPath(import.meta.url));
const MINI = join(here, "fixtures", "mini-deck.pptx");
const SECTIONS = join(here, "fixtures", "spec099-sections.pptx");
const SECTIONS_ODP = join(here, "fixtures", "spec099-sections.odp");

function sorter(page: Page): Locator {
  return page.getByTestId("slides-sorter");
}
function cards(page: Page): Locator {
  return sorter(page).getByTestId("sorter-slide");
}
function live(page: Page): Locator {
  return page.getByTestId("sorter-live");
}

async function openFile(page: Page, file: string): Promise<void> {
  await page.getByTestId("open-document-input").setInputFiles(file);
  await expect(page).toHaveURL(/\/slides\/.+/);
  await expect(page.getByTestId("presentation-view")).toBeVisible();
}

/** Each visible card's slide title (its accessible name ends "— Title"). */
async function cardTitles(page: Page): Promise<string[]> {
  const labels = await cards(page).evaluateAll((els) => els.map((el) => el.getAttribute("aria-label") ?? ""));
  return labels.map((label) => label.split(" — ")[1] ?? "");
}

/** Section names with their slide counts, in deck order. */
async function sectionSummary(page: Page): Promise<string[]> {
  return sorter(page)
    .getByTestId("sorter-section")
    .evaluateAll((els) =>
      els.map((el) => {
        const name = el.querySelector('[data-testid="sorter-section-name"]')?.textContent?.trim() ?? "";
        const count = el.querySelector('[data-testid="sorter-section-count"]')?.textContent?.trim() ?? "";
        return `${name}: ${count}`;
      }),
    );
}

async function downloadTo(page: Page, name: string): Promise<string> {
  const downloadPromise = page.waitForEvent("download");
  await page.getByTestId("presentation-download").click();
  const download = await downloadPromise;
  const filePath = join(tmpdir(), name);
  await download.saveAs(filePath);
  return filePath;
}

test("sorter: sections, multi-select copy/paste, collapse + drag, reuse from another deck, round trip (AC-3..17)", async ({
  page,
  context,
}) => {
  // The Paste button reads the clipboard through the async API.
  await context.grantPermissions(["clipboard-read", "clipboard-write"]);

  // A second deck in the library first: mini-deck, then back to the library.
  await page.goto("/");
  await openFile(page, MINI);
  await page.getByTestId("product-name").click();
  await expect(page).not.toHaveURL(/\/slides\//);
  await openFile(page, SECTIONS);

  // Sorter: six cards under two section headers.
  const toggle = page.getByTestId("slides-sorter-toggle");
  await expect(toggle).toHaveAttribute("aria-pressed", "false");
  await toggle.click();
  await expect(toggle).toHaveAttribute("aria-pressed", "true");
  await expect(sorter(page)).toBeVisible();
  await expect(page.getByTestId("slide-filmstrip")).toHaveCount(0);
  const listbox = sorter(page).getByRole("listbox", { name: "Slides" });
  await expect(listbox).toHaveAttribute("aria-multiselectable", "true");
  await expect(cards(page)).toHaveCount(6);
  await expect(sorter(page).getByTestId("sorter-section-name")).toHaveText(["Opening", "Details"]);
  expect(await sectionSummary(page)).toEqual(["Opening: 3 slides", "Details: 3 slides"]);
  expect(await cardTitles(page)).toEqual(["Welcome", "Agenda", "Team photo", "Budget", "Timeline", "Next steps"]);

  // Shift+click selects slides 2–3.
  await cards(page).nth(1).click();
  await cards(page).nth(2).click({ modifiers: ["Shift"] });
  await expect(cards(page).nth(1)).toHaveAttribute("aria-selected", "true");
  await expect(cards(page).nth(2)).toHaveAttribute("aria-selected", "true");
  await expect(cards(page).nth(0)).toHaveAttribute("aria-selected", "false");
  await expect(live(page)).toHaveText("2 slides selected");

  // Copy → click slide 5 → Paste: the copies land after it, inside Details.
  await page.getByTestId("sorter-copy").click();
  await expect(live(page)).toHaveText("2 slides copied");
  await cards(page).nth(4).click();
  await expect(cards(page).nth(4)).toHaveAttribute("aria-selected", "true");
  await page.getByTestId("sorter-paste").click();
  await expect(cards(page)).toHaveCount(8);
  await expect(live(page)).toHaveText("2 slides pasted");
  expect(await cardTitles(page)).toEqual([
    "Welcome",
    "Agenda",
    "Team photo",
    "Budget",
    "Timeline",
    "Agenda",
    "Team photo",
    "Next steps",
  ]);
  expect(await sectionSummary(page)).toEqual(["Opening: 3 slides", "Details: 5 slides"]);
  await expect(cards(page).nth(5)).toHaveAttribute("aria-selected", "true");
  await expect(cards(page).nth(6)).toHaveAttribute("aria-selected", "true");

  // Collapse the first section: its cards hide and its count shows.
  const openingToggle = sorter(page).getByTestId("sorter-section-toggle").first();
  await openingToggle.click();
  await expect(openingToggle).toHaveAttribute("aria-expanded", "false");
  await expect(cards(page)).toHaveCount(5);
  await expect(sorter(page).getByTestId("sorter-section-count").first()).toHaveText("3 slides");

  // Drag "Next steps" into the second section, before "Budget".
  const visible = await cardTitles(page);
  expect(visible).toEqual(["Budget", "Timeline", "Agenda", "Team photo", "Next steps"]);
  await cards(page).last().click();
  await cards(page).last().dragTo(cards(page).first(), { targetPosition: { x: 8, y: 40 } });
  expect(await cardTitles(page)).toEqual(["Next steps", "Budget", "Timeline", "Agenda", "Team photo"]);
  await expect(live(page)).toHaveText("Slide moved to position 4");
  expect(await sectionSummary(page)).toEqual(["Opening: 3 slides", "Details: 5 slides"]);

  // Reuse slides from mini-deck: two of them, with this deck's theme.
  await page.getByTestId("slides-reuse-toggle").click();
  const dialog = page.getByTestId("slides-reuse-dialog");
  await expect(dialog).toBeVisible();
  await dialog.getByTestId("reuse-deck-select").selectOption({ label: "mini-deck" });
  const reuseSlides = dialog.getByTestId("reuse-slide");
  await expect(reuseSlides).toHaveCount(4);
  await reuseSlides.nth(0).click();
  await reuseSlides.nth(3).click();
  await expect(reuseSlides.nth(3)).toHaveAttribute("aria-selected", "true");
  await expect(dialog.getByTestId("reuse-use-destination")).toBeChecked();
  await dialog.getByTestId("reuse-use-destination").check();
  await dialog.getByTestId("reuse-insert").click();
  await expect(dialog).toBeHidden();

  // The report names what it kept.
  const report = page.getByTestId("sorter-import-report");
  await expect(report).toBeVisible();
  await expect(report).toContainText("Inserted 2 slides");
  await expect(report).toContainText("Kept 2 slides: “Quarterly Review”, “Extras”.");
  expect(await cardTitles(page)).toEqual(["Next steps", "Quarterly Review", "Extras", "Budget", "Timeline", "Agenda", "Team photo"]);
  expect(await sectionSummary(page)).toEqual(["Opening: 3 slides", "Details: 7 slides"]);
  await page.getByTestId("sorter-import-report-dismiss").click();
  await expect(report).toHaveCount(0);

  // Expand Opening again and take the whole deck's order.
  await openingToggle.click();
  await expect(cards(page)).toHaveCount(10);
  const before = await cardTitles(page);
  const sectionsBefore = await sectionSummary(page);

  // Download, reopen: sections, names and order unchanged.
  const saved = await downloadTo(page, "spec099-roundtrip.pptx");
  await page.goto("/");
  await openFile(page, saved);
  await expect(page.getByTestId("slides-sorter-toggle")).toHaveAttribute("aria-pressed", "true");
  await expect(cards(page)).toHaveCount(10);
  expect(await sectionSummary(page)).toEqual(sectionsBefore);
  expect(await cardTitles(page)).toEqual(before);

  // The reused slide's image renders (on its sorter card and on the stage).
  const extras = cards(page).nth(5);
  await expect(extras).toHaveAttribute("aria-label", /Extras$/);
  const thumbImage = extras.locator("img.frame-image");
  await expect(thumbImage).toHaveCount(1);
  await expect.poll(() => thumbImage.evaluate((img) => (img as HTMLImageElement).naturalWidth)).toBeGreaterThan(0);
  await extras.click();
  await page.getByTestId("slides-sorter-toggle").click();
  await expect(page.getByTestId("slides-sorter")).toHaveCount(0);
  const stageImage = page.getByTestId("stage-workspace").locator("img.frame-image");
  await expect(stageImage).toHaveCount(1);
  await expect.poll(() => stageImage.evaluate((img) => (img as HTMLImageElement).naturalWidth)).toBeGreaterThan(0);
});

// "Add section" hands the keyboard to the new section's name: typing replaces
// it and Enter commits it. Typed through page.keyboard, not locator.press(),
// which focuses its target itself and so cannot see where the app put focus.
test("Add section focuses the new section's name for typing (AC-6)", async ({ page }) => {
  await page.goto("/");
  await openFile(page, SECTIONS);
  await page.getByTestId("slides-sorter-toggle").click();
  await expect(cards(page)).toHaveCount(6);
  const input = sorter(page).getByTestId("section-rename-input");

  // From the sorter's toolbar, at slide 5.
  await cards(page).nth(4).click();
  await sorter(page).getByTestId("sorter-add-section").click();
  await expect(input).toBeFocused();
  await expect(input).toHaveValue("New section");
  await page.keyboard.type("Appendix");
  await page.keyboard.press("Enter");
  await expect(input).toHaveCount(0);
  expect(await sectionSummary(page)).toEqual(["Opening: 3 slides", "Details: 1 slide", "Appendix: 2 slides"]);

  // From a section's menu, at slide 2.
  await cards(page).nth(1).click();
  await sorter(page).getByTestId("sorter-section-menu").first().click();
  await page.getByTestId("section-add").click();
  await expect(input).toBeFocused();
  await page.keyboard.type("Middle");
  await page.keyboard.press("Enter");
  await expect(input).toHaveCount(0);
  expect(await sectionSummary(page)).toEqual([
    "Opening: 1 slide",
    "Middle: 2 slides",
    "Details: 1 slide",
    "Appendix: 2 slides",
  ]);
});

test("sections: rename, reject an empty name, move a section with its run, remove it (AC-6)", async ({ page }) => {
  await page.goto("/");
  await openFile(page, SECTIONS);
  await page.getByTestId("slides-sorter-toggle").click();
  await expect(cards(page)).toHaveCount(6);

  // Double-click → rename; an empty name is rejected, Escape cancels, Enter commits.
  await sorter(page).getByTestId("sorter-section-name").nth(1).dblclick();
  const input = sorter(page).getByTestId("section-rename-input");
  await expect(input).toBeFocused();
  await input.fill("   ");
  await input.press("Enter");
  await expect(input).toBeVisible();
  await input.press("Escape");
  await expect(input).toHaveCount(0);
  await expect(sorter(page).getByTestId("sorter-section-name").nth(1)).toHaveText("Details");
  await sorter(page).getByTestId("sorter-section-menu").nth(1).click();
  await page.getByTestId("section-rename").click();
  await sorter(page).getByTestId("section-rename-input").fill("  Deep dive  ");
  await sorter(page).getByTestId("section-rename-input").press("Enter");
  await expect(sorter(page).getByTestId("sorter-section-name")).toHaveText(["Opening", "Deep dive"]);

  // Move the second section up: its whole run moves with it.
  await sorter(page).getByTestId("sorter-section-menu").nth(1).click();
  await page.getByTestId("section-move-up").click();
  await expect(sorter(page).getByTestId("sorter-section-name")).toHaveText(["Deep dive", "Opening"]);
  expect(await cardTitles(page)).toEqual(["Budget", "Timeline", "Next steps", "Welcome", "Agenda", "Team photo"]);

  // Undo restores the order (one step per section command).
  await page.getByTestId("presentation-undo").click();
  await expect(sorter(page).getByTestId("sorter-section-name")).toHaveText(["Opening", "Deep dive"]);

  // Remove the first section (its slides join the next).
  await sorter(page).getByTestId("sorter-section-menu").first().click();
  await page.getByTestId("section-remove").click();
  await expect(sorter(page).getByTestId("sorter-section-name")).toHaveText(["Deep dive"]);
  await expect(cards(page)).toHaveCount(6);

  // A paste with no slides on the clipboard is a no-op that says so.
  await cards(page).first().click();
  await page.evaluate(() => {
    const data = new DataTransfer();
    data.setData("text/plain", "Just some text");
    window.dispatchEvent(new ClipboardEvent("paste", { clipboardData: data, bubbles: true }));
  });
  await expect(live(page)).toHaveText("The clipboard doesn't hold any slides");
  await expect(cards(page)).toHaveCount(6);
});

// The odp mapping says once that Impress will not show sections; a download
// with no edits is byte-identical to the source in both formats, and a
// section-only edit survives a reopen.
test("odp sections note, no-edit downloads stay byte-identical, a rename round-trips (AC-8, AC-15)", async ({
  page,
}) => {
  for (const file of [SECTIONS, SECTIONS_ODP]) {
    await page.goto("/");
    await openFile(page, file);
    const toggle = page.getByTestId("slides-sorter-toggle");
    if ((await toggle.getAttribute("aria-pressed")) !== "true") {
      await toggle.click();
    }
    await expect(cards(page)).toHaveCount(6);
    await expect(sorter(page).getByTestId("sorter-section-name")).toHaveText(["Opening", "Details"]);
    await expect(page.getByTestId("sorter-odp-note")).toHaveCount(file === SECTIONS_ODP ? 1 : 0);
    const ext = file === SECTIONS_ODP ? "odp" : "pptx";
    const untouched = await downloadTo(page, `spec099-noedit.${ext}`);
    expect(readFileSync(untouched).equals(readFileSync(file))).toBe(true);

    await sorter(page).getByTestId("sorter-section-name").first().dblclick();
    await expect(sorter(page).getByTestId("section-rename-input")).toBeFocused();
    await page.keyboard.press("ControlOrMeta+a");
    await page.keyboard.type("Prologue");
    await page.keyboard.press("Enter");
    await expect(sorter(page).getByTestId("sorter-section-name")).toHaveText(["Prologue", "Details"]);
    const renamed = await downloadTo(page, `spec099-renamed.${ext}`);
    expect(readFileSync(renamed).equals(readFileSync(file))).toBe(false);

    await page.goto("/");
    await openFile(page, renamed);
    await expect(cards(page)).toHaveCount(6);
    expect(await sectionSummary(page)).toEqual(["Prologue: 3 slides", "Details: 3 slides"]);
    expect(await cardTitles(page)).toEqual(["Welcome", "Agenda", "Team photo", "Budget", "Timeline", "Next steps"]);
  }
});
