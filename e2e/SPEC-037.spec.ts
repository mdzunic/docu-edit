import { test, expect, type Page } from "@playwright/test";
import { MOD } from "./keys";

// SPEC-037 — document outline navigator: a collapsible left sidebar built
// from Heading 1-6, indented by level, whose highlight follows the caret;
// clicking or keyboard-navigating an entry jumps to that heading; entries
// with children collapse/expand; a heading can be dragged among its siblings
// (one undo step); a body chevron collapses a heading's content without
// touching the file; the pane toggles from the View menu and its open state
// persists per user; the pane is a nav landmark with a tree role.

const content = (page: Page) => page.getByTestId("document-content");
const pane = (page: Page) => page.getByTestId("outline-pane");
const tree = (page: Page) => page.getByTestId("outline-tree");
const caretStatus = (page: Page) => page.getByTestId("caret-status");

function paneItem(page: Page, name: string) {
  return pane(page).getByRole("treeitem").filter({ hasText: name });
}

async function createDocument(page: Page): Promise<void> {
  await page.goto("/");
  await page.getByTestId("new-document").click();
  await expect(page).toHaveURL(/\/doc\/.+/);
}

/** Types a heading paragraph via the SPEC-021 shortcut, then a body
 * paragraph — mirrors e2e/SPEC-033.spec.ts's helper. */
async function typeHeadingThenBody(
  page: Page,
  headingText: string,
  level: number,
  bodyText: string,
): Promise<void> {
  await page.keyboard.type(headingText);
  await page.keyboard.press(`${MOD}+Alt+${level}`);
  await page.keyboard.press("Enter");
  await page.keyboard.press(`${MOD}+Alt+0`);
  await page.keyboard.type(bodyText);
  await page.keyboard.press("Enter");
}

// Since SPEC-065 the outline is a tab in the consolidated side panel and is
// closed by default; open it from the View menu (its command id is unchanged).
async function openOutline(page: Page): Promise<void> {
  await page.getByTestId("menu-title-view").click();
  await page.getByTestId("view-menu-outline").click();
}

async function seedHeadings(page: Page): Promise<void> {
  await content(page).click();
  await typeHeadingThenBody(page, "Chapter One", 1, "Body under chapter one.");
  await typeHeadingThenBody(page, "Section A", 2, "Body under section A.");
  await typeHeadingThenBody(page, "Chapter Two", 1, "Body under chapter two.");
  await openOutline(page);
}

test("the pane is a nav landmark with a tree, opened from the View menu, listing headings indented by level (AC-1, AC-2, AC-15, AC-20)", async ({
  page,
}) => {
  await createDocument(page);
  await seedHeadings(page);
  const landmark = pane(page);
  await expect(landmark).toBeVisible();
  expect((await landmark.evaluate((el) => el.tagName)).toLowerCase()).toBe("nav");
  await expect(tree(page)).toHaveAttribute("role", "tree");
  const items = tree(page).getByRole("treeitem");
  await expect(items).toHaveCount(3);
  await expect(paneItem(page, "Chapter One")).toHaveAttribute("aria-level", "1");
  await expect(paneItem(page, "Section A")).toHaveAttribute("aria-level", "2");
  await expect(paneItem(page, "Chapter Two")).toHaveAttribute("aria-level", "1");
});

test("a document with no headings shows the empty-state message (AC-12)", async ({ page }) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("Just a plain paragraph.");
  await openOutline(page);
  await expect(page.getByTestId("outline-empty")).toHaveText("Headings you add will appear here");
});

test("the highlight follows the caret (AC-3, AC-4)", async ({ page }) => {
  await createDocument(page);
  await seedHeadings(page);
  // Caret currently sits after the last heading's body — Chapter Two is active.
  await expect(paneItem(page, "Chapter Two")).toHaveClass(/is-active/);
  // Move the caret up into Chapter One's territory.
  await content(page).getByRole("heading", { name: "Chapter One" }).click();
  await expect(paneItem(page, "Chapter One")).toHaveClass(/is-active/);
  await expect(paneItem(page, "Chapter Two")).not.toHaveClass(/is-active/);
});

test("clicking an entry scrolls to the heading and places the caret at its start (AC-5, AC-6)", async ({
  page,
}) => {
  await createDocument(page);
  await seedHeadings(page);
  await paneItem(page, "Section A").click();
  await expect(paneItem(page, "Section A")).toHaveClass(/is-active/);
  // The caret landed on the "Section A" heading paragraph (paragraph 3).
  await expect(caretStatus(page)).toContainText("Paragraph 3");
  await expect(caretStatus(page)).toContainText("character 1");
});

test("entries with children collapse and expand in the pane (AC-7, AC-8)", async ({ page }) => {
  await createDocument(page);
  await seedHeadings(page);
  // Chapter One has a child (Section A); Chapter Two does not.
  await expect(tree(page).getByRole("treeitem")).toHaveCount(3);
  await paneItem(page, "Chapter One").getByRole("button", { name: "Collapse" }).click();
  await expect(paneItem(page, "Section A")).toHaveCount(0);
  await expect(tree(page).getByRole("treeitem")).toHaveCount(2);
  await paneItem(page, "Chapter One").getByRole("button", { name: "Expand" }).click();
  await expect(paneItem(page, "Section A")).toHaveCount(1);
});

test("the pane is keyboard navigable — arrows move, Enter jumps (AC-18, AC-19)", async ({ page }) => {
  await createDocument(page);
  await seedHeadings(page);
  await tree(page).focus();
  // Home lands on the first row (AC-19a); one ArrowDown then selects
  // "Section A" and Enter jumps to it.
  await page.keyboard.press("Home");
  await page.keyboard.press("ArrowDown");
  await page.keyboard.press("Enter");
  await expect(paneItem(page, "Section A")).toHaveClass(/is-active/);
  await expect(caretStatus(page)).toContainText("Paragraph 3");
});

test("dragging a heading reorders it among its siblings as one undo step (AC-9, AC-10)", async ({
  page,
}) => {
  await createDocument(page);
  await seedHeadings(page);
  const headingsBefore = await content(page).getByRole("heading").allInnerTexts();
  expect(headingsBefore).toEqual(["Chapter One", "Section A", "Chapter Two"]);
  // Drag "Chapter Two" to just before "Chapter One" (both level 1).
  await paneItem(page, "Chapter Two").dragTo(paneItem(page, "Chapter One"), {
    targetPosition: { x: 30, y: 2 },
  });
  await expect(async () => {
    const headings = await content(page).getByRole("heading").allInnerTexts();
    expect(headings).toEqual(["Chapter Two", "Chapter One", "Section A"]);
  }).toPass();
  // One undo restores the original order.
  await content(page).click();
  await page.keyboard.press(`${MOD}+z`);
  await expect(async () => {
    const headings = await content(page).getByRole("heading").allInnerTexts();
    expect(headings).toEqual(["Chapter One", "Section A", "Chapter Two"]);
  }).toPass();
});

test("the View menu toggles the pane and the open state persists per user (AC-13, AC-14)", async ({
  page,
}) => {
  await createDocument(page);
  await seedHeadings(page);
  await expect(pane(page)).toBeVisible();
  // Toggle off via the View menu.
  await page.getByTestId("menu-title-view").click();
  await page.getByTestId("view-menu-outline").click();
  await expect(pane(page)).toHaveCount(0);
  // The closed state survives a reload.
  await page.reload();
  await expect(pane(page)).toHaveCount(0);
  // Toggle back on; that too persists.
  await page.getByTestId("menu-title-view").click();
  await page.getByTestId("view-menu-outline").click();
  await expect(pane(page)).toBeVisible();
  await page.reload();
  await expect(pane(page)).toBeVisible();
});

test("a body chevron collapses a heading's content, and the content is preserved (AC-16, AC-17)", async ({
  page,
}) => {
  await createDocument(page);
  await seedHeadings(page);
  const body = content(page).getByText("Body under chapter one.");
  await expect(body).toBeVisible();
  // Collapse Chapter One from the body chevron.
  await content(page)
    .getByRole("heading", { name: "Chapter One" })
    .getByRole("button", { name: "Collapse heading" })
    .click();
  await expect(body).toBeHidden();
  // Its own sub-heading is hidden too.
  await expect(content(page).getByRole("heading", { name: "Section A" })).toBeHidden();
  // The hidden state is view-only: expanding brings the content back verbatim.
  await content(page)
    .getByRole("heading", { name: "Chapter One" })
    .getByRole("button", { name: "Expand heading" })
    .click();
  await expect(content(page).getByText("Body under chapter one.")).toBeVisible();
});
