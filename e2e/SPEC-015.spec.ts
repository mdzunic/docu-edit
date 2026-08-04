import { test, expect, type Page } from "@playwright/test";
import { LINE_START } from "./keys";

// SPEC-015 — undo/redo: shortcuts, toolbar buttons, typing grouped into one
// step and broken by a caret move / pause / non-typing command, formatting
// always its own step, the restored selection, and the redo stack cleared
// by a new edit. Selection setup follows e2e/SPEC-012.spec.ts.

async function createDocument(page: Page): Promise<string> {
  await page.goto("/");
  await page.getByTestId("new-document").click();
  await expect(page).toHaveURL(/\/doc\/.+/);
  const url = page.url();
  return url.slice(url.lastIndexOf("/") + 1);
}

const content = (page: Page) => page.getByTestId("document-content");
const undoButton = (page: Page) => page.getByTestId("undo-button");
const redoButton = (page: Page) => page.getByTestId("redo-button");

async function selectRange(page: Page, from: number, to: number): Promise<void> {
  await page.keyboard.press(LINE_START);
  for (let i = 0; i < from; i++) {
    await page.keyboard.press("ArrowRight");
  }
  await page.keyboard.down("Shift");
  for (let i = from; i < to; i++) {
    await page.keyboard.press("ArrowRight");
  }
  await page.keyboard.up("Shift");
}

test("Undo/Redo start disabled with a generic tooltip, and enable once there is something to act on (AC-4..8)", async ({
  page,
}) => {
  await createDocument(page);
  await expect(undoButton(page)).toBeDisabled();
  await expect(redoButton(page)).toBeDisabled();
  await expect(undoButton(page)).toHaveAttribute("title", "Undo (Ctrl+Z)");
  await expect(redoButton(page)).toHaveAttribute("title", "Redo (Ctrl+Shift+Z)");

  await content(page).click();
  await page.keyboard.type("hello");
  await expect(undoButton(page)).toBeEnabled();
  await expect(undoButton(page)).toHaveAttribute("title", "Undo typing (Ctrl+Z)");
  await expect(redoButton(page)).toBeDisabled();
});

test("Ctrl+Z undoes, Ctrl+Shift+Z and Ctrl+Y redo (AC-1, AC-2, AC-3)", async ({ page }) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("abc");
  await expect(content(page)).toHaveText("abc");

  await page.keyboard.press("Control+z");
  await expect(content(page)).toHaveText("");

  await page.keyboard.press("Control+Shift+z");
  await expect(content(page)).toHaveText("abc");

  await page.keyboard.press("Control+z");
  await expect(content(page)).toHaveText("");
  await page.keyboard.press("Control+y");
  await expect(content(page)).toHaveText("abc");
});

test("the toolbar buttons perform undo and redo (AC-4)", async ({ page }) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("abc");
  await expect(content(page)).toHaveText("abc");

  await undoButton(page).click();
  await expect(content(page)).toHaveText("");
  await redoButton(page).click();
  await expect(content(page)).toHaveText("abc");
});

test("consecutive typed characters group into one undo step, broken by a caret move (AC-9, AC-10)", async ({
  page,
}) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("hello");
  await expect(content(page)).toHaveText("hello");

  // One undo removes the whole run of consecutive typing.
  await page.keyboard.press("Control+z");
  await expect(content(page)).toHaveText("");

  await page.keyboard.press("Control+Shift+z");
  await page.keyboard.press(LINE_START);
  await page.keyboard.type("X");
  await expect(content(page)).toHaveText("Xhello");

  // The caret move before typing "X" broke the group: one undo removes only
  // "X", a second is needed for the original "hello" run.
  await page.keyboard.press("Control+z");
  await expect(content(page)).toHaveText("hello");
  await page.keyboard.press("Control+z");
  await expect(content(page)).toHaveText("");
});

test("a pause of 2 s breaks the typing group (AC-11)", async ({ page }) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("ab");
  await page.waitForTimeout(2100);
  await page.keyboard.type("cd");
  await expect(content(page)).toHaveText("abcd");

  await page.keyboard.press("Control+z");
  await expect(content(page)).toHaveText("ab");
  await page.keyboard.press("Control+z");
  await expect(content(page)).toHaveText("");
});

test("a formatting command breaks the typing group and is always its own undo step (AC-12, AC-13)", async ({
  page,
}) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("hi");
  await selectRange(page, 0, 2);
  await page.keyboard.press("Control+b");
  await expect(undoButton(page)).toHaveAttribute("title", "Undo bold (Ctrl+Z)");
  await selectRange(page, 0, 2);
  await page.keyboard.press("Control+i");
  await expect(undoButton(page)).toHaveAttribute("title", "Undo italic (Ctrl+Z)");

  // Bold and italic never merge, even applied over the same selection.
  await page.keyboard.press("Control+z");
  await expect(undoButton(page)).toHaveAttribute("title", "Undo bold (Ctrl+Z)");
  await page.keyboard.press("Control+z");
  await expect(undoButton(page)).toHaveAttribute("title", "Undo typing (Ctrl+Z)");
});

test("undo restores the selection that was replaced (AC-14)", async ({ page }) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("Hello world");
  await selectRange(page, 6, 11);
  await page.keyboard.type("X");
  await expect(content(page)).toHaveText("Hello X");

  await page.keyboard.press("Control+z");
  await expect(content(page)).toHaveText("Hello world");
  // The restored selection covers "world" again, so typing over it replaces
  // it rather than inserting at a bare caret.
  await page.keyboard.type("Y");
  await expect(content(page)).toHaveText("Hello Y");
});

test("a new edit after undo clears the redo stack (AC-15)", async ({ page }) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("abc");
  await page.keyboard.press("Control+z");
  await expect(redoButton(page)).toBeEnabled();

  await page.keyboard.type("x");
  await expect(redoButton(page)).toBeDisabled();
});

test("undo history is per document (AC-18, AC-19)", async ({ page }) => {
  await createDocument(page);
  await content(page).click();
  await page.keyboard.type("hello");
  await expect(undoButton(page)).toBeEnabled();

  await createDocument(page);
  await expect(undoButton(page)).toBeDisabled();
});
