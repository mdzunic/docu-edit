import { test, expect, type Page } from "@playwright/test";

// SPEC-042 — the custom styles manager: a "Styles" side pane lists every
// style grouped by type with in-use tracking (AC-1, AC-2); New/Modify/
// Delete manage the catalogue (AC-3..AC-8, AC-11); character styles apply
// independently of the paragraph style and both show at the caret (AC-9,
// AC-10); a colliding name is rejected inline (AC-7 in the intake's own
// numbering — "name collision" AC); custom styles round-trip through
// .docx/.odt (AC-13, AC-14); and "Import styles from another document"
// merges a second file's catalogue with per-conflict resolution (AC-19..24).

async function createDocument(page: Page): Promise<void> {
  await page.goto("/");
  await page.getByTestId("new-document").first().click();
  await expect(page).toHaveURL(/\/doc\/.+/);
}

const editor = (page: Page) => page.locator('[contenteditable="true"]').first();

async function openStylesPane(page: Page): Promise<void> {
  await page.getByTestId("menu-title-view").click();
  await page.getByTestId("view-menu-styles-pane").click();
  await expect(page.getByTestId("styles-pane")).toBeVisible();
}

async function newStyle(
  page: Page,
  type: "paragraph" | "character" | "list" | "table",
  name: string,
  options: { basedOn?: string; bold?: boolean; italic?: boolean; color?: string } = {},
): Promise<void> {
  await page.getByTestId(`new-style-${type}`).click();
  await expect(page.getByTestId("style-editor-dialog")).toBeVisible();
  await page.getByTestId("style-name-input").fill(name);
  if (options.basedOn) {
    await page.getByTestId("style-based-on-select").selectOption({ label: options.basedOn });
  }
  if (options.bold) {
    await page.getByTestId("style-bold-checkbox").check();
  }
  if (options.italic) {
    await page.getByTestId("style-italic-checkbox").check();
  }
  if (options.color) {
    await page.getByTestId("style-color-input").fill(options.color);
  }
  await page.getByTestId("style-editor-save").click();
  await expect(page.getByTestId("style-editor-dialog")).not.toBeVisible();
}

function styleRow(page: Page, name: string) {
  return page.locator(`[data-testid^="style-row-"]:has-text("${name}")`);
}

test.describe("Styles pane (AC-1, AC-2)", () => {
  test("lists styles grouped by type, and marks an applied style in use", async ({ page }) => {
    await createDocument(page);
    await editor(page).click();
    await page.keyboard.type("Some text");
    await openStylesPane(page);

    await expect(page.getByTestId("styles-group-paragraph")).toBeVisible();
    await expect(page.getByTestId("styles-group-character")).toBeVisible();
    await expect(page.getByTestId("styles-group-list")).toBeVisible();
    await expect(page.getByTestId("styles-group-table")).toBeVisible();
    // Every built-in paragraph style is listed even before anything custom
    // is created.
    await expect(styleRow(page, "Heading 1")).toBeVisible();

    await newStyle(page, "paragraph", "My Body");
    await expect(styleRow(page, "My Body").getByTestId(/style-in-use-/)).toBeVisible();
  });
});

test.describe("New / Modify style (AC-3, AC-4, AC-5, AC-11)", () => {
  test("modifying a parent style re-bakes a child style's formatting without breaking its identity", async ({
    page,
  }) => {
    await createDocument(page);
    await editor(page).click();
    await page.keyboard.type("Parent paragraph.");
    await page.keyboard.press("Enter");
    await page.keyboard.type("Child paragraph.");
    await openStylesPane(page);

    await page.keyboard.press("Control+Home");
    await editor(page).click();
    await page.keyboard.press("Control+Home");
    await newStyle(page, "paragraph", "Parent Style", { bold: true });

    await page.keyboard.press("Control+End");
    await newStyle(page, "paragraph", "Child Style", { basedOn: "Parent Style", italic: true });

    // Both styles are in use before the modify.
    await expect(styleRow(page, "Parent Style").getByTestId(/style-in-use-/)).toBeVisible();
    await expect(styleRow(page, "Child Style").getByTestId(/style-in-use-/)).toBeVisible();

    await styleRow(page, "Parent Style").getByTestId(/style-modify-/).click();
    await expect(page.getByTestId("style-editor-dialog")).toBeVisible();
    await page.getByTestId("style-color-input").fill("#0000ff");
    await page.getByTestId("style-editor-save").click();
    await expect(page.getByTestId("style-editor-dialog")).not.toBeVisible();

    // The regression under test: Child Style must still be in use — its
    // paragraph must not have been silently re-pointed at Parent Style.
    await expect(styleRow(page, "Child Style").getByTestId(/style-in-use-/)).toBeVisible();
    await expect(styleRow(page, "Parent Style").getByTestId(/style-in-use-/)).toBeVisible();

    // The caret's readout on the child paragraph still names Child Style.
    await editor(page).click();
    await page.keyboard.press("Control+End");
    await expect(page.getByTestId("caret-paragraph-style")).toHaveText("Child Style");
  });

  test("rejects a name that collides with an existing style, inline", async ({ page }) => {
    await createDocument(page);
    await editor(page).click();
    await page.keyboard.type("Text");
    await openStylesPane(page);

    await page.getByTestId("new-style-paragraph").click();
    await page.getByTestId("style-name-input").fill("Normal");
    await expect(page.getByTestId("style-name-error")).toBeVisible();
    await expect(page.getByTestId("style-editor-save")).toBeDisabled();
    await page.getByTestId("style-editor-cancel").click();
  });
});

test.describe("Delete style (AC-7, AC-8)", () => {
  test("asks for confirmation and reverts paragraphs to Normal", async ({ page }) => {
    await createDocument(page);
    await editor(page).click();
    await page.keyboard.type("Styled paragraph.");
    await openStylesPane(page);

    await page.keyboard.press("Control+Home");
    await newStyle(page, "paragraph", "Temp Style", { bold: true });

    await styleRow(page, "Temp Style").getByTestId(/style-delete-/).click();
    await expect(page.getByTestId("style-delete-dialog")).toBeVisible();
    await expect(page.getByTestId("style-delete-in-use-warning")).toBeVisible();
    await page.getByTestId("style-delete-confirm").click();
    await expect(page.getByTestId("style-delete-dialog")).not.toBeVisible();

    await expect(styleRow(page, "Temp Style")).toHaveCount(0);
    await expect(page.getByTestId("caret-paragraph-style")).toHaveText("Normal");
  });
});

test.describe("Character styles (AC-9, AC-10)", () => {
  test("applies independently of the paragraph style and both show at the caret", async ({ page }) => {
    await createDocument(page);
    await editor(page).click();
    await page.keyboard.type("A run of text.");
    await openStylesPane(page);

    await page.keyboard.press("Home");
    await page.keyboard.down("Shift");
    for (let i = 0; i < 6; i++) {
      await page.keyboard.press("ArrowRight");
    }
    await page.keyboard.up("Shift");

    await newStyle(page, "character", "Emphasis Style", { italic: true, color: "#cc0000" });
    await expect(styleRow(page, "Emphasis Style").getByTestId(/style-in-use-/)).toBeVisible();

    await page.keyboard.press("Home");
    await page.keyboard.press("ArrowRight");
    await expect(page.getByTestId("caret-character-style")).toHaveText("Emphasis Style");
    await expect(page.getByTestId("caret-paragraph-style")).toHaveText("Normal");

    await page.keyboard.press("End");
    await expect(page.getByTestId("caret-character-style")).toHaveText("None");
  });
});

test.describe("Custom style round trip (AC-13, AC-14)", () => {
  test("a custom paragraph style survives a .docx download and reopen", async ({ page }) => {
    await createDocument(page);
    await editor(page).click();
    await page.keyboard.type("Custom styled paragraph.");
    await openStylesPane(page);
    await page.keyboard.press("Control+Home");
    await newStyle(page, "paragraph", "RoundTrip Style", { bold: true });

    await page.getByTestId("download-menu").click();
    const [download] = await Promise.all([
      page.waitForEvent("download"),
      page.getByTestId("download-docx").click(),
    ]);
    const path = await download.path();
    expect(path).toBeTruthy();

    await page.goto("/");
    await page.getByTestId("open-document-input").setInputFiles(path!);
    await expect(page).toHaveURL(/\/doc\/.+/);
    await openStylesPane(page);
    await expect(styleRow(page, "RoundTrip Style")).toBeVisible();
    await page.keyboard.press("Control+Home");
    await expect(page.getByTestId("caret-paragraph-style")).toHaveText("RoundTrip Style");
  });
});

test.describe("Import styles from another document (AC-19..24)", () => {
  test("merges non-conflicting styles automatically and lets the user resolve a conflict", async ({ page }) => {
    // Source document: a custom style whose name collides with the target's
    // own "Shared Name", plus one that does not.
    await createDocument(page);
    await editor(page).click();
    await page.keyboard.type("Source doc text.");
    await openStylesPane(page);
    await page.keyboard.press("Control+Home");
    await newStyle(page, "paragraph", "Shared Name", { bold: true });
    await newStyle(page, "paragraph", "Imported Unique Style", { italic: true });

    await page.getByTestId("download-menu").click();
    const [sourceDownload] = await Promise.all([
      page.waitForEvent("download"),
      page.getByTestId("download-docx").click(),
    ]);
    const sourcePath = await sourceDownload.path();
    expect(sourcePath).toBeTruthy();

    // Target document: its own "Shared Name" with different formatting.
    await createDocument(page);
    await editor(page).click();
    await page.keyboard.type("Target doc text.");
    await openStylesPane(page);
    await page.keyboard.press("Control+Home");
    await newStyle(page, "paragraph", "Shared Name", { italic: true });

    await page.getByTestId("menu-title-format").click();
    await page.getByTestId("format-menu-import-styles").click();
    await expect(page.getByTestId("import-styles-dialog")).toBeVisible();
    await page.getByTestId("import-styles-file-input").setInputFiles(sourcePath!);
    await expect(page.getByTestId("import-styles-non-conflicting-count")).toContainText("1");
    const conflictRow = page.locator('[data-testid^="import-conflict-"]', { hasText: "Shared Name" });
    await expect(conflictRow).toBeVisible();
    await conflictRow.getByTestId(/-replace$/).check();

    await page.getByTestId("import-styles-import").click();
    // Both documents also carry the full built-in set (materialized into
    // the catalogue the moment any custom style is created) — those collide
    // by name too and default to "keep", hence 11 skipped alongside the one
    // real conflict this test resolves as "replace".
    await expect(page.getByTestId("import-styles-summary")).toContainText("Added 1, replaced 1, skipped 11");
    await page.getByTestId("import-styles-done").click();
    await expect(page.getByTestId("import-styles-dialog")).not.toBeVisible();

    await expect(styleRow(page, "Imported Unique Style")).toBeVisible();
  });
});
