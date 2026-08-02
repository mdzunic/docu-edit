import { test, expect } from "@playwright/test";

// SPEC-002 — create a new blank document.

test("clicking 'New document' creates a document and navigates to /doc/:id (AC-1)", async ({
  page,
}) => {
  await page.goto("/");
  await page.getByRole("button", { name: "New document" }).click();
  await expect(page).toHaveURL(/\/doc\/.+/);
});

test("a new document defaults the title to 'Untitled document' in the app bar (AC-4, AC-5)", async ({
  page,
}) => {
  await page.goto("/");
  await page.getByRole("button", { name: "New document" }).click();
  await expect(page.getByTestId("document-title")).toHaveText(
    "Untitled document",
  );
});

test("a new document renders a white, centred, margined page (AC-2, AC-3)", async ({
  page,
}) => {
  await page.goto("/");
  await page.getByRole("button", { name: "New document" }).click();

  const canvas = (await page.getByTestId("canvas").boundingBox())!;
  const doc = (await page.getByTestId("document-page").boundingBox())!;
  const margin = (await page.getByTestId("page-margin").boundingBox())!;

  await expect(page.getByTestId("document-page")).toHaveCSS(
    "background-color",
    "rgb(255, 255, 255)",
  );

  // Centred horizontally within the canvas (small tolerance for scrollbars).
  const leftGap = doc.x - canvas.x;
  const rightGap = canvas.x + canvas.width - (doc.x + doc.width);
  expect(Math.abs(leftGap - rightGap)).toBeLessThan(20);

  // Margins are visible: the margin box sits strictly inside the page on
  // all four sides.
  expect(margin.x).toBeGreaterThan(doc.x);
  expect(margin.y).toBeGreaterThan(doc.y);
  expect(margin.x + margin.width).toBeLessThan(doc.x + doc.width);
  expect(margin.y + margin.height).toBeLessThan(doc.y + doc.height);
});

test("clicking the title makes it editable, and Enter commits the new value (AC-6, AC-7)", async ({
  page,
}) => {
  await page.goto("/");
  await page.getByRole("button", { name: "New document" }).click();

  await page.getByTestId("document-title").click();
  const input = page.getByTestId("document-title-input");
  await expect(input).toBeVisible();
  await expect(input).toHaveValue("Untitled document");

  await input.fill("Meeting Notes");
  await input.press("Enter");

  await expect(page.getByTestId("document-title")).toHaveText(
    "Meeting Notes",
  );
});

test("blurring the title input commits the new value (AC-7)", async ({
  page,
}) => {
  await page.goto("/");
  await page.getByRole("button", { name: "New document" }).click();

  await page.getByTestId("document-title").click();
  await page.getByTestId("document-title-input").fill("Blurred Title");
  await page.getByTestId("canvas").click();

  await expect(page.getByTestId("document-title")).toHaveText(
    "Blurred Title",
  );
});

test("pressing Escape reverts the title without committing (AC-8)", async ({
  page,
}) => {
  await page.goto("/");
  await page.getByRole("button", { name: "New document" }).click();

  await page.getByTestId("document-title").click();
  await page.getByTestId("document-title-input").fill("Discarded Title");
  await page.getByTestId("document-title-input").press("Escape");

  await expect(page.getByTestId("document-title")).toHaveText(
    "Untitled document",
  );
});

test("a whitespace-only title reverts to 'Untitled document' (AC-9)", async ({
  page,
}) => {
  await page.goto("/");
  await page.getByRole("button", { name: "New document" }).click();

  await page.getByTestId("document-title").click();
  await page.getByTestId("document-title-input").fill("Renamed First");
  await page.getByTestId("document-title-input").press("Enter");
  await expect(page.getByTestId("document-title")).toHaveText(
    "Renamed First",
  );

  await page.getByTestId("document-title").click();
  await page.getByTestId("document-title-input").fill("   ");
  await page.getByTestId("document-title-input").press("Enter");

  await expect(page.getByTestId("document-title")).toHaveText(
    "Untitled document",
  );
});
