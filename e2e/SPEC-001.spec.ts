import { test, expect } from "@playwright/test";

// SPEC-001 — app shell and document workspace.
//
// The whole file runs at a 1024px-wide viewport: AC-16/AC-17 require the
// layout to work at exactly that width, and every other criterion must hold
// there too.
//
// Not covered here: the "Document not found" state. Per the spec's Decisions,
// every routable id is treated as valid until persistence lands, so no URL can
// reach that state; it is covered by unit tests (src/views/DocumentView.spec.ts).
test.use({ viewport: { width: 1024, height: 768 } });

test("/ shows the app shell: app bar, toolbar strip, canvas (AC-1)", async ({
  page,
}) => {
  await page.goto("/");
  await expect(page.getByTestId("app-bar")).toBeVisible();
  await expect(page.getByTestId("toolbar")).toBeVisible();
  await expect(page.getByTestId("canvas")).toBeVisible();
});

test("app bar shows the product name and 'No document' (AC-2, AC-3, AC-4)", async ({
  page,
}) => {
  await page.goto("/");
  await expect(page.getByTestId("product-name")).toHaveText("docu-edit");
  await expect(page.getByTestId("document-title")).toHaveText("No document");
});

test("empty state shows 'No document open' and both buttons (AC-5..AC-8)", async ({
  page,
}) => {
  await page.goto("/");
  await expect(page.getByTestId("canvas")).toContainText("No document open");
  await expect(
    page.getByRole("button", { name: "New document" }),
  ).toBeVisible();
  await expect(
    page.getByRole("button", { name: "Open document" }),
  ).toBeVisible();
});

test("/doc/:id titles the app bar with the client-side timestamp", async ({
  page,
}) => {
  await page.goto("/doc/any-id");
  await expect(page.getByTestId("document-title")).toHaveText(
    /^\d{4}-\d{2}-\d{2} \d{2}:\d{2}$/,
  );
});

test("toolbar scrolls horizontally on overflow — no wrap, no clip (AC-11, AC-12, AC-13)", async ({
  page,
}) => {
  await page.goto("/");
  const toolbar = page.getByTestId("toolbar");
  const heightBefore = (await toolbar.boundingBox())!.height;

  // The strip ships empty in SPEC-001, so overflow it the way later specs
  // will: with a row of controls wider than the viewport.
  await toolbar.evaluate((el) => {
    for (let i = 0; i < 40; i += 1) {
      const button = document.createElement("button");
      button.textContent = `Control ${i}`;
      el.append(button);
    }
  });

  const metrics = await toolbar.evaluate((el) => ({
    scrollWidth: el.scrollWidth,
    clientWidth: el.clientWidth,
  }));
  expect(metrics.scrollWidth).toBeGreaterThan(metrics.clientWidth);

  // No wrap: the strip stays a single row.
  const heightAfter = (await toolbar.boundingBox())!.height;
  expect(heightAfter).toBe(heightBefore);

  // No clip: the overflowing content is reachable by scrolling.
  const scrolled = await toolbar.evaluate((el) => {
    el.scrollLeft = 10_000;
    return el.scrollLeft;
  });
  expect(scrolled).toBeGreaterThan(0);
  expect(scrolled + metrics.clientWidth).toBeGreaterThanOrEqual(
    metrics.scrollWidth - 1,
  );
});

test("layout fills the viewport and only the canvas scrolls (AC-14, AC-15)", async ({
  page,
}) => {
  await page.goto("/");

  const pageScroll = () =>
    page.evaluate(() => ({
      scrollHeight: document.documentElement.scrollHeight,
      innerHeight: window.innerHeight,
    }));

  // No page-level vertical scrollbar.
  const before = await pageScroll();
  expect(before.scrollHeight).toBeLessThanOrEqual(before.innerHeight);

  // Overflow the canvas: the page still must not scroll, the canvas must.
  const canvas = page.getByTestId("canvas");
  await canvas.evaluate((el) => {
    const filler = document.createElement("div");
    filler.style.height = "3000px";
    el.append(filler);
  });

  const after = await pageScroll();
  expect(after.scrollHeight).toBeLessThanOrEqual(after.innerHeight);

  const canvasScrolled = await canvas.evaluate((el) => {
    el.scrollTop = 500;
    return el.scrollTop;
  });
  expect(canvasScrolled).toBeGreaterThan(0);
});

test("at 1024px the bars stack without overlap and nothing is cut off (AC-16, AC-17)", async ({
  page,
}) => {
  await page.goto("/");
  const appBar = (await page.getByTestId("app-bar").boundingBox())!;
  const toolbar = (await page.getByTestId("toolbar").boundingBox())!;
  const canvas = (await page.getByTestId("canvas").boundingBox())!;

  // Stacked vertically with no overlap (0.5px tolerance for subpixel layout).
  expect(toolbar.y).toBeGreaterThanOrEqual(appBar.y + appBar.height - 0.5);
  expect(canvas.y).toBeGreaterThanOrEqual(toolbar.y + toolbar.height - 0.5);

  // Every region fits the 1024x768 viewport — nothing cut off.
  for (const box of [appBar, toolbar, canvas]) {
    expect(box.x).toBeGreaterThanOrEqual(0);
    expect(box.y).toBeGreaterThanOrEqual(0);
    expect(box.x + box.width).toBeLessThanOrEqual(1024.5);
    expect(box.y + box.height).toBeLessThanOrEqual(768.5);
  }

  // The empty-state actions are fully inside the canvas, hence usable.
  const openButton = (await page
    .getByRole("button", { name: "Open document" })
    .boundingBox())!;
  expect(openButton.y).toBeGreaterThanOrEqual(canvas.y);
  expect(openButton.y + openButton.height).toBeLessThanOrEqual(
    canvas.y + canvas.height + 0.5,
  );
});
