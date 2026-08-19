import { test, expect, type Page } from "@playwright/test";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";

// SPEC-057 — native drawing-group containers. The SPEC-051 fixtures carry a
// legacy `de:group` pair that migrates on open to one native DrawingGroup
// container (AC-65), which renders as a selectable, enterable group (AC-77).
// The rest is exercised through the create/ungroup gestures on a fresh
// document, so it needs no new binary fixtures.
//
// Fixture drawing-index map after migration (see e2e/SPEC-051.spec.ts):
//   0..13 catalogue shapes, 14 migrated group container, 15 text box,
//   16 unsupported geometry.

const fixturesDir = join(dirname(fileURLToPath(import.meta.url)), "fixtures");
const docxShapesFixture = join(fixturesDir, "word-shapes.docx");
const odtShapesFixture = join(fixturesDir, "lo-shapes.odt");
const docxGroupsFixture = join(fixturesDir, "word-groups.docx");
const odtGroupsFixture = join(fixturesDir, "lo-groups.odt");

const content = (page: Page) => page.getByTestId("document-content");
const drawings = (page: Page) => content(page).locator('[data-testid^="drawing-"]:not([data-testid*="-member-"])');
const drawingAt = (page: Page, index: number) => content(page).locator(`[data-testid="drawing-${index}"]`);

async function openFixture(page: Page, path: string): Promise<void> {
  await page.goto("/");
  await page.getByTestId("open-document-input").setInputFiles(path);
  await expect(page).toHaveURL(/\/doc\/.+/);
}

async function createDocument(page: Page): Promise<void> {
  await page.goto("/");
  await page.getByTestId("new-document").click();
  await expect(page).toHaveURL(/\/doc\/.+/);
}

async function insertShapeAt(page: Page, geometry: string, x: number, y: number): Promise<void> {
  await content(page).click({ position: { x: 20, y: 20 } });
  await page.getByTestId("menu-title-insert").click();
  await page.getByTestId(`shape-gallery-${geometry}`).click();
  await content(page).click({ position: { x, y } });
}

for (const [label, fixture] of [
  ["word-shapes.docx", docxShapesFixture],
  ["lo-shapes.odt", odtShapesFixture],
] as const) {
  test.describe(`SPEC-057 — migrated ${label} opens as a native container`, () => {
    test("the legacy de:group pair renders as one container, selectable and enterable (AC-65, AC-77)", async ({ page }) => {
      await openFixture(page, fixture);
      const container = drawingAt(page, 14);
      await expect(container).toHaveAttribute("role", "group");
      const memberA = page.locator('[data-testid="drawing-14-member-0"]');
      await expect(memberA).toHaveCount(1);
      await expect(page.locator('[data-testid="drawing-14-member-1"]')).toHaveCount(1);

      // First click selects the outermost container (AC-22): its toolbar shows.
      await memberA.click();
      await expect(page.getByTestId("group-toolbar")).toBeVisible();
      await expect(page.getByTestId("group-handles")).toBeVisible();

      // A second click descends to the member under the pointer (AC-23) — the
      // member proxy becomes focusable (tabindex 0) once its container is
      // entered.
      await memberA.click();
      await expect(memberA).toHaveAttribute("tabindex", "0");

      // Escape steps back out one level (AC-24): the container is selected
      // again, its member no longer focusable.
      await page.keyboard.press("Escape");
      await expect(page.getByTestId("group-toolbar")).toBeVisible();
      await expect(memberA).toHaveAttribute("tabindex", "-1");
    });
  });
}

test.describe("SPEC-057 — Group and Ungroup gestures", () => {
  test("Group replaces two shapes with a container; Ctrl+Shift+G ungroups them back (AC-3, AC-5)", async ({ page }) => {
    await createDocument(page);
    await insertShapeAt(page, "rect", 400, 250);
    await insertShapeAt(page, "ellipse", 600, 250);
    await expect(drawings(page)).toHaveCount(2);

    const [first, second] = await drawings(page).all();
    const firstBox = await first.boundingBox();
    const secondBox = await second.boundingBox();
    if (!firstBox || !secondBox) {
      throw new Error("a drawing has no bounding box");
    }
    const [leftShape, rightShape] = firstBox.x <= secondBox.x ? [first, second] : [second, first];
    const [leftBox, rightBox] = firstBox.x <= secondBox.x ? [firstBox, secondBox] : [secondBox, firstBox];
    await leftShape.click({ position: { x: leftBox.width / 2, y: leftBox.height * 0.72 } });
    await expect(page.getByTestId("shape-toolbar")).toBeVisible();
    await rightShape.click({ modifiers: ["Shift"], position: { x: rightBox.width / 2, y: rightBox.height * 0.72 } });

    // Ctrl/Cmd+G groups the multi-selection into one container.
    await page.keyboard.press("ControlOrMeta+g");
    await expect(drawings(page)).toHaveCount(1);
    await expect(drawingAt(page, 0)).toHaveAttribute("role", "group");
    await expect(page.getByTestId("group-toolbar")).toBeVisible();

    // Ctrl/Cmd+Shift+G ungroups one level: the two shapes re-emerge.
    await page.keyboard.press("ControlOrMeta+Shift+g");
    await expect(drawings(page)).toHaveCount(2);
  });

  test("grouping a selection that includes a group nests it as one member; Ungroup peels one level (AC-7, AC-8)", async ({ page }) => {
    await createDocument(page);
    await insertShapeAt(page, "rect", 300, 250);
    await insertShapeAt(page, "ellipse", 480, 250);
    await insertShapeAt(page, "rect", 680, 250);
    await expect(drawings(page)).toHaveCount(3);

    // Order the three by x so the gesture is deterministic.
    const all = await drawings(page).all();
    const boxed = [];
    for (const d of all) {
      const box = await d.boundingBox();
      if (!box) {
        throw new Error("a drawing has no bounding box");
      }
      boxed.push({ d, box });
    }
    boxed.sort((a, b) => a.box.x - b.box.x);

    // Group the two leftmost into a container (Shift-click multi-select + Ctrl+G).
    await boxed[0].d.click({ position: { x: boxed[0].box.width / 2, y: boxed[0].box.height / 2 } });
    await expect(page.getByTestId("shape-toolbar")).toBeVisible();
    await boxed[1].d.click({ modifiers: ["Shift"], position: { x: boxed[1].box.width / 2, y: boxed[1].box.height * 0.72 } });
    await page.keyboard.press("ControlOrMeta+g");
    await expect(drawings(page)).toHaveCount(2);

    // One container (role=group) and one loose shape remain; Group leaves the
    // new container selected (its toolbar shows).
    const containerSel = '[data-testid^="drawing-"][role="group"]:not([data-testid*="-member-"])';
    const looseSel = '[data-testid^="drawing-"]:not([data-testid*="-member-"]):not([role="group"])';
    await expect(content(page).locator(containerSel)).toHaveCount(1);
    await expect(page.getByTestId("group-toolbar")).toBeVisible();

    // Shift-click the loose shape — the selected container and the shape now
    // form one cross-kind Group selection (AC-7, the fix this round).
    const loose = content(page).locator(looseSel);
    const looseBox = await loose.boundingBox();
    if (!looseBox) {
      throw new Error("loose shape has no bounding box");
    }
    await loose.click({ modifiers: ["Shift"], position: { x: looseBox.width / 2, y: looseBox.height / 2 } });

    // Ctrl+G nests the group as one member of a new outer container.
    await page.keyboard.press("ControlOrMeta+g");
    await expect(drawings(page)).toHaveCount(1);
    await expect(drawingAt(page, 0)).toHaveAttribute("role", "group");
    // One of the outer container's members is itself a group (nesting).
    await expect(page.locator('[data-testid^="drawing-0-member-"][role="group"]')).toHaveCount(1);

    // Group leaves the new outer container selected; Ungroup peels exactly one
    // level, so the inner container survives as a top-level sibling.
    await expect(page.getByTestId("group-toolbar")).toBeVisible();
    await page.keyboard.press("ControlOrMeta+Shift+g");
    await expect(drawings(page)).toHaveCount(2);
    await expect(content(page).locator(containerSel)).toHaveCount(1);
  });

  test("Group moves no member on screen (AC-4)", async ({ page }) => {
    await createDocument(page);
    await insertShapeAt(page, "rect", 400, 250);
    await insertShapeAt(page, "ellipse", 600, 250);
    await expect(drawings(page)).toHaveCount(2);

    const [first, second] = await drawings(page).all();
    const firstBox = await first.boundingBox();
    const secondBox = await second.boundingBox();
    if (!firstBox || !secondBox) {
      throw new Error("a drawing has no bounding box");
    }
    const [leftShape, rightShape] = firstBox.x <= secondBox.x ? [first, second] : [second, first];
    const [leftBox, rightBox] = firstBox.x <= secondBox.x ? [firstBox, secondBox] : [secondBox, firstBox];
    await leftShape.click({ position: { x: leftBox.width / 2, y: leftBox.height * 0.72 } });
    await expect(page.getByTestId("shape-toolbar")).toBeVisible();
    await rightShape.click({ modifiers: ["Shift"], position: { x: rightBox.width / 2, y: rightBox.height * 0.72 } });

    await page.keyboard.press("ControlOrMeta+g");
    await expect(drawingAt(page, 0)).toHaveAttribute("role", "group");

    // The two members must land exactly where they were before Group — the
    // container's offset is anchor-relative, so grouping is a pure re-parenting
    // with no visible translation.
    const member0 = await page.locator('[data-testid="drawing-0-member-0"]').boundingBox();
    const member1 = await page.locator('[data-testid="drawing-0-member-1"]').boundingBox();
    if (!member0 || !member1) {
      throw new Error("a member has no bounding box");
    }
    const memberBoxes = [member0, member1];
    for (const original of [leftBox, rightBox]) {
      const match = memberBoxes.find((m) => Math.abs(m.x - original.x) < 2 && Math.abs(m.y - original.y) < 2);
      expect(match, `a member should stay at (${original.x}, ${original.y})`).toBeTruthy();
    }
  });

  test("a second click selects a member, whose toolbar offers the reorder verbs; Delete dissolves a two-member container (AC-23, AC-31/32, AC-33..36)", async ({ page }) => {
    await openFixture(page, docxShapesFixture);
    const member0 = page.locator('[data-testid="drawing-14-member-0"]');
    // First click selects the container, second descends to the member.
    await member0.click();
    await member0.click();
    await expect(page.getByTestId("group-member-reorder")).toBeVisible();
    // Send it to the back — reorders it among its siblings (member order is
    // paint order).
    await page.getByTestId("group-member-reorder").selectOption("back");
    // Deleting a member of a two-member container dissolves it into the
    // remaining member: the container proxy at index 14 is gone, replaced by a
    // plain shape.
    await page.locator('[data-testid="drawing-14-member-0"]').click();
    await page.locator('[data-testid="drawing-14-member-0"]').click();
    await page.keyboard.press("Delete");
    await expect(drawingAt(page, 14)).toHaveAttribute("role", "img");
  });

  test("double-clicking a member enters its text, editable in place through SPEC-051's lifecycle (AC-25)", async ({ page }) => {
    await createDocument(page);
    await insertShapeAt(page, "rect", 400, 250);
    await insertShapeAt(page, "ellipse", 600, 250);
    await expect(drawings(page)).toHaveCount(2);

    const [first, second] = await drawings(page).all();
    const firstBox = await first.boundingBox();
    const secondBox = await second.boundingBox();
    if (!firstBox || !secondBox) {
      throw new Error("a drawing has no bounding box");
    }
    const [leftShape, rightShape] = firstBox.x <= secondBox.x ? [first, second] : [second, first];
    const [leftBox, rightBox] = firstBox.x <= secondBox.x ? [firstBox, secondBox] : [secondBox, firstBox];
    await leftShape.click({ position: { x: leftBox.width / 2, y: leftBox.height * 0.72 } });
    await expect(page.getByTestId("shape-toolbar")).toBeVisible();
    await rightShape.click({ modifiers: ["Shift"], position: { x: rightBox.width / 2, y: rightBox.height * 0.72 } });
    await page.keyboard.press("ControlOrMeta+g");
    await expect(drawingAt(page, 0)).toHaveAttribute("role", "group");

    // Select a member (first click selects the container, second descends),
    // then double-click to enter its text — the SPEC-051 gesture.
    const member0 = page.locator('[data-testid="drawing-0-member-0"]');
    await member0.click();
    await member0.click();
    await expect(member0).toHaveAttribute("tabindex", "0");
    await member0.dblclick();

    const text = page.getByTestId("member-shape-text");
    await expect(text).toHaveAttribute("contenteditable", "true");
    await page.keyboard.type("In group");
    await expect(text).toContainText("In group");

    // Clicking away re-renders the member read-only from the model — the text
    // surviving proves it landed in the member's Shape.text, not just the
    // transient contenteditable DOM.
    await content(page).click({ position: { x: 60, y: 560 } });
    await expect(page.getByTestId("member-shape-text")).toContainText("In group");
  });

  test("dragging inside a member's live text frame selects text instead of moving the member (AC-25)", async ({ page }) => {
    await createDocument(page);
    await insertShapeAt(page, "rect", 400, 250);
    await insertShapeAt(page, "ellipse", 600, 250);
    await expect(drawings(page)).toHaveCount(2);

    const [first, second] = await drawings(page).all();
    const firstBox = await first.boundingBox();
    const secondBox = await second.boundingBox();
    if (!firstBox || !secondBox) {
      throw new Error("a drawing has no bounding box");
    }
    const [leftShape, rightShape] = firstBox.x <= secondBox.x ? [first, second] : [second, first];
    const [leftBox, rightBox] = firstBox.x <= secondBox.x ? [firstBox, secondBox] : [secondBox, firstBox];
    await leftShape.click({ position: { x: leftBox.width / 2, y: leftBox.height * 0.72 } });
    await expect(page.getByTestId("shape-toolbar")).toBeVisible();
    await rightShape.click({ modifiers: ["Shift"], position: { x: rightBox.width / 2, y: rightBox.height * 0.72 } });
    await page.keyboard.press("ControlOrMeta+g");
    await expect(drawingAt(page, 0)).toHaveAttribute("role", "group");

    const member0 = page.locator('[data-testid="drawing-0-member-0"]');
    await member0.click();
    await member0.click();
    await member0.dblclick();
    const text = page.getByTestId("member-shape-text");
    await expect(text).toHaveAttribute("contenteditable", "true");
    await page.keyboard.type("In group");

    const before = await member0.boundingBox();
    const lineBox = await text.locator(".doc-paragraph").first().boundingBox();
    if (!before || !lineBox) {
      throw new Error("the member or its text line has no bounding box");
    }
    // Drag across the rendered text line inside the live frame — the SPEC-051
    // select-text gesture. The mousedown bubbles to the canvas inside
    // .doc-group-anchor while the member is selected; it must fall through to
    // native text selection, not arm a member move. Drag on the line itself
    // (not the frame's empty lower half), inset from the edges so the member's
    // resize handles — live during editing at top level too — can't catch it.
    const lineY = lineBox.y + lineBox.height / 2;
    await page.mouse.move(lineBox.x + lineBox.width * 0.15, lineY);
    await page.mouse.down();
    await page.mouse.move(lineBox.x + lineBox.width * 0.85, lineY, { steps: 8 });
    await page.mouse.up();

    // The drag selected a real run of the member's text (exactly which
    // characters depends on glyph metrics, so assert containment instead).
    const selected = await page.evaluate(() => document.getSelection()?.toString() ?? "");
    expect(selected.length).toBeGreaterThanOrEqual(2);
    expect("In group").toContain(selected);
    // …the frame is still live…
    await expect(text).toHaveAttribute("contenteditable", "true");
    // …and the member stayed put (no moveMember commit on release).
    const after = await member0.boundingBox();
    if (!after) {
      throw new Error("the member lost its bounding box after the drag");
    }
    expect(Math.abs(after.x - before.x)).toBeLessThan(1);
    expect(Math.abs(after.y - before.y)).toBeLessThan(1);
  });

  test("the container carries editable alt text from its toolbar (AC-49)", async ({ page }) => {
    await createDocument(page);
    await insertShapeAt(page, "rect", 400, 250);
    await insertShapeAt(page, "ellipse", 600, 250);
    const [first, second] = await drawings(page).all();
    const firstBox = await first.boundingBox();
    const secondBox = await second.boundingBox();
    if (!firstBox || !secondBox) {
      throw new Error("a drawing has no bounding box");
    }
    const [leftShape, rightShape] = firstBox.x <= secondBox.x ? [first, second] : [second, first];
    await leftShape.click({ position: { x: firstBox.width / 2, y: firstBox.height * 0.72 } });
    await expect(page.getByTestId("shape-toolbar")).toBeVisible();
    await rightShape.click({ modifiers: ["Shift"], position: { x: secondBox.width / 2, y: secondBox.height * 0.72 } });
    await page.keyboard.press("ControlOrMeta+g");

    const altInput = page.getByTestId("group-alt-text");
    await expect(altInput).toBeVisible();
    await altInput.fill("My diagram");
    // The container proxy reflects the alt text as its aria-label.
    await expect(drawingAt(page, 0)).toHaveAttribute("aria-label", "My diagram");
  });
});

// SPEC-057 — the hand-built native-group fixtures (build-spec057-fixtures.mjs).
// Each holds five containers in document order (drawing-0..4): a plain pair, a
// nested group, a mixed shape/image/text-box group, a rotated container and a
// preserved-member (sealed) group. AC-77..80, AC-82.
for (const [label, fixture] of [
  ["word-groups.docx", docxGroupsFixture],
  ["lo-groups.odt", odtGroupsFixture],
] as const) {
  test.describe(`SPEC-057 — native ${label} renders its groups`, () => {
    test("all five containers render as selectable, enterable groups; the sealed one does not enter (AC-77..80, AC-82)", async ({ page }) => {
      await openFixture(page, fixture);
      await expect(drawings(page)).toHaveCount(5);

      // The four unsealed groups: a first click selects the container (its
      // toolbar and resize handles show), a second descends into a member
      // (its proxy becomes focusable), Escape steps back out.
      for (const index of [0, 1, 2, 3]) {
        const container = drawingAt(page, index);
        await expect(container).toHaveAttribute("role", "group");
        const member = page.locator(`[data-testid="drawing-${index}-member-0"]`);
        await expect(member).toHaveCount(1);

        await member.click();
        await expect(page.getByTestId("group-toolbar")).toBeVisible();
        await expect(page.getByTestId("group-handles")).toBeVisible();

        await member.click();
        await expect(member).toHaveAttribute("tabindex", "0");

        await page.keyboard.press("Escape");
        await expect(member).toHaveAttribute("tabindex", "-1");
        await page.keyboard.press("Escape");
      }

      // The mixed group (index 2) nests an image member proxy.
      await expect(page.locator('[data-testid="drawing-2-member-1"]')).toHaveAttribute("role", "img");

      // The preserved-member group (index 4) is sealed: it selects but offers
      // no resize/rotation handles and does not enter — a second click keeps
      // the container selected, its members never become focusable.
      const sealed = drawingAt(page, 4);
      await expect(sealed).toHaveAttribute("role", "group");
      const sealedMember = page.locator('[data-testid="drawing-4-member-0"]');
      await sealedMember.click();
      await expect(page.getByTestId("group-toolbar")).toBeVisible();
      await expect(page.getByTestId("group-handles")).toHaveCount(0);
      await sealedMember.click();
      await expect(sealedMember).toHaveAttribute("tabindex", "-1");
    });
  });
}
