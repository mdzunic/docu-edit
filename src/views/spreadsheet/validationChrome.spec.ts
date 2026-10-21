// SPEC-130's chrome placement: the input message below-right of the cell,
// flipping left and up to stay inside the cell area; the dropdown button
// against the cell's right edge (inside it at the area's edge); the list
// under the cell (above when there is no room), eight rows tall at most and
// always inside the area — the roomier side, scrolling, when neither side
// holds it; shifted left at the right edge.
// The list's keys are covered through ValidationDropdown.spec.ts.
import { describe, expect, it } from "vitest";
import { DROPDOWN_BUTTON_PX, INPUT_MESSAGE_GAP_PX, LIST_ROW_PX, nextItemStartingWith, placeDropdownButton, placeInputMessage, placeValidationList } from "./validationChrome";

const BOUNDS = { width: 600, height: 400 };
const CELL = { left: 100, top: 40, width: 64, height: 20 };

describe("placeInputMessage", () => {
  it("sits below-right of the cell", () => {
    expect(placeInputMessage(CELL, { width: 120, height: 40 }, BOUNDS)).toEqual({ left: 164, top: 60 + INPUT_MESSAGE_GAP_PX });
  });

  it("flips to the cell's left past the right edge and above it past the bottom", () => {
    const corner = { left: 520, top: 370, width: 64, height: 20 };
    expect(placeInputMessage(corner, { width: 120, height: 40 }, BOUNDS)).toEqual({ left: 400, top: 370 - 40 - INPUT_MESSAGE_GAP_PX });
  });

  it("never leaves the area's top-left", () => {
    expect(placeInputMessage({ left: 10, top: 10, width: 20, height: 20 }, { width: 900, height: 900 }, BOUNDS)).toEqual({ left: 0, top: 0 });
  });

  it("flips left at the right edge and stays below when there is room there (QA's right-edge cell)", () => {
    const edge = { left: 520, top: 40, width: 64, height: 20 };
    expect(placeInputMessage(edge, { width: 221, height: 40 }, BOUNDS)).toEqual({ left: 520 - 221, top: 60 + INPUT_MESSAGE_GAP_PX });
  });

  it("fitting on neither side, stays inside the area as near below-right as it can", () => {
    // Too wide for either side of the cell: inside the right edge.
    expect(placeInputMessage({ left: 250, top: 40, width: 64, height: 20 }, { width: 280, height: 40 }, { width: 300, height: 400 }).left).toBe(20);
    // Too tall for either side: as low as stays inside, not pinned at the top.
    expect(placeInputMessage({ left: 100, top: 100, width: 64, height: 20 }, { width: 120, height: 150 }, { width: 600, height: 200 }).top).toBe(50);
  });
});

describe("placeDropdownButton", () => {
  it("sits just outside the cell's right edge, the cell's height tall", () => {
    expect(placeDropdownButton(CELL, BOUNDS)).toEqual({ left: 164, top: 40, width: DROPDOWN_BUTTON_PX, height: 20 });
  });

  it("moves inside the cell where the area ends", () => {
    expect(placeDropdownButton({ left: 536, top: 0, width: 64, height: 20 }, BOUNDS).left).toBe(600 - DROPDOWN_BUTTON_PX);
  });
});

describe("placeValidationList", () => {
  it("opens under the cell, at least its width, at most eight rows tall", () => {
    expect(placeValidationList(CELL, 3, BOUNDS)).toEqual({ left: 100, top: 60, minWidth: 64 + DROPDOWN_BUTTON_PX, maxWidth: 500, maxHeight: 3 * LIST_ROW_PX + 2 });
    expect(placeValidationList(CELL, 50, BOUNDS).maxHeight).toBe(8 * LIST_ROW_PX + 2);
  });

  it("opens above the cell when the room is there and not below", () => {
    expect(placeValidationList({ ...CELL, top: 380 }, 8, BOUNDS).top).toBe(380 - (8 * LIST_ROW_PX + 2));
    // A cell the area's bottom edge cuts through (QA's B10 in a short window).
    expect(placeValidationList({ ...CELL, top: 190 }, 3, { width: 600, height: 200 })).toMatchObject({ top: 190 - (3 * LIST_ROW_PX + 2), maxHeight: 3 * LIST_ROW_PX + 2 });
  });

  it("with room on neither side, takes the roomier one and scrolls in what it holds", () => {
    const short = { width: 600, height: 200 };
    // 100 px above, 80 below: above, from the area's top.
    expect(placeValidationList({ ...CELL, top: 100 }, 8, short)).toMatchObject({ top: 0, maxHeight: 100 });
    // 60 px above, 120 below: below.
    expect(placeValidationList({ ...CELL, top: 60 }, 8, short)).toMatchObject({ top: 80, maxHeight: 120 });
    // Neither side holds a row: over the cell, inside the area.
    expect(placeValidationList({ left: 100, top: 5, width: 64, height: 20 }, 8, { width: 600, height: 30 })).toMatchObject({ top: 0, maxHeight: 30 });
  });

  it("shifts left to keep its width inside the right edge, a long item cut there", () => {
    expect(placeValidationList({ ...CELL, left: 560 }, 3, BOUNDS)).toMatchObject({ left: 600 - (64 + DROPDOWN_BUTTON_PX), maxWidth: 64 + DROPDOWN_BUTTON_PX });
    expect(placeValidationList({ ...CELL, left: 0, width: 700 }, 3, BOUNDS)).toMatchObject({ left: 0, minWidth: 700 + DROPDOWN_BUTTON_PX });
  });
});

describe("nextItemStartingWith", () => {
  it("finds the next match after `from`, wrapping, case-insensitively", () => {
    const texts = ["apple", "Banana", "avocado", "blueberry"];
    expect(nextItemStartingWith(texts, -1, "B")).toBe(1);
    expect(nextItemStartingWith(texts, 1, "b")).toBe(3);
    expect(nextItemStartingWith(texts, 3, "b")).toBe(1);
    expect(nextItemStartingWith(texts, 0, "z")).toBe(-1);
    expect(nextItemStartingWith([], -1, "a")).toBe(-1);
  });
});
