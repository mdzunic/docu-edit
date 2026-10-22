// SPEC-125 sheetGestures: the outline band a press or hover starts a
// drag-move in (MOVE_HIT_PX on screen, either side of the edge, at any
// zoom), and the modifiers that keep a press on it the selection's — Shift,
// and Cmd on a Mac — while Alt/Option and Ctrl on Windows/Linux, the move's
// copy modifiers, still start one.
import { describe, expect, it } from "vitest";
import { MOVE_HIT_PX, nearOutline, outlinePressMoves } from "./sheetGestures";
import type { RectItem } from "./useSheetOverlays";
import { COLUMN_HEADER_HEIGHT_PX, ROW_HEADER_WIDTH_PX } from "./useSheetViewport";

const press = (init: { shiftKey?: boolean; metaKey?: boolean } = {}) => ({ shiftKey: false, metaKey: false, ...init });

describe("SPEC-125 sheetGestures", () => {
  it("the outline band reaches MOVE_HIT_PX screen px either side of the edge, at any zoom, and no deeper inside", () => {
    const outline: RectItem = { key: "outline", left: 64, top: 20, width: 128, height: 40 };
    const near = (x: number, y: number, zoom = 1): boolean => nearOutline(ROW_HEADER_WIDTH_PX + x, COLUMN_HEADER_HEIGHT_PX + y, outline, zoom);
    expect(near(64, 40)).toBe(true);
    expect(near(64 - MOVE_HIT_PX, 40)).toBe(true);
    expect(near(64 - MOVE_HIT_PX - 1, 40)).toBe(false);
    expect(near(64 + MOVE_HIT_PX + 1, 40)).toBe(false);
    expect(near(128, 60 + MOVE_HIT_PX)).toBe(true);
    // At 200%, 4 screen px are 2 paper px.
    expect(near(64 - 2, 40, 2)).toBe(true);
    expect(near(64 - 3, 40, 2)).toBe(false);
    expect(nearOutline(0, 0, null, 1)).toBe(false);
  });

  it("a press with Shift, or Cmd on a Mac, stays the selection's; a plain one, Alt/Option or Ctrl moves", () => {
    for (const mac of [false, true]) {
      expect(outlinePressMoves(press(), mac)).toBe(true);
      expect(outlinePressMoves(press({ shiftKey: true }), mac)).toBe(false);
    }
    expect(outlinePressMoves(press({ metaKey: true }), true)).toBe(false);
    expect(outlinePressMoves(press({ metaKey: true }), false)).toBe(true);
  });
});
