import { describe, it, expect, beforeEach } from "vitest";
import {
  HIGHLIGHT_PICKER_ITEMS,
  loadRecentColors,
  parseColorInput,
  pushRecentColor,
  RECENT_COLORS_LIMIT,
  TEXT_COLOR_PALETTE,
} from "./colors";
import { HIGHLIGHT_NAMES } from "./colors";

describe("TEXT_COLOR_PALETTE", () => {
  it("is a standard palette of lowercase #rrggbb colours including black and primaries", () => {
    for (const hex of TEXT_COLOR_PALETTE) {
      expect(hex).toMatch(/^#[0-9a-f]{6}$/);
    }
    expect(TEXT_COLOR_PALETTE).toContain("#000000");
    expect(TEXT_COLOR_PALETTE).toContain("#ff0000");
    expect(new Set(TEXT_COLOR_PALETTE).size).toBe(TEXT_COLOR_PALETTE.length);
  });
});

describe("parseColorInput", () => {
  it("accepts hex with and without the #", () => {
    expect(parseColorInput("#FF8800")).toBe("#ff8800");
    expect(parseColorInput("ff8800")).toBe("#ff8800");
    expect(parseColorInput("  #ff8800  ")).toBe("#ff8800");
  });

  it("expands 3-digit hex", () => {
    expect(parseColorInput("#f80")).toBe("#ff8800");
  });

  it("accepts rgb() and normalises to hex", () => {
    expect(parseColorInput("rgb(255, 136, 0)")).toBe("#ff8800");
    expect(parseColorInput("rgb(0,0,0)")).toBe("#000000");
  });

  it("rejects out-of-range channels and garbage", () => {
    expect(parseColorInput("rgb(300, 0, 0)")).toBeUndefined();
    expect(parseColorInput("#ff88")).toBeUndefined();
    expect(parseColorInput("not a colour")).toBeUndefined();
    expect(parseColorInput("")).toBeUndefined();
  });
});

describe("recent colours", () => {
  beforeEach(() => {
    localStorage.clear();
  });

  it("starts empty and persists pushes most-recent-first", () => {
    expect(loadRecentColors()).toEqual([]);
    let recent = pushRecentColor("#ff0000", []);
    recent = pushRecentColor("#00ff00", recent);
    expect(recent).toEqual(["#00ff00", "#ff0000"]);
    expect(loadRecentColors()).toEqual(["#00ff00", "#ff0000"]);
  });

  it("de-duplicates by moving a re-applied colour to the front", () => {
    let recent = pushRecentColor("#ff0000", []);
    recent = pushRecentColor("#00ff00", recent);
    recent = pushRecentColor("#ff0000", recent);
    expect(recent).toEqual(["#ff0000", "#00ff00"]);
  });

  it("caps the list at the limit", () => {
    let recent: string[] = [];
    for (let i = 0; i < RECENT_COLORS_LIMIT + 3; i++) {
      recent = pushRecentColor(`#0000${i.toString(16).padStart(2, "0")}`, recent);
    }
    expect(recent).toHaveLength(RECENT_COLORS_LIMIT);
    expect(loadRecentColors()).toHaveLength(RECENT_COLORS_LIMIT);
  });

  it("ignores corrupt stored data", () => {
    localStorage.setItem("docu-edit:recent-text-colors", "{not json");
    expect(loadRecentColors()).toEqual([]);
    localStorage.setItem("docu-edit:recent-text-colors", JSON.stringify(["red", 5, "#ff0000"]));
    expect(loadRecentColors()).toEqual(["#ff0000"]);
  });
});

describe("HIGHLIGHT_PICKER_ITEMS", () => {
  it("offers exactly the 15 Word highlight colours (white stays out) (AC)", () => {
    expect(HIGHLIGHT_PICKER_ITEMS).toHaveLength(15);
    expect(HIGHLIGHT_PICKER_ITEMS.some((item) => item.name === "white")).toBe(false);
    for (const item of HIGHLIGHT_PICKER_ITEMS) {
      // Every picker swatch maps back through the shared 16-entry palette.
      expect(HIGHLIGHT_NAMES[item.hex]).toBe(item.name);
      expect(item.label.length).toBeGreaterThan(0);
    }
  });

  it("uses Word's UI names for the divergent entries", () => {
    const byName = new Map(HIGHLIGHT_PICKER_ITEMS.map((item) => [item.name, item.label]));
    expect(byName.get("green")).toBe("Bright Green");
    expect(byName.get("cyan")).toBe("Turquoise");
    expect(byName.get("magenta")).toBe("Pink");
    expect(byName.get("darkCyan")).toBe("Teal");
  });
});
