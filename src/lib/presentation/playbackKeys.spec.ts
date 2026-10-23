// SPEC-084 — the normative playback key map. These assertions pin the exact
// bindings the play view and (SPEC-088) the HTML export both depend on, and the
// classifier's case-insensitive letter handling. If a downstream spec needs to
// change the map, it changes here first.
import { describe, expect, it } from "vitest";
import {
  advanceKeys,
  backKeys,
  blackoutKeys,
  firstKeys,
  lastKeys,
  normaliseKey,
  playbackActionForKey,
  whiteoutKeys,
} from "./playbackKeys";

describe("playbackKeys map", () => {
  it("advance keys cover click-equivalents, arrows, space, page-down, enter, N", () => {
    for (const key of ["ArrowRight", "ArrowDown", " ", "Spacebar", "PageDown", "Enter", "n"]) {
      expect(advanceKeys).toContain(key);
      expect(playbackActionForKey(key)).toBe("advance");
    }
  });

  it("back keys cover arrows, page-up, P, backspace", () => {
    for (const key of ["ArrowLeft", "ArrowUp", "PageUp", "p", "Backspace"]) {
      expect(backKeys).toContain(key);
      expect(playbackActionForKey(key)).toBe("back");
    }
  });

  it("Home/End jump to the first/last slide", () => {
    expect(firstKeys).toContain("Home");
    expect(lastKeys).toContain("End");
    expect(playbackActionForKey("Home")).toBe("first");
    expect(playbackActionForKey("End")).toBe("last");
  });

  it("letter bindings are case-insensitive (Shift-held N/P still navigate)", () => {
    expect(normaliseKey("N")).toBe("n");
    expect(playbackActionForKey("N")).toBe("advance");
    expect(playbackActionForKey("P")).toBe("back");
  });

  it("multi-character keys pass through normaliseKey unchanged", () => {
    expect(normaliseKey("ArrowRight")).toBe("ArrowRight");
  });

  it("returns null for keys it does not own", () => {
    expect(playbackActionForKey("a")).toBeNull();
    expect(playbackActionForKey("Escape")).toBeNull();
    expect(playbackActionForKey("Tab")).toBeNull();
  });

  it("PageDown/PageUp are the editor's two shared keys (must not drift)", () => {
    // usePresentationEditing imports these very entries for its slide nav.
    expect(advanceKeys).toContain("PageDown");
    expect(backKeys).toContain("PageUp");
  });

  // SPEC-100: the blank-screen rows SPEC-084 deferred.
  it("b and . black the screen out; w and , white it out (case-insensitive)", () => {
    expect(blackoutKeys).toEqual(["b", "."]);
    expect(whiteoutKeys).toEqual(["w", ","]);
    for (const key of ["b", "B", "."]) expect(playbackActionForKey(key)).toBe("blackout");
    for (const key of ["w", "W", ","]) expect(playbackActionForKey(key)).toBe("whiteout");
  });

  it("the blank rows change no existing binding — every key keeps its action", () => {
    const existing: Array<[string, string]> = [
      ...["ArrowRight", "ArrowDown", " ", "Spacebar", "PageDown", "Enter", "n", "N"].map((k): [string, string] => [k, "advance"]),
      ...["ArrowLeft", "ArrowUp", "PageUp", "p", "P", "Backspace"].map((k): [string, string] => [k, "back"]),
      ["Home", "first"],
      ["End", "last"],
    ];
    for (const [key, action] of existing) {
      expect(playbackActionForKey(key), key).toBe(action);
    }
    expect(advanceKeys).toEqual(["ArrowRight", "ArrowDown", " ", "Spacebar", "PageDown", "Enter", "n"]);
    expect(backKeys).toEqual(["ArrowLeft", "ArrowUp", "PageUp", "p", "Backspace"]);
    expect(firstKeys).toEqual(["Home"]);
    expect(lastKeys).toEqual(["End"]);
    // No blank key overlaps a navigation key.
    for (const key of [...blackoutKeys, ...whiteoutKeys]) {
      expect([...advanceKeys, ...backKeys, ...firstKeys, ...lastKeys]).not.toContain(key);
    }
    // Keys outside the map still own nothing.
    expect(playbackActionForKey("g")).toBeNull();
    expect(playbackActionForKey("f")).toBeNull();
  });
});
