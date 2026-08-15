import { describe, expect, it } from "vitest";
import { listMarkers, listRunAround } from "./lists";
import {
  bulletStyleForChar,
  clampListLevel,
  defaultMarkerStyle,
  detectListAutoformat,
  formatOrdinal,
  listIndentCm,
  markerStyleOf,
  markerText,
  markerType,
  planListLevels,
} from "./core/lists";
import type { DocumentModel, Paragraph } from "./documentModel";
import type { ListFormatting } from "./core/lists";
import { createDocument } from "./documentModel";

function paragraph(text: string, list?: ListFormatting): Paragraph {
  const p: Paragraph = {
    type: "paragraph",
    runs: text.length > 0 ? [{ text, formatting: {} }] : [],
  };
  if (list) {
    p.list = list;
  }
  return p;
}

function docWith(...blocks: Paragraph[]): DocumentModel {
  const model = createDocument("lists-fixture");
  model.sections = [{ blocks }];
  return model;
}

describe("SPEC-018 AC-5 — default marker cycles", () => {
  it("bullets cycle •, ◦, ▪ over levels 1..9", () => {
    const styles = [1, 2, 3, 4, 5, 6, 7, 8, 9].map((level) => defaultMarkerStyle("bullet", level));
    expect(styles).toEqual([
      "disc", "circle", "square",
      "disc", "circle", "square",
      "disc", "circle", "square",
    ]);
  });

  it("numbers cycle decimal, lower-alpha, lower-roman over levels 1..9", () => {
    const styles = [1, 2, 3, 4, 5, 6, 7, 8, 9].map((level) => defaultMarkerStyle("number", level));
    expect(styles).toEqual([
      "decimal", "lowerAlpha", "lowerRoman",
      "decimal", "lowerAlpha", "lowerRoman",
      "decimal", "lowerAlpha", "lowerRoman",
    ]);
  });

  it("renders the marker characters and ordinal formats of the spec", () => {
    expect(markerText("disc", 0)).toBe("•");
    expect(markerText("circle", 0)).toBe("◦");
    expect(markerText("square", 0)).toBe("▪");
    expect(markerText("decimal", 1)).toBe("1.");
    expect(markerText("lowerAlpha", 1)).toBe("a.");
    expect(markerText("lowerRoman", 1)).toBe("i.");
    expect(markerText("upperAlpha", 3)).toBe("C.");
    expect(markerText("upperRoman", 4)).toBe("IV.");
  });
});

describe("ordinal formatting", () => {
  it("formats alphabetics bijectively past z", () => {
    expect(formatOrdinal("lowerAlpha", 26)).toBe("z");
    expect(formatOrdinal("lowerAlpha", 27)).toBe("aa");
    expect(formatOrdinal("upperAlpha", 28)).toBe("AB");
  });

  it("formats roman numerals", () => {
    expect(formatOrdinal("lowerRoman", 3)).toBe("iii");
    expect(formatOrdinal("lowerRoman", 9)).toBe("ix");
    expect(formatOrdinal("upperRoman", 1944)).toBe("MCMXLIV");
  });
});

describe("marker helpers", () => {
  it("markerType classifies bullet and number styles", () => {
    expect(markerType("square")).toBe("bullet");
    expect(markerType("upperRoman")).toBe("number");
  });

  it("markerStyleOf uses the explicit marker only when it fits the type", () => {
    expect(markerStyleOf({ type: "bullet", level: 1, marker: "square" })).toBe("square");
    // A stale number marker on a bullet item falls back to the cycle.
    expect(markerStyleOf({ type: "bullet", level: 1, marker: "decimal" })).toBe("disc");
  });

  it("maps producer bullet characters onto the supported styles", () => {
    expect(bulletStyleForChar("•")).toBe("disc");
    expect(bulletStyleForChar("o")).toBe("circle");
    expect(bulletStyleForChar("■")).toBe("square");
    expect(bulletStyleForChar("★")).toBeUndefined();
  });

  it("clamps levels to 1..9", () => {
    expect(clampListLevel(0)).toBe(1);
    expect(clampListLevel(12)).toBe(9);
  });

  it("defaults the indent to level × 1.27 cm", () => {
    expect(listIndentCm({ type: "bullet", level: 3 })).toBeCloseTo(3.81, 10);
    expect(listIndentCm({ type: "bullet", level: 2, indentCm: 2 })).toBe(2);
  });
});

describe("SPEC-018 AC-10/AC-11/AC-22 — derived numbering", () => {
  it("numbers consecutive same-level items continuously (AC-10)", () => {
    const model = docWith(
      paragraph("one", { type: "number", level: 1 }),
      paragraph("two", { type: "number", level: 1 }),
      paragraph("three", { type: "number", level: 1 }),
    );
    expect(listMarkers(model)[0]).toEqual(["1.", "2.", "3."]);
  });

  it("restarts numbering after a non-list interruption (AC-11)", () => {
    const model = docWith(
      paragraph("one", { type: "number", level: 1 }),
      paragraph("two", { type: "number", level: 1 }),
      paragraph("plain"),
      paragraph("restart", { type: "number", level: 1 }),
    );
    expect(listMarkers(model)[0]).toEqual(["1.", "2.", "", "1."]);
  });

  it("restarts at the list's own start value after an interruption (AC-11, AC-22)", () => {
    const model = docWith(
      paragraph("five", { type: "number", level: 1, startAt: 5 }),
      paragraph("six", { type: "number", level: 1, startAt: 5 }),
      paragraph("plain"),
      paragraph("five again", { type: "number", level: 1, startAt: 5 }),
    );
    expect(listMarkers(model)[0]).toEqual(["5.", "6.", "", "5."]);
  });

  it("numbers sub-levels independently and restarts them after a shallower item", () => {
    const model = docWith(
      paragraph("1", { type: "number", level: 1 }),
      paragraph("1a", { type: "number", level: 2 }),
      paragraph("1b", { type: "number", level: 2 }),
      paragraph("2", { type: "number", level: 1 }),
      paragraph("2a", { type: "number", level: 2 }),
    );
    expect(listMarkers(model)[0]).toEqual(["1.", "a.", "b.", "2.", "a."]);
  });

  it("keeps numbering across an interleaved bullet item (still a list item)", () => {
    const model = docWith(
      paragraph("one", { type: "number", level: 1 }),
      paragraph("dash", { type: "bullet", level: 1 }),
      paragraph("two", { type: "number", level: 1 }),
    );
    expect(listMarkers(model)[0]).toEqual(["1.", "•", "2."]);
  });

  it("uses explicit markers and level-cycled defaults (AC-5)", () => {
    const model = docWith(
      paragraph("a", { type: "bullet", level: 1 }),
      paragraph("b", { type: "bullet", level: 2 }),
      paragraph("c", { type: "bullet", level: 4 }),
      paragraph("d", { type: "bullet", level: 2, marker: "square" }),
      paragraph("e", { type: "number", level: 5, startAt: 3 }),
    );
    // Level 4 cycles back to •; the explicit ▪ overrides level 2's ◦;
    // level 5 numbers in lower-alpha from its start value.
    expect(listMarkers(model)[0]).toEqual(["•", "◦", "•", "▪", "c."]);
  });
});

describe("listRunAround — the current list (AC-13)", () => {
  it("finds the contiguous run of list paragraphs around a block", () => {
    const model = docWith(
      paragraph("plain"),
      paragraph("a", { type: "bullet", level: 1 }),
      paragraph("b", { type: "bullet", level: 2 }),
      paragraph("c", { type: "number", level: 1 }),
      paragraph("plain again"),
      paragraph("d", { type: "bullet", level: 1 }),
    );
    expect(listRunAround(model.sections[0].blocks, 2)).toEqual({ first: 1, last: 3 });
    expect(listRunAround(model.sections[0].blocks, 5)).toEqual({ first: 5, last: 5 });
    expect(listRunAround(model.sections[0].blocks, 0)).toBeNull();
  });
});

describe("SPEC-018 AC-14/AC-15/AC-17 — autoformat detection", () => {
  it("recognises exactly the four trigger prefixes", () => {
    expect(detectListAutoformat("- ")).toBe("bullet");
    expect(detectListAutoformat("* ")).toBe("bullet");
    expect(detectListAutoformat("1. ")).toBe("number");
    expect(detectListAutoformat("1) ")).toBe("number");
  });

  it("rejects anything with text or whitespace before the prefix (AC-17)", () => {
    expect(detectListAutoformat(" - ")).toBeNull();
    expect(detectListAutoformat("x- ")).toBeNull();
    expect(detectListAutoformat("2. ")).toBeNull();
    expect(detectListAutoformat("-")).toBeNull();
    expect(detectListAutoformat("- x")).toBeNull();
  });
});

describe("planListLevels — export planning", () => {
  it("defines each level from the run's first item at that level", () => {
    const levels = planListLevels([
      { type: "number", level: 1, startAt: 4 },
      { type: "number", level: 2, marker: "upperRoman" },
      { type: "number", level: 1, startAt: 9 }, // later item does not redefine
    ]);
    expect(levels[0]).toEqual({ type: "number", style: "decimal", startAt: 4, indentCm: 1.27 });
    expect(levels[1]).toEqual({ type: "number", style: "upperRoman", startAt: 1, indentCm: 2.54 });
    // Unused levels fall back to the leading type's default cycle.
    expect(levels[2].style).toBe("lowerRoman");
    expect(levels).toHaveLength(9);
  });
});
