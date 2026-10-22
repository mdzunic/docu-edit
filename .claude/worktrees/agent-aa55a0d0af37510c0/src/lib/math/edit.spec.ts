// SPEC-056 in-equation editing engine tests — one rule per AC, pinned on the
// pure model so the interactions are provable without mounting a component.
import { describe, expect, it } from "vitest";
import {
  backspace,
  createEditState,
  editLinearForm,
  insertConstruct,
  insertGlyph,
  insertTemplate,
  moveLeft,
  moveRight,
  nextSlot,
  prevSlot,
  rowAt,
  typeString,
  CONSTRUCT_TEMPLATES,
  type EditState,
} from "./edit";
import type { MathNode } from "../documentModel";

const start = (body: MathNode[] = []) => createEditState(body);

describe("live linear-input typing (AC-4)", () => {
  it("turns \\frac{a}{b} into a fraction as the closing brace is typed", () => {
    let s = start();
    s = typeString(s, "\\frac{a}{b}");
    expect(s.body).toHaveLength(1);
    expect(s.body[0].kind).toBe("frac");
    const frac = s.body[0] as Extract<MathNode, { kind: "frac" }>;
    expect(frac.num[0]).toMatchObject({ kind: "sym", text: "a" });
    expect(frac.den[0]).toMatchObject({ kind: "sym", text: "b" });
  });

  it("renders the fraction before blur — it exists mid-way through typing", () => {
    let s = start();
    s = typeString(s, "\\frac{a}");
    // A fraction node is already present with the numerator filled.
    expect(s.body[0].kind).toBe("frac");
  });

  it("builds a script from x^2 live", () => {
    let s = start();
    s = typeString(s, "x^2");
    expect(s.body[0].kind).toBe("script");
  });

  it("resolves a \\alpha name against the catalogue while typing (AC-16)", () => {
    let s = start();
    s = typeString(s, "\\alpha");
    expect(s.body[0]).toMatchObject({ kind: "sym", text: "α" });
  });
});

describe("template palette (AC-5)", () => {
  it("offers a template for every construct kind in the set", () => {
    const kinds = new Set(CONSTRUCT_TEMPLATES.map((t) => t.make().kind));
    for (const kind of ["frac", "script", "radical", "nary", "delim", "matrix", "accent", "func", "text"]) {
      expect(kinds.has(kind as MathNode["kind"])).toBe(true);
    }
  });

  it("inserts a construct with empty slots and drops the caret into the first", () => {
    let s = start();
    s = insertTemplate(s, "frac");
    expect(s.body[0]).toMatchObject({ kind: "frac", num: [], den: [] });
    // Caret is inside the numerator (first slot).
    expect(s.caret.path).toEqual([{ index: 0, slot: "num" }]);
    expect(s.caret.offset).toBe(0);
  });

  it("typing after insertion fills the current slot", () => {
    let s = start();
    s = insertTemplate(s, "frac");
    s = typeString(s, "a");
    const frac = s.body[0] as Extract<MathNode, { kind: "frac" }>;
    expect(frac.num[0]).toMatchObject({ kind: "sym", text: "a" });
    expect(frac.den).toEqual([]);
  });
});

describe("Tab / Shift+Tab slot navigation (AC-6/7)", () => {
  it("Tab moves from the numerator to the denominator", () => {
    let s = insertTemplate(start(), "frac");
    s = typeString(s, "a");
    s = nextSlot(s);
    expect(s.caret.path).toEqual([{ index: 0, slot: "den" }]);
    s = typeString(s, "b");
    const frac = s.body[0] as Extract<MathNode, { kind: "frac" }>;
    expect(frac.den[0]).toMatchObject({ kind: "sym", text: "b" });
  });

  it("Tab from the last slot exits the construct", () => {
    let s = insertTemplate(start(), "frac");
    s = nextSlot(s); // -> den
    s = nextSlot(s); // -> out
    expect(s.caret.path).toEqual([]);
    expect(s.caret.offset).toBe(1);
  });

  it("Shift+Tab moves back to the previous slot", () => {
    let s = insertTemplate(start(), "frac");
    s = nextSlot(s); // den
    s = prevSlot(s); // num
    expect(s.caret.path).toEqual([{ index: 0, slot: "num" }]);
  });
});

describe("arrow structural navigation (AC-12)", () => {
  it("Right from the numerator's end enters the denominator", () => {
    let s = insertTemplate(start(), "frac");
    s = typeString(s, "a"); // caret at end of numerator
    s = moveRight(s);
    expect(s.caret.path).toEqual([{ index: 0, slot: "den" }]);
    expect(s.caret.offset).toBe(0);
  });

  it("Right descends into a construct from the row", () => {
    // body: frac, caret before it.
    const s0 = insertConstruct(start(), { kind: "frac", num: [], den: [], bar: true });
    // Move caret to top-level start.
    const s1: EditState = { ...s0, caret: { path: [], offset: 0 }, pending: "", runStart: 0 };
    const s2 = moveRight(s1);
    expect(s2.caret.path).toEqual([{ index: 0, slot: "num" }]);
  });

  it("Left mirrors Right", () => {
    let s = insertTemplate(start(), "frac");
    // caret in numerator, offset 0. Left exits before the construct.
    s = moveLeft(s);
    expect(s.caret.path).toEqual([]);
    expect(s.caret.offset).toBe(0);
  });
});

describe("backspace (AC-13)", () => {
  it("deletes a character of an in-progress run", () => {
    let s = typeString(start(), "ab");
    s = backspace(s);
    expect(editLinearForm(s)).toBe("a");
  });

  it("unwraps a construct at the start of its first slot, keeping content", () => {
    // Build \frac{a}{b}, then place caret at start of numerator.
    let s = insertTemplate(start(), "frac");
    s = typeString(s, "a");
    s = nextSlot(s);
    s = typeString(s, "b");
    // Caret is at end of denominator; go to start of numerator (first slot).
    s = prevSlot(s); // -> numerator, at end
    s = { ...s, caret: { path: s.caret.path, offset: 0 }, pending: "", runStart: 0 };
    s = backspace(s); // unwrap
    expect(s.body.every((n) => n.kind !== "frac")).toBe(true);
    expect(editLinearForm(s)).toContain("a");
    expect(editLinearForm(s)).toContain("b");
  });

  it("deletes the atom to the left at a row position", () => {
    let s = typeString(start(), "a");
    s = insertGlyph(s, "β");
    s = backspace(s);
    expect(editLinearForm(s)).toBe("a");
  });
});

describe("character-picker glyph insertion (AC-14/15)", () => {
  it("inserts a Maths glyph as an upright sym", () => {
    let s = insertGlyph(start(), "∞");
    expect(s.body[0]).toMatchObject({ kind: "sym", text: "∞", variant: "upright" });
  });

  it("inserts a Greek glyph as an italic identifier", () => {
    let s = insertGlyph(start(), "β");
    expect(s.body[0]).toMatchObject({ kind: "sym", text: "β", variant: "italic" });
  });

  it("inserts into the active slot", () => {
    let s = insertTemplate(start(), "sqrt");
    s = insertGlyph(s, "π");
    const rad = s.body[0] as Extract<MathNode, { kind: "radical" }>;
    expect(rad.radicand[0]).toMatchObject({ kind: "sym", text: "π" });
  });
});

describe("rowAt", () => {
  it("resolves a nested slot row", () => {
    const s = insertTemplate(start(), "frac");
    const row = rowAt(s.body, [{ index: 0, slot: "num" }]);
    expect(row).toEqual([]);
  });
});
