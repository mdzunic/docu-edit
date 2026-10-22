import { describe, expect, it } from "vitest";
import {
  EQUATION_SEQ_CODE,
  insertDisplayEquation,
  insertInlineEquation,
  insertNumberedDisplayEquation,
} from "./commands.equations";
import { createDocument, type DocumentModel, type Paragraph } from "./documentModel";
import { parseLinear } from "./math/linearParse";

function docWithText(text: string): DocumentModel {
  const model = createDocument("d");
  const para = model.sections[0].blocks[0] as Paragraph;
  para.runs = text ? [{ text, formatting: {} }] : [];
  return model;
}

const caretAt = (block: number, offset: number) => ({
  anchor: { section: 0, block, offset },
  focus: { section: 0, block, offset },
});

describe("insertInlineEquation (AC-1)", () => {
  it("inserts an empty inline equation run at the caret", () => {
    const model = docWithText("ab");
    const { model: out } = insertInlineEquation(model, caretAt(0, 1));
    const runs = (out.sections[0].blocks[0] as Paragraph).runs;
    const eqRun = runs.find((r) => r.equation);
    expect(eqRun?.equation).toEqual({ display: "inline", body: [] });
    expect(eqRun?.text).toBe("1"); // one-character caret atom
  });

  it("can seed the equation from parsed linear input", () => {
    const model = docWithText("");
    const { model: out } = insertInlineEquation(model, caretAt(0, 0), parseLinear("\\frac{a}{b}"));
    const eqRun = (out.sections[0].blocks[0] as Paragraph).runs.find((r) => r.equation);
    expect(eqRun?.equation?.body[0].kind).toBe("frac");
  });
});

describe("insertDisplayEquation (AC-2)", () => {
  it("inserts a centred paragraph whose only run is the block equation", () => {
    const model = docWithText("");
    const { model: out, selection } = insertDisplayEquation(model, caretAt(0, 0));
    const para = out.sections[0].blocks[selection.focus.block] as Paragraph;
    expect(para.formatting?.alignment).toBe("center");
    expect(para.runs.length).toBe(1);
    expect(para.runs[0].equation).toEqual({ display: "block", body: [] });
  });

  it("splits the paragraph so the equation gets its own paragraph", () => {
    const model = docWithText("before after");
    const { model: out } = insertDisplayEquation(model, caretAt(0, 6));
    const blocks = out.sections[0].blocks as Paragraph[];
    // before | equation | after
    expect(blocks.some((b) => b.runs.some((r) => r.equation?.display === "block"))).toBe(true);
    expect(blocks[0].runs.map((r) => r.text).join("")).toContain("before");
  });
});

describe("insertNumberedDisplayEquation (AC-17/18)", () => {
  it("adds a SEQ Equation number at a right tab stop on the equation paragraph", () => {
    const model = docWithText("");
    const { model: out, selection } = insertNumberedDisplayEquation(model, caretAt(0, 0));
    const para = out.sections[0].blocks[selection.focus.block] as Paragraph;
    // The number is an unsupported SEQ field carrying the Equation label —
    // preserved verbatim, never reparsed (docx/fields.spec.ts fidelity guard).
    const seq = para.runs.find((r) => r.field?.kind === "unsupported");
    expect(seq?.field).toMatchObject({ kind: "unsupported", code: EQUATION_SEQ_CODE, cached: "1" });
    // A right tab stop pushes the number to the margin (SPEC-054).
    expect(para.formatting?.tabStops?.some((s) => s.alignment === "right")).toBe(true);
    // The equation is still centred and present.
    expect(para.formatting?.alignment).toBe("center");
    expect(para.runs[0].equation?.display).toBe("block");
    // There is a tab run before the number.
    expect(para.runs.some((r) => r.text === "\t")).toBe(true);
  });

  it("makes the equation a valid cross-reference target via a bookmark (AC-18)", () => {
    const model = docWithText("");
    const { model: out, selection } = insertNumberedDisplayEquation(model, caretAt(0, 0));
    const para = out.sections[0].blocks[selection.focus.block] as Paragraph;
    expect(para.bookmarks?.length).toBe(1);
    expect(para.bookmarks?.[0].name).toMatch(/^Equation\d+$/);
  });

  it("continues the sequence for a second numbered equation", () => {
    const model = docWithText("");
    const first = insertNumberedDisplayEquation(model, caretAt(0, 0));
    const secondCaret = { anchor: { section: 0, block: first.selection.focus.block + 1, offset: 0 }, focus: { section: 0, block: first.selection.focus.block + 1, offset: 0 } };
    const second = insertNumberedDisplayEquation(first.model, secondCaret);
    const numbers = second.model.sections[0].blocks
      .filter((b): b is Paragraph => b.type === "paragraph")
      .flatMap((b) => b.runs)
      .filter((r) => r.field?.kind === "unsupported" && r.field.code.startsWith("SEQ Equation"))
      .map((r) => (r.field?.kind === "unsupported" ? r.field.cached : ""));
    expect(numbers).toEqual(["1", "2"]);
    // Distinct bookmark names.
    const names = second.model.sections[0].blocks
      .filter((b): b is Paragraph => b.type === "paragraph")
      .flatMap((b) => b.bookmarks ?? [])
      .map((bm) => bm.name);
    expect(new Set(names).size).toBe(names.length);
  });
});
