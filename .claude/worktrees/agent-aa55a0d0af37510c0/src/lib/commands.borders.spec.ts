// SPEC-074 AC-11/AC-17 — what the Borders and shading dialog's OK writes.
// The interesting case is "off": borders and shading are inherited from the
// paragraph style until the paragraph sets its own, so a paragraph whose box
// comes from its style cannot be unboxed by dropping the field — that
// re-inherits. It needs a value that is present and paints nothing.
import { describe, expect, it } from "vitest";
import { applyParagraphBorders, caretSelection, type ModelSelection } from "./commands";
import { NO_SHADING_FILL } from "./paragraphBorders";
import { paintedParagraphFormatting } from "./styles";
import {
  createDocument,
  type DocumentModel,
  type NamedStyle,
  type Paragraph,
  type ParagraphFormatting,
} from "./documentModel";
import type { CaretPosition } from "./textEditing";

const LINE = { style: "single", widthPt: 1, color: "#000000" } as const;
const BOX = { top: LINE, bottom: LINE, left: LINE, right: LINE };

function caret(block: number, offset = 0): CaretPosition {
  return { section: 0, block, offset };
}

function range(anchor: CaretPosition, focus: CaretPosition): ModelSelection {
  return { anchor, focus };
}

function paragraph(text: string, styleId?: string): Paragraph {
  const block: Paragraph = { type: "paragraph", runs: [{ text, formatting: {} }] };
  if (styleId !== undefined) {
    block.styleId = styleId;
  }
  return block;
}

/** A document of plain paragraphs, with an optional style catalogue — every
 * paragraph names no style, so what it inherits is Normal's (AC-17). */
function fixture(styles?: NamedStyle[], ...texts: string[]): DocumentModel {
  const model = createDocument("borders-commands");
  model.sections = [{ blocks: (texts.length > 0 ? texts : ["One", "Two"]).map((t) => paragraph(t)) }];
  if (styles) {
    model.styles = styles;
  }
  return model;
}

function formattingOf(model: DocumentModel, block = 0): ParagraphFormatting | undefined {
  return (model.sections[0].blocks[block] as Paragraph).formatting;
}

function paintedOf(model: DocumentModel, block = 0): ParagraphFormatting | undefined {
  return paintedParagraphFormatting(model, model.sections[0].blocks[block] as Paragraph);
}

describe("applyParagraphBorders — setting borders and shading (AC-11)", () => {
  it("applies to every paragraph the selection touches", () => {
    const model = fixture(undefined, "One", "Two", "Three");
    const { model: next } = applyParagraphBorders(model, range(caret(0, 1), caret(1, 2)), {
      borders: BOX,
      shadingFill: "#fff2cc",
    });
    expect(formattingOf(next, 0)?.borders).toEqual(BOX);
    expect(formattingOf(next, 1)?.shadingFill).toBe("#fff2cc");
    expect(formattingOf(next, 2)).toBeUndefined();
  });

  it("stores a theme-referenced shading fill verbatim (SPEC-077 AC-23)", () => {
    const model = fixture(undefined, "One");
    const reference = { theme: "accent1", lumMod: 0.2, lumOff: 0.8 } as const;
    const { model: next } = applyParagraphBorders(model, range(caret(0, 0), caret(0, 1)), {
      shadingFill: reference,
    });
    expect(formattingOf(next, 0)?.shadingFill).toEqual(reference);
  });
});

describe("applyParagraphBorders — switching off with nothing to override (AC-17)", () => {
  const plain: NamedStyle[] = [{ id: "Normal", name: "Normal" }];

  it("leaves no husk behind when the paragraph had borders of its own", () => {
    const boxed = applyParagraphBorders(fixture(plain), caretSelection(caret(0)), {
      borders: BOX,
      shadingFill: "#fff2cc",
    }).model;
    const cleared = applyParagraphBorders(boxed, caretSelection(caret(0)), {
      borders: null,
      shadingFill: null,
    }).model;
    // Nothing to override, so "off" is simply the absent field — an empty
    // borders object here would stop the paragraph merging with an ordinary
    // unbordered neighbour later (AC-12).
    expect(formattingOf(cleared, 0)?.borders).toBeUndefined();
    expect(formattingOf(cleared, 0)?.shadingFill).toBeUndefined();
  });

  it("leaves a paragraph that never had any untouched", () => {
    const next = applyParagraphBorders(fixture(plain), caretSelection(caret(0)), {
      borders: null,
      shadingFill: null,
    }).model;
    expect(formattingOf(next, 0)).toBeUndefined();
  });
});

describe("applyParagraphBorders — switching an inherited box off (AC-17)", () => {
  /** Normal itself carries the box, so every paragraph in the document
   * inherits it — including the ones that name no style at all. */
  const boxedNormal: NamedStyle[] = [
    { id: "Normal", name: "Normal", paragraphFormatting: { borders: BOX, shadingFill: "#ffd9d9" } },
  ];

  it("writes an override that is present and paints nothing", () => {
    const model = fixture(boxedNormal);
    expect(paintedOf(model, 0)?.borders).toEqual(BOX);

    const next = applyParagraphBorders(model, caretSelection(caret(0)), {
      borders: null,
      shadingFill: null,
    }).model;
    expect(formattingOf(next, 0)?.borders).toEqual({});
    expect(formattingOf(next, 0)?.shadingFill).toBe(NO_SHADING_FILL);
  });

  it("stops the style's box and fill actually painting", () => {
    const next = applyParagraphBorders(fixture(boxedNormal), caretSelection(caret(0)), {
      borders: null,
      shadingFill: null,
    }).model;
    const painted = paintedOf(next, 0);
    expect(painted?.borders).toEqual({});
    expect(painted?.shadingFill).toBe(NO_SHADING_FILL);
    // The untouched neighbour still inherits, so this really is an override
    // on one paragraph and not a change to the style.
    expect(paintedOf(next, 1)?.borders).toEqual(BOX);
  });

  it("decides borders and shading independently", () => {
    const fillOff = applyParagraphBorders(fixture(boxedNormal), caretSelection(caret(0)), {
      shadingFill: null,
    }).model;
    expect(formattingOf(fillOff, 0)?.borders).toBeUndefined();
    expect(paintedOf(fillOff, 0)?.borders).toEqual(BOX);
    expect(paintedOf(fillOff, 0)?.shadingFill).toBe(NO_SHADING_FILL);
  });

  it("takes a non-empty edge set as the override where the user picks one", () => {
    const next = applyParagraphBorders(fixture(boxedNormal), caretSelection(caret(0)), {
      borders: { top: LINE, bottom: LINE },
    }).model;
    expect(paintedOf(next, 0)?.borders).toEqual({ top: LINE, bottom: LINE });
  });

  it("comes back on when the paragraph is given borders again", () => {
    const off = applyParagraphBorders(fixture(boxedNormal), caretSelection(caret(0)), {
      borders: null,
      shadingFill: null,
    }).model;
    const on = applyParagraphBorders(off, caretSelection(caret(0)), {
      borders: BOX,
      shadingFill: "#fff2cc",
    }).model;
    expect(formattingOf(on, 0)?.borders).toEqual(BOX);
    expect(formattingOf(on, 0)?.shadingFill).toBe("#fff2cc");
  });

  it("resolves the style per paragraph, not once for the selection", () => {
    const model = fixture(
      [
        ...boxedNormal,
        { id: "Plain", name: "Plain", paragraphFormatting: {} } as NamedStyle,
      ],
      "One",
      "Two",
    );
    (model.sections[0].blocks[1] as Paragraph).styleId = "Plain";
    const next = applyParagraphBorders(model, range(caret(0, 0), caret(1, 3)), {
      borders: null,
      shadingFill: null,
    }).model;
    // The first inherits Normal's box and needs the override; the second's
    // own style paints nothing, so it needs no husk.
    expect(formattingOf(next, 0)?.borders).toEqual({});
    expect(formattingOf(next, 1)?.borders).toBeUndefined();
  });
});
