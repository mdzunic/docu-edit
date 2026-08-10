import { describe, expect, it } from "vitest";
import {
  applyCharacterStyle,
  applyParagraphStyle,
  applyTableStyle,
  caretSelection,
  createStyle,
  deleteStyle,
  modifyStyle,
  renameStyle,
  toolbarCharacterStyleState,
  type ModelSelection,
} from "./commands";
import { asParagraph, createDocument, type DocumentModel, type Paragraph, type Table } from "./documentModel";
import type { CaretPosition } from "./textEditing";
import { isStyleInUse, styleNameCollides } from "./styles";

// SPEC-042 — the custom styles manager's commands. `applyParagraphStyle`
// coverage lives in commands.styles.spec.ts (SPEC-019); these tests cover
// the four families/CRUD SPEC-042 adds, and — the sole prior QA failure —
// that modifying a parent style never corrupts a child style's identity
// (AC-11).

function paragraph(...runs: Array<[string, Record<string, unknown>?]>): Paragraph {
  return {
    type: "paragraph",
    runs: runs.map(([text, formatting]) => ({ text, formatting: formatting ?? {} })),
  };
}

function caret(block: number, offset: number): CaretPosition {
  return { section: 0, block, offset };
}

function range(anchor: CaretPosition, focus: CaretPosition): ModelSelection {
  return { anchor, focus };
}

function blocks(model: DocumentModel): Paragraph[] {
  return model.sections[0].blocks.map((block) => asParagraph(block));
}

function fixture(): DocumentModel {
  const model = createDocument("custom-styles-commands");
  model.sections = [
    {
      blocks: [
        paragraph(["First paragraph", {}]),
        paragraph(["Second ", {}], ["bold", { bold: true }]),
        paragraph(["Third paragraph", {}]),
      ],
    },
  ];
  return model;
}

function tableFixture(): DocumentModel {
  const model = createDocument("custom-styles-table");
  const table: Table = {
    type: "table",
    columnWidthsCm: [5, 5],
    rows: [
      {
        cells: [
          { blocks: [paragraph(["A", {}])] },
          { blocks: [paragraph(["B", {}])] },
        ],
      },
    ],
  };
  model.sections = [{ blocks: [table] }];
  return model;
}

describe("createStyle (AC-3)", () => {
  it("creates a paragraph style from given formatting and applies it to the selection", () => {
    const model = fixture();
    const { model: next } = createStyle(model, caretSelection(caret(0, 0)), {
      name: "My Body",
      type: "paragraph",
      basedOn: "Normal",
      runFormatting: { italic: true },
    });
    expect(next.styles?.find((s) => s.name === "My Body")).toMatchObject({
      basedOn: "Normal",
      runFormatting: { italic: true },
    });
    expect(blocks(next)[0].styleId).toBe(next.styles?.find((s) => s.name === "My Body")?.id);
  });

  it("mints a fresh id that never collides with an existing style", () => {
    const model = fixture();
    const first = createStyle(model, caretSelection(caret(0, 0)), { name: "Emphasis", type: "character" });
    const second = createStyle(first.model, caretSelection(caret(1, 0)), { name: "Emphasis 2", type: "character" });
    const ids = second.model.styles?.map((s) => s.id) ?? [];
    expect(new Set(ids).size).toBe(ids.length);
  });

  it("rejects a colliding name via styleNameCollides (AC-7 — checked by the dialog before calling createStyle)", () => {
    const model = fixture();
    expect(styleNameCollides(model, "Normal")).toBe(true);
    expect(styleNameCollides(model, "  normal  ")).toBe(true);
    expect(styleNameCollides(model, "Totally New")).toBe(false);
  });
});

describe("applyCharacterStyle (AC-9)", () => {
  it("applies a character style to a selection independently of the paragraph style", () => {
    const model = fixture();
    const withParagraphStyle = applyParagraphStyle(model, caretSelection(caret(0, 0)), "Heading1").model;
    const withCharStyle = createStyle(withParagraphStyle, range(caret(0, 0), caret(0, 5)), {
      name: "Emphasis Style",
      type: "character",
      runFormatting: { italic: true, color: "#ff0000" },
    }).model;
    const first = blocks(withCharStyle)[0];
    expect(first.styleId).toBe("Heading1");
    expect(first.runs[0].styleId).toBe(
      withCharStyle.styles?.find((s) => s.name === "Emphasis Style")?.id,
    );
    expect(first.runs[0].formatting).toMatchObject({ italic: true, color: "#ff0000" });
    // The rest of the paragraph keeps the paragraph style's own formatting,
    // untouched by the character style.
    expect(first.runs[1]?.styleId).toBeUndefined();
  });

  it("reports the character style at the caret independently of the paragraph readout", () => {
    const model = fixture();
    const charStyleId = createStyle(model, range(caret(0, 0), caret(0, 5)), {
      name: "Emphasis",
      type: "character",
    }).model.styles?.find((s) => s.name === "Emphasis")?.id;
    const applied = applyCharacterStyle(model, range(caret(0, 0), caret(0, 5)), charStyleId!).model;
    expect(toolbarCharacterStyleState(applied, caretSelection(caret(0, 2))).id).toBe(charStyleId);
    expect(toolbarCharacterStyleState(applied, caretSelection(caret(0, 10))).id).toBe("");
  });
});

describe("modifyStyle inheritance (AC-4, AC-5, AC-11 — the QA regression)", () => {
  it("re-bakes a child style's paragraphs without rewriting their styleId to the parent's", () => {
    let model = fixture();
    model = createStyle(model, caretSelection(caret(0, 0)), {
      name: "Parent Style",
      type: "paragraph",
      runFormatting: { bold: true },
    }).model;
    const parentId = model.styles!.find((s) => s.name === "Parent Style")!.id;
    model = createStyle(model, caretSelection(caret(1, 0)), {
      name: "Child Style",
      type: "paragraph",
      basedOn: parentId,
      runFormatting: { italic: true },
    }).model;
    const childId = model.styles!.find((s) => s.name === "Child Style")!.id;

    model = modifyStyle(model, caretSelection(caret(0, 0)), parentId, {
      name: "Parent Style",
      type: "paragraph",
      runFormatting: { bold: true, color: "#0000ff" },
    }).model;

    const childParagraph = blocks(model)[1];
    // The identity bug under test: the child paragraph must keep
    // referencing Child Style, not get silently re-pointed at Parent Style.
    expect(childParagraph.styleId).toBe(childId);
    expect(isStyleInUse(model, childId)).toBe(true);
    // But it still inherits the parent's new (non-overridden) property.
    expect(childParagraph.runs[0].formatting).toMatchObject({ italic: true, color: "#0000ff" });

    const parentParagraph = blocks(model)[0];
    expect(parentParagraph.styleId).toBe(parentId);
    expect(parentParagraph.runs[0].formatting).toMatchObject({ bold: true, color: "#0000ff" });
  });

  it("is a no-op for a style the editor could not fully model (unmodeled)", () => {
    const model = fixture();
    model.styles = [{ id: "Weird", name: "Weird", type: "table", unmodeled: true, rawXml: "<x/>" }];
    const { model: next } = modifyStyle(model, caretSelection(caret(0, 0)), "Weird", {
      name: "Renamed anyway",
      type: "table",
    });
    expect(next.styles?.[0]).toEqual(model.styles[0]);
  });
});

describe("renameStyle (AC-6)", () => {
  it("changes only the display name, leaving every paragraph's styleId untouched", () => {
    let model = fixture();
    model = createStyle(model, caretSelection(caret(0, 0)), { name: "Old Name", type: "paragraph" }).model;
    const id = model.styles!.find((s) => s.name === "Old Name")!.id;
    const { model: renamed } = renameStyle(model, caretSelection(caret(0, 0)), id, "New Name");
    expect(renamed.styles?.find((s) => s.id === id)?.name).toBe("New Name");
    expect(blocks(renamed)[0].styleId).toBe(id);
  });
});

describe("deleteStyle (AC-7, AC-8)", () => {
  it("reverts paragraphs to the parent style, or Normal if there is none", () => {
    let model = fixture();
    model = createStyle(model, caretSelection(caret(0, 0)), {
      name: "Parent",
      type: "paragraph",
      runFormatting: { bold: true },
    }).model;
    const parentId = model.styles!.find((s) => s.name === "Parent")!.id;
    model = createStyle(model, range(caret(1, 0), caret(2, 0)), {
      name: "Child",
      type: "paragraph",
      basedOn: parentId,
      runFormatting: { italic: true },
    }).model;

    model = deleteStyle(model, caretSelection(caret(0, 0)), parentId).model;
    // Child's basedOn pointed at Parent, which is gone — its own paragraphs
    // are untouched (deleteStyle only re-points *direct* users of the
    // deleted style, per AC-8's literal "paragraphs that used it").
    expect(model.styles?.some((s) => s.id === parentId)).toBe(false);

    const childId = model.styles!.find((s) => s.name === "Child")!.id;
    model = deleteStyle(model, caretSelection(caret(0, 0)), childId).model;
    expect(model.styles?.some((s) => s.id === childId)).toBe(false);
    expect(blocks(model)[1].styleId).toBeUndefined();
    expect(blocks(model)[2].styleId).toBeUndefined();
  });
});

describe("applyTableStyle (AC-1, AC-2 for table styles)", () => {
  it("bakes the style's borders/shading onto every non-covered cell and records the reference", () => {
    let model = tableFixture();
    model = createStyle(model, caretSelection(caret(0, 0)), {
      name: "Grid Style",
      type: "table",
      cellShadingFill: "#eeeeee",
      cellBorders: { top: { style: "single", widthPt: 1, color: "#000000" } },
    }).model;
    const table = model.sections[0].blocks[0] as Table;
    const styleId = model.styles!.find((s) => s.name === "Grid Style")!.id;
    expect(table.styleId).toBe(styleId);
    expect(table.rows[0].cells[0].shadingFill).toBe("#eeeeee");
    expect(table.rows[0].cells[0].borders?.top).toEqual({ style: "single", widthPt: 1, color: "#000000" });
    expect(isStyleInUse(model, styleId)).toBe(true);

    const cleared = applyTableStyle(model, caretSelection(caret(0, 0)), "").model;
    expect((cleared.sections[0].blocks[0] as Table).styleId).toBeUndefined();
  });
});
