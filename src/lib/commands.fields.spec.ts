import { describe, it, expect } from "vitest";
import {
  createDocument,
  createFieldSpecRun,
  type DocumentModel,
  type FieldSpec,
  type Paragraph,
  type TableOfContents,
} from "./documentModel";
import { caretSelection, commitHeaderFooterModel, withHeaderFooterModel, type ModelSelection } from "./commands";
import {
  fieldSpanningSelection,
  insertFieldCommand,
  setFieldLockedAtCaret,
  unlinkFieldAtCaret,
  updateAllFields,
  updateFieldAtCaret,
  type FieldUpdateContext,
} from "./commands.fields";
import { noTocPageNumbers } from "./toc";

const INSTANT = new Date(2026, 2, 4, 14, 30, 0);
const CTX: FieldUpdateContext = { now: INSTANT, fileName: "report.docx" };

function fieldDoc(...fields: FieldSpec[]): DocumentModel {
  const model = createDocument("cmd-fields");
  model.defaultStyles.language = "en-GB";
  model.sections = [{ blocks: fields.map((field) => ({ type: "paragraph", runs: [createFieldSpecRun(field)] })) }];
  model.properties = { author: "Ada", customProperties: [{ name: "Client", type: "text", value: "Acme" }] };
  return model;
}

function firstField(model: DocumentModel, block = 0): FieldSpec | undefined {
  return (model.sections[0].blocks[block] as Paragraph).runs[0].field;
}

function caretAt(block: number, offset: number): ModelSelection {
  return caretSelection({ section: 0, block, offset });
}

describe("insertFieldCommand (SPEC-053)", () => {
  it("inserts a field as a one-character caret atom", () => {
    const model = createDocument("x");
    model.sections = [{ blocks: [{ type: "paragraph", runs: [{ text: "hi", formatting: {} }] }] }];
    const out = insertFieldCommand(model, caretAt(0, 2), { kind: "author", cached: "" });
    const runs = (out.model.sections[0].blocks[0] as Paragraph).runs;
    const fieldRun = runs.find((r) => r.field);
    expect(fieldRun?.text).toHaveLength(1);
    expect(fieldRun?.field).toEqual({ kind: "author", cached: "" });
  });
});

describe("fieldSpanningSelection (SPEC-053 AC-7)", () => {
  it("returns the field when the selection spans exactly its atom", () => {
    const field: FieldSpec = { kind: "date", format: "d MMMM yyyy", cached: "4 March 2026" };
    const model = fieldDoc(field);
    const found = fieldSpanningSelection(
      model,
      { section: 0, block: 0, offset: 0 },
      { section: 0, block: 0, offset: 1 },
    );
    expect(found).toEqual(field);
  });

  it("returns null for a collapsed selection", () => {
    const model = fieldDoc({ kind: "author", cached: "Ada" });
    expect(
      fieldSpanningSelection(model, { section: 0, block: 0, offset: 0 }, { section: 0, block: 0, offset: 0 }),
    ).toBeNull();
  });

  it("returns null when the selection does not start at the field's edge", () => {
    const model = createDocument("span");
    model.sections = [
      {
        blocks: [
          {
            type: "paragraph",
            runs: [{ text: "ab", formatting: {} }, createFieldSpecRun({ kind: "author", cached: "Ada" })],
          },
        ],
      },
    ];
    // Offsets 1..2 sit inside the plain text, not on the field atom at 2..3.
    expect(
      fieldSpanningSelection(model, { section: 0, block: 0, offset: 1 }, { section: 0, block: 0, offset: 2 }),
    ).toBeNull();
  });

  it("ignores merge chips (SPEC-046 owns those)", () => {
    const model = fieldDoc({ kind: "merge", name: "client_name" });
    expect(
      fieldSpanningSelection(model, { section: 0, block: 0, offset: 0 }, { section: 0, block: 0, offset: 1 }),
    ).toBeNull();
  });
});

describe("F9 — updateFieldAtCaret (SPEC-053 AC-15)", () => {
  it("recomputes the field at the caret", () => {
    const model = fieldDoc({ kind: "date", format: "d MMMM yyyy", cached: "old value" });
    const out = updateFieldAtCaret(model, caretAt(0, 0), CTX);
    expect(firstField(out.model)).toEqual({ kind: "date", format: "d MMMM yyyy", cached: "4 March 2026" });
  });

  it("skips a locked field (AC-19 rule 4)", () => {
    const model = fieldDoc({ kind: "date", format: "d MMMM yyyy", cached: "old value", locked: true });
    const out = updateFieldAtCaret(model, caretAt(0, 0), CTX);
    expect(firstField(out.model)).toMatchObject({ cached: "old value" });
  });

  it("leaves other fields untouched", () => {
    const model = fieldDoc(
      { kind: "date", format: "d MMMM yyyy", cached: "old date" },
      { kind: "author", cached: "old author" },
    );
    const out = updateFieldAtCaret(model, caretAt(0, 0), CTX);
    expect(firstField(out.model, 0)).toMatchObject({ cached: "4 March 2026" });
    expect(firstField(out.model, 1)).toMatchObject({ cached: "old author" });
  });
});

// F9 with the caret in a header/footer band (AC-15 with AC-33). The band
// addresses its content as a throwaway single-section model, so the update
// runs in that coordinate space while values still come from the real
// document — the composition useHeadersFooters' `updateFieldInBand` performs.
describe("F9 in a header/footer band (SPEC-053 AC-15, AC-33)", () => {
  function headerDoc(...header: FieldSpec[]): DocumentModel {
    const model = createDocument("cmd-fields-header");
    model.defaultStyles.language = "en-GB";
    model.sections = [
      {
        blocks: [
          {
            type: "paragraph",
            runs: [{ text: "Body words go here", formatting: {} }],
            bookmarks: [{ name: "intro", offset: 0 }],
          },
        ],
        headers: {
          default: {
            blocks: header.map((field) => ({ type: "paragraph", runs: [createFieldSpecRun(field)] })),
          },
        },
      },
    ];
    return model;
  }

  function headerField(model: DocumentModel, block = 0): FieldSpec | undefined {
    const paragraph = model.sections[0].headers?.default?.blocks[block] as Paragraph;
    return paragraph.runs[0].field;
  }

  it("recomputes the header field at the caret and folds it back into the header", () => {
    const model = headerDoc({ kind: "date", format: "d MMMM yyyy", cached: "1 January 2020" });
    const synthetic = withHeaderFooterModel(model, 0, "header", "default");
    const out = updateFieldAtCaret(synthetic, caretAt(0, 0), CTX, model);
    const committed = commitHeaderFooterModel(model, 0, "header", "default", out.model);
    expect(headerField(committed)).toMatchObject({ cached: "4 March 2026" });
    // The band's own commit must not disturb the body it was addressed from.
    expect((committed.sections[0].blocks[0] as Paragraph).runs[0].text).toBe("Body words go here");
  });

  it("evaluates against the real document, not the band's synthetic model", () => {
    // Both would silently evaluate wrong against the synthetic model: the
    // header holds no bookmark to resolve, and its own word count is not the
    // document's.
    const model = headerDoc(
      { kind: "ref", target: "intro", show: "text", cached: "stale" },
      { kind: "wordCount", cached: "0" },
    );
    const synthetic = withHeaderFooterModel(model, 0, "header", "default");
    let out = updateFieldAtCaret(synthetic, caretAt(0, 0), CTX, model);
    out = updateFieldAtCaret(out.model, caretAt(1, 0), CTX, model);
    const committed = commitHeaderFooterModel(model, 0, "header", "default", out.model);
    expect(headerField(committed, 0)).toMatchObject({ cached: "Body words go here" });
    expect(headerField(committed, 1)).toMatchObject({ cached: "4" });
  });

  it("skips a locked header field", () => {
    const model = headerDoc({ kind: "date", format: "d MMMM yyyy", cached: "1 January 2020", locked: true });
    const synthetic = withHeaderFooterModel(model, 0, "header", "default");
    const out = updateFieldAtCaret(synthetic, caretAt(0, 0), CTX, model);
    const committed = commitHeaderFooterModel(model, 0, "header", "default", out.model);
    expect(headerField(committed)).toMatchObject({ cached: "1 January 2020" });
  });
});

describe("Update all fields (SPEC-053 AC-16..18, AC-28)", () => {
  it("updates every unlocked field and skips locked ones", () => {
    const model = fieldDoc(
      { kind: "date", format: "d MMMM yyyy", cached: "old" },
      { kind: "author", cached: "old" },
      { kind: "docProperty", name: "Client", cached: "old" },
      { kind: "author", cached: "keep", locked: true },
    );
    const out = updateAllFields(model, caretAt(0, 0), CTX, noTocPageNumbers());
    expect(firstField(out.model, 0)).toMatchObject({ cached: "4 March 2026" });
    expect(firstField(out.model, 1)).toMatchObject({ cached: "Ada" });
    expect(firstField(out.model, 2)).toMatchObject({ cached: "Acme" });
    expect(firstField(out.model, 3)).toMatchObject({ cached: "keep" });
  });

  it("lists a deleted-target cross-reference rather than failing (AC-28)", () => {
    const model = fieldDoc({ kind: "ref", target: "gone", show: "text", cached: "stale" });
    const out = updateAllFields(model, caretAt(0, 0), CTX, noTocPageNumbers());
    expect(out.brokenRefs).toContain("gone");
    expect(firstField(out.model, 0)).toMatchObject({ cached: "Error! Reference source not found." });
  });

  it("never recomputes an unsupported field (AC-43)", () => {
    const model = fieldDoc({ kind: "unsupported", code: "SEQ x", cached: "42" });
    const out = updateAllFields(model, caretAt(0, 0), CTX, noTocPageNumbers());
    expect(firstField(out.model, 0)).toEqual({ kind: "unsupported", code: "SEQ x", cached: "42" });
  });

  it("rebuilds a table of contents through the shared path (AC-17)", () => {
    const model = createDocument("toc");
    const toc: TableOfContents = {
      type: "toc",
      options: { levels: [1, 2], showPageNumbers: true, hyperlinksOnly: false },
      entries: [],
    };
    model.sections = [
      {
        blocks: [
          toc,
          { type: "paragraph", runs: [{ text: "Chapter One", formatting: {} }], headingLevel: 1 },
        ],
      },
    ];
    const out = updateAllFields(model, caretAt(0, 0), CTX, noTocPageNumbers());
    const rebuilt = out.model.sections[0].blocks[0] as TableOfContents;
    expect(rebuilt.entries.map((e) => e.text)).toContain("Chapter One");
  });
});

describe("Lock / Unlink (SPEC-053 AC-13, AC-14)", () => {
  it("locks and unlocks the caret field", () => {
    const model = fieldDoc({ kind: "author", cached: "Ada" });
    const locked = setFieldLockedAtCaret(model, caretAt(0, 0), true);
    expect(firstField(locked.model)).toMatchObject({ locked: true });
    const unlocked = setFieldLockedAtCaret(locked.model, caretAt(0, 0), false);
    expect((firstField(unlocked.model) as { locked?: true }).locked).toBeUndefined();
  });

  it("unlink replaces the field with plain text of its current value", () => {
    const model = fieldDoc({ kind: "date", format: "d MMMM yyyy", cached: "old" });
    const out = unlinkFieldAtCaret(model, caretAt(0, 0), CTX);
    const run = (out.model.sections[0].blocks[0] as Paragraph).runs[0];
    expect(run.field).toBeUndefined();
    expect(run.text).toBe("4 March 2026");
  });
});
