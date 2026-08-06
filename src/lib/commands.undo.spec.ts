import { describe, expect, it } from "vitest";
import * as commands from "./commands";
import { caretSelection, type CommandResult, type ModelSelection } from "./commands";
import { createDocument, type DocumentModel, type Paragraph } from "./documentModel";
import type { CaretPosition } from "./textEditing";
import { equalStructural } from "./equality";
import { clearHistory, recordEdit, undo } from "./undoHistory";

// SPEC-015 AC-20/AC-21 — "every command introduced by any later spec is
// undoable". The undo history (undoHistory.ts) is snapshot-based, not a
// per-command inverse: an undo step just restores the untouched `before`
// model every command (see commands.ts's header comment) promises never to
// mutate. So a command is undoable for free as long as it upholds that
// promise — which is exactly what this test checks, end-to-end through the
// real undo history module, for every command commands.ts exports.
//
// REGISTRY below must have one entry per command. QUERIES lists every other
// exported function — read-only queries, selection builders and the
// `.vue`-script typed-null seeds (see their doc comments in commands.ts) —
// which are not commands and carry no undo obligation. The completeness
// check below fails if a newly exported function lands in neither list,
// which is how a later spec's new command gets caught if nobody wires it
// into this test.

function paragraph(...runs: Array<[string, Record<string, unknown>?]>): Paragraph {
  return {
    type: "paragraph",
    runs: runs.map(([text, formatting]) => ({ text, formatting: formatting ?? {} })),
  };
}

function caret(block: number, offset: number, section = 0): CaretPosition {
  return { section, block, offset };
}

function range(anchor: CaretPosition, focus: CaretPosition): ModelSelection {
  return { anchor, focus };
}

/** One fixture document, freshly built for every scenario: two paragraphs
 * with plain and bold text, enough for every command — text edits,
 * character/run formatting, paragraph formatting and page margins — to have
 * something meaningful to touch, plus one list item (SPEC-018) so the list
 * commands genuinely mutate. */
function fixtureDocument(): DocumentModel {
  const model = createDocument("undo-fixture");
  const listItem = paragraph(["List item", {}]);
  listItem.list = { type: "bullet", level: 2 };
  model.sections = [
    {
      blocks: [
        paragraph(["Hello ", {}], ["world", { bold: true }]),
        paragraph(["Second paragraph", {}]),
        listItem,
      ],
    },
  ];
  return model;
}

type Scenario = (model: DocumentModel) => CommandResult;

const REGISTRY: Record<string, Scenario> = {
  insertText: (model) => commands.insertText(model, caretSelection(caret(0, 0)), "X"),
  insertLineBreak: (model) => commands.insertLineBreak(model, caretSelection(caret(0, 3))),
  splitBlock: (model) => commands.splitBlock(model, caretSelection(caret(0, 3))),
  deleteRange: (model) => commands.deleteRange(model, range(caret(0, 0), caret(0, 5))),
  deleteBackward: (model) => commands.deleteBackward(model, caretSelection(caret(0, 3))),
  deleteForward: (model) => commands.deleteForward(model, caretSelection(caret(0, 3))),
  setFormat: (model) => commands.setFormat(model, range(caret(0, 0), caret(0, 5)), "bold", true),
  toggleFormat: (model) => commands.toggleFormat(model, range(caret(0, 0), caret(0, 5)), "italic"),
  applyFormat: (model) =>
    commands.applyFormat(model, range(caret(0, 6), caret(0, 11)), { fontFamily: "Georgia" }),
  toggleVerticalAlign: (model) =>
    commands.toggleVerticalAlign(model, range(caret(0, 0), caret(0, 5)), "superscript"),
  applyParagraphFormat: (model) =>
    commands.applyParagraphFormat(model, caretSelection(caret(0, 0)), { alignment: "center" }),
  setAlignment: (model) => commands.setAlignment(model, caretSelection(caret(0, 0)), "right"),
  adjustIndent: (model) =>
    commands.adjustIndent(model, caretSelection(caret(0, 0)), commands.INDENT_STEP_CM),
  toggleSpaceBefore: (model) => commands.toggleSpaceBefore(model, caretSelection(caret(0, 0))),
  setPageMargins: (model) =>
    commands.setPageMargins(model, caretSelection(caret(0, 0)), { leftCm: 3 }),
  setPageSetup: (model) =>
    commands.setPageSetup(model, caretSelection(caret(0, 0)), {
      pageSize: { widthCm: 29.7, heightCm: 21 },
      margins: { topCm: 3, rightCm: 3, bottomCm: 3, leftCm: 3 },
      measurementUnit: "in",
    }),
  toggleList: (model) => commands.toggleList(model, caretSelection(caret(0, 0)), "bullet"),
  changeListLevel: (model) => commands.changeListLevel(model, caretSelection(caret(2, 0)), 1),
  setListMarker: (model) =>
    commands.setListMarker(model, caretSelection(caret(2, 0)), "square"),
  removeListFormat: (model) => commands.removeListFormat(model, caretSelection(caret(2, 0))),
  applyListAutoformat: (model) =>
    commands.applyListAutoformat(model, caretSelection(caret(1, 0)), "number"),
  applyParagraphStyle: (model) =>
    commands.applyParagraphStyle(model, caretSelection(caret(0, 0)), "Heading1"),
  resetParagraphStyle: (model) =>
    commands.resetParagraphStyle(model, caretSelection(caret(0, 8))),
  updateStyleFromSelection: (model) =>
    commands.updateStyleFromSelection(model, caretSelection(caret(0, 0)), "Heading1"),
  clearFormatting: (model) =>
    commands.clearFormatting(model, range(caret(0, 0), caret(0, 5))),
  applyPainterFormat: (model) =>
    commands.applyPainterFormat(model, range(caret(1, 0), caret(1, 6)), { run: { italic: true } }),
};

const QUERIES = new Set([
  "caretSelection",
  "noSelection",
  "noCommandResult",
  "noParagraphCommand",
  "noPendingFormat",
  "isCollapsed",
  "orderedRange",
  "caretFormatting",
  "formatState",
  "formatValue",
  "verticalAlignState",
  "emptyToolbarFormatValues",
  "toolbarFormatValues",
  "touchedBlocks",
  "firstTouchedFormatting",
  "emptyParagraphFormatting",
  "hasSpaceBefore",
  "documentStart",
  "documentEnd",
  "selectAll",
  "blockCount",
  "blockOrdinal",
  "allTouchedOfListType",
  "emptyToolbarListState",
  "toolbarListState",
  "emptyToolbarStyleState",
  "toolbarStyleState",
  "noPainterFormat",
  "isWholeParagraphSelected",
  "capturePainterFormat",
]);

describe("SPEC-015 AC-20 — every command is registered for the undo test", () => {
  it("has no commands.ts export that neither REGISTRY nor QUERIES accounts for", () => {
    const unaccounted = Object.keys(commands).filter((name) => {
      const value = (commands as Record<string, unknown>)[name];
      return typeof value === "function" && !(name in REGISTRY) && !QUERIES.has(name);
    });
    expect(
      unaccounted,
      "a new commands.ts export was found that is neither a registered command " +
        "(add it to REGISTRY) nor a documented query (add it to QUERIES) in commands.undo.spec.ts",
    ).toEqual([]);
  });
});

describe("SPEC-015 AC-21 — applying then undoing any registered command restores the prior state", () => {
  for (const [name, run] of Object.entries(REGISTRY)) {
    it(name, () => {
      const docId = `undo-registry-${name}`;
      const model = fixtureDocument();
      const before = structuredClone(model);

      const result = run(model);

      recordEdit(docId, {
        before: { model, selection: caretSelection(caret(0, 0)) },
        after: { model: result.model, selection: result.selection },
        label: name,
        groupable: false,
      });
      const restored = undo(docId);
      clearHistory(docId);

      expect(restored).not.toBeNull();
      expect(equalStructural(restored?.model, before)).toBe(true);
    });
  }
});
