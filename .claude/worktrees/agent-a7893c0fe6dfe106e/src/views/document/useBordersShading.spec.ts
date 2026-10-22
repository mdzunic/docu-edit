// Format → Borders and shading… driven through the REAL wiring (SPEC-074
// AC-6..AC-11): the prefill openBordersDialog produces for the selection at
// the moment of opening, and what applyBordersShading commits. The dialog's
// own controls are covered by BordersShadingDialog.spec.ts; this file exists
// for the seam between the live (non-reactive) selection and the seed.
import { shallowRef, ref } from "vue";
import { describe, expect, it, vi } from "vitest";
import { createDocument, type DocumentModel, type ParagraphBorders } from "../../lib/documentModel";
import type { CommandResult, ModelSelection } from "../../lib/commands";
import { applyParagraphBorders } from "../../lib/commands.paragraph";

import { useBordersShading } from "./useBordersShading";

const BOX: ParagraphBorders = {
  top: { style: "single", widthPt: 0.5, color: "#000000" },
  bottom: { style: "single", widthPt: 0.5, color: "#000000" },
  left: { style: "single", widthPt: 0.5, color: "#000000" },
  right: { style: "single", widthPt: 0.5, color: "#000000" },
};

/** A two-paragraph document whose FIRST paragraph is boxed and filled. */
function twoParagraphs(): DocumentModel {
  const base = createDocument("d");
  const first = base.sections[0].blocks[0];
  const model: DocumentModel = {
    ...base,
    sections: [
      {
        ...base.sections[0],
        blocks: [first, { ...first }],
      },
    ],
  };
  const at = { section: 0, block: 0, offset: 0 };
  return applyParagraphBorders(model, { anchor: at, focus: at }, { borders: BOX, shadingFill: "#fff2cc" }).model;
}

function harness(model: DocumentModel = twoParagraphs()) {
  const doc = shallowRef<DocumentModel>(model);
  // The caret lives OUTSIDE Vue's reactivity, exactly as the view's does:
  // paragraphSelection reads document.getSelection() and a module-level
  // caret, neither of which a computed can track.
  let caret = { section: 0, block: 0, offset: 0 };
  const applyEdit = vi.fn();
  const deps = {
    doc,
    contentRoot: ref<Element | null>({ focus: () => {} } as unknown as Element),
    editable: ref(true),
    applyEdit,
    paragraphSelection: (): ModelSelection => ({ anchor: caret, focus: caret }),
  };
  const api = useBordersShading(deps as unknown as Parameters<typeof useBordersShading>[0]);
  return {
    api,
    doc,
    applyEdit,
    moveCaretTo(block: number) {
      caret = { section: 0, block, offset: 0 };
    },
  };
}

describe("Borders and shading prefill (SPEC-074 AC-6..AC-11)", () => {
  it("seeds the dialog from the first paragraph the selection touches", () => {
    const { api } = harness();
    api.openBordersDialog();
    expect(api.bordersDialogOpen.value).toBe(true);
    expect(api.bordersDialogInitial.value).toEqual({ borders: BOX, shadingFill: "#fff2cc" });
  });

  it("re-seeds on every open, even when only the caret moved in between", () => {
    const { api, moveCaretTo } = harness();
    // First open: on the boxed paragraph.
    api.openBordersDialog();
    expect(api.bordersDialogInitial.value.borders).toEqual(BOX);
    api.closeBordersDialog();

    // The regression: the caret moves to the plain paragraph with NO edit
    // in between (doc and contentRoot untouched), and the dialog reopens.
    // A computed over the selection stayed cached on the boxed paragraph,
    // so the plain one showed — and on OK received — the box and the fill.
    moveCaretTo(1);
    api.openBordersDialog();
    expect(api.bordersDialogOpen.value).toBe(true);
    expect(api.bordersDialogInitial.value).toEqual({ borders: undefined, shadingFill: undefined });

    // ...and back again, still without an edit.
    api.closeBordersDialog();
    moveCaretTo(0);
    api.openBordersDialog();
    expect(api.bordersDialogInitial.value).toEqual({ borders: BOX, shadingFill: "#fff2cc" });
  });

  it("OK applies to the selection as it is at that moment, not at open time", () => {
    const { api, applyEdit, moveCaretTo } = harness();
    moveCaretTo(1);
    api.openBordersDialog();
    api.applyBordersShading({ borders: BOX, shadingFill: null });
    expect(api.bordersDialogOpen.value).toBe(false);
    const result = applyEdit.mock.calls[0][0] as CommandResult;
    const blocks = result.model.sections[0].blocks;
    expect(blocks[1].type === "paragraph" && blocks[1].formatting?.borders).toEqual(BOX);
    // The boxed first paragraph keeps its fill; the plain one gained none.
    expect(blocks[0].type === "paragraph" && blocks[0].formatting?.shadingFill).toBe("#fff2cc");
    expect(blocks[1].type === "paragraph" && blocks[1].formatting?.shadingFill).toBeUndefined();
  });

  it("seeds from what the paragraph paints, its style's box included (AC-17)", () => {
    // No paragraph sets anything; Normal draws the box. Seeding from the
    // paragraph's DIRECT formatting would show "None" here, and OK would
    // then apply that None and strip a box the user never asked to lose.
    const model = twoParagraphs();
    model.styles = [
      { id: "Normal", name: "Normal", paragraphFormatting: { borders: BOX, shadingFill: "#ffd9d9" } },
    ];
    const { api, moveCaretTo } = harness(model);
    moveCaretTo(1);
    api.openBordersDialog();
    expect(api.bordersDialogInitial.value).toEqual({ borders: BOX, shadingFill: "#ffd9d9" });
  });

  it("shows an explicit no-fill override as no fill, not as a colour", () => {
    const model = twoParagraphs();
    model.styles = [
      { id: "Normal", name: "Normal", paragraphFormatting: { borders: BOX, shadingFill: "#ffd9d9" } },
    ];
    const at = { section: 0, block: 1, offset: 0 };
    const off = applyParagraphBorders(model, { anchor: at, focus: at }, {
      borders: null,
      shadingFill: null,
    }).model;
    const { api, moveCaretTo } = harness(off);
    moveCaretTo(1);
    api.openBordersDialog();
    expect(api.bordersDialogInitial.value).toEqual({ borders: {}, shadingFill: undefined });
  });

  it("does not open on a read-only document", () => {
    const doc = shallowRef<DocumentModel>(twoParagraphs());
    const at = { section: 0, block: 0, offset: 0 };
    const api = useBordersShading({
      doc,
      contentRoot: ref<Element | null>(null),
      editable: ref(false),
      applyEdit: vi.fn(),
      paragraphSelection: (): ModelSelection => ({ anchor: at, focus: at }),
    } as unknown as Parameters<typeof useBordersShading>[0]);
    api.openBordersDialog();
    expect(api.bordersDialogOpen.value).toBe(false);
  });
});
