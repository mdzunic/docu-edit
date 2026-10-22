// SPEC-064 AC-2a/3a/3b/5a regression: a vertical-ruler drag fires one edit per
// pointermove, and the marker being dragged (a bottom margin, footer band, or a
// lower table row) is often a page below the caret, so the user has scrolled it
// into view. Every drag branch must therefore pass { skipCaretScroll: true } to
// applyEdit — otherwise keepCaretVisible() snaps the viewport back to the caret
// mid-drag and desyncs the fixed pointer from clientToContentPx's page-rect
// mapping, so every following pointermove reads a wild value (QA scroll-jump).
import { shallowRef, ref } from "vue";
import { describe, expect, it, vi } from "vitest";
import { useVerticalRuler } from "./useVerticalRuler";
import { createDocument, type DocumentModel } from "../../lib/documentModel";
import type { ApplyEditOptions } from "./context";

interface ApplyEditCall {
  label: string;
  options?: ApplyEditOptions;
}

function harness() {
  const doc = shallowRef<DocumentModel>(createDocument("d"));
  const calls: ApplyEditCall[] = [];
  const applyEdit = vi.fn(
    (result: { model: DocumentModel } | null, _before: unknown, label: string, options?: ApplyEditOptions) => {
      calls.push({ label, options });
      if (result) doc.value = result.model;
    },
  );
  const bodySelection = { anchor: { section: 0, block: 0, offset: 0 }, focus: { section: 0, block: 0, offset: 0 } };
  const deps = {
    doc,
    contentRoot: ref<Element | null>({} as Element),
    editable: ref(true),
    applyEdit,
    currentSelection: () => bodySelection,
    clientToContentPx: (_x: number, y: number) => ({ x: 0, y }),
    zoom: ref(1),
    getCaret: () => ({ section: 0, block: 0, offset: 0 }),
    onScroll: () => {},
    pageRef: ref<Element | null>(null),
    paragraphSelection: () => bodySelection,
    getCaretSectionIndex: () => 0,
    activeHeaderFooter: ref(null),
    viewMode: ref("print"),
    caretPosition: ref({ section: 0, block: 0, offset: 0 }),
  };
  // Only the fields useVerticalRuler reads are supplied (see the *Deps convention).
  const ruler = useVerticalRuler(deps as unknown as Parameters<typeof useVerticalRuler>[0]);
  return { ruler, calls };
}

describe("useVerticalRuler drag skips the caret-scroll nudge (AC-2a/3a/3b)", () => {
  it("passes skipCaretScroll when dragging the top margin", () => {
    const { ruler, calls } = harness();
    ruler.handleVerticalRulerUpdate({ kind: "marginTopCm", valueCm: 3, row: 0, minHeightCm: 0 });
    expect(calls).toHaveLength(1);
    expect(calls[0].label).toBe("margins");
    expect(calls[0].options?.skipCaretScroll).toBe(true);
  });

  it("passes skipCaretScroll when dragging the bottom margin", () => {
    const { ruler, calls } = harness();
    ruler.handleVerticalRulerUpdate({ kind: "marginBottomCm", valueCm: 4, row: 0, minHeightCm: 0 });
    expect(calls[0].label).toBe("margins");
    expect(calls[0].options?.skipCaretScroll).toBe(true);
  });

  it("passes skipCaretScroll when dragging the header boundary", () => {
    const { ruler, calls } = harness();
    ruler.handleVerticalRulerUpdate({ kind: "headerDistanceCm", valueCm: 1.8, row: 0, minHeightCm: 0 });
    expect(calls[0].label).toBe("header distance");
    expect(calls[0].options?.skipCaretScroll).toBe(true);
  });

  it("passes skipCaretScroll when dragging the footer boundary", () => {
    const { ruler, calls } = harness();
    ruler.handleVerticalRulerUpdate({ kind: "footerDistanceCm", valueCm: 1.8, row: 0, minHeightCm: 0 });
    expect(calls[0].label).toBe("footer distance");
    expect(calls[0].options?.skipCaretScroll).toBe(true);
  });
});
