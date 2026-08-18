// SPEC-056 AC-1/2/3/17 regression: inserting an equation must hand focus to
// the equation editor's own surface, not let the document's caret-restore
// reclaim it. The wiring guarantee that makes that possible is that every
// fresh-insert path passes { skipCaretRestore: true } to applyEdit, so the
// contenteditable root is never re-selected (which in Chromium implicitly
// refocuses it) after EquationEditor.vue focuses its surface on mount. Editing
// an existing equation and committing it must NOT skip caret restore — there
// the caret belongs back in the document.
import { shallowRef, ref } from "vue";
import { describe, expect, it, vi } from "vitest";
import { useEquations } from "./useEquations";
import { createDocument, type DocumentModel, type Paragraph } from "../../lib/documentModel";
import type { ApplyEditOptions } from "./context";

interface ApplyEditCall {
  label: string;
  options?: ApplyEditOptions;
}

function harness() {
  const doc = shallowRef<DocumentModel>(createDocument("d"));
  const calls: ApplyEditCall[] = [];
  const applyEdit = vi.fn(
    (
      result: { model: DocumentModel } | null,
      _before: unknown,
      label: string,
      options?: ApplyEditOptions,
    ) => {
      calls.push({ label, options });
      if (result) doc.value = result.model;
    },
  );
  const caret = { section: 0, block: 0, offset: 0 };
  const deps = {
    doc,
    contentRoot: ref<Element | null>({} as Element),
    editable: ref(true),
    currentSelection: () => ({ anchor: caret, focus: caret }),
    applyEdit,
    getCaret: () => caret,
  };
  // Only the fields useEquations reads are supplied; the composable ignores
  // the rest of EditorContext (see the *Deps convention).
  return { deps: deps as unknown as Parameters<typeof useEquations>[0], calls, doc };
}

describe("useEquations focus hand-off (AC-1/2/3/17)", () => {
  it("skips caret restore when inserting an inline equation", () => {
    const { deps, calls } = harness();
    useEquations(deps).insertEquation();
    expect(calls).toHaveLength(1);
    expect(calls[0].label).toBe("insert equation");
    expect(calls[0].options?.skipCaretRestore).toBe(true);
  });

  it("skips caret restore when inserting a display equation", () => {
    const { deps, calls } = harness();
    useEquations(deps).insertDisplayEquationAtCaret();
    expect(calls[0].options?.skipCaretRestore).toBe(true);
  });

  it("skips caret restore when inserting a numbered display equation", () => {
    const { deps, calls } = harness();
    useEquations(deps).insertNumberedDisplayEquationAtCaret();
    expect(calls[0].options?.skipCaretRestore).toBe(true);
  });

  it("restores the document caret when committing an edited equation", () => {
    const { deps, calls } = harness();
    const eq = useEquations(deps);
    eq.insertEquation();
    // Type into the just-opened session, then commit — the caret must return
    // to the document, so this applyEdit must NOT skip caret restore.
    eq.equationType("a");
    eq.commitEquation();
    const commit = calls.find((c) => c.label === "edit equation");
    expect(commit).toBeDefined();
    expect(commit?.options?.skipCaretRestore).toBeUndefined();
  });
});
