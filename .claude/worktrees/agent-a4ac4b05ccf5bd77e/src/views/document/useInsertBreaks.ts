// Insert → Break: a manual line/column/page break (Insert menu) and a
// section break at the caret. Both resolve the paragraph selection, apply
// the documentModel command as one undo entry, refocus the canvas, then
// follow the caret to its new — very likely off-screen — position. The
// Ctrl/Cmd+Enter page break stays in DocumentView (it uses the raw caret
// selection, not paragraphSelection).
import { asBreakKind, asSectionStart } from "../../lib/documentModel";
import { insertBreak, insertSectionBreak } from "../../lib/commands";
import type { ModelSelection } from "../../lib/commands";
import type { EditorContext } from "./context";

export interface InsertBreaksDeps extends EditorContext {
  /** DocumentView's paragraphSelection always resolves to a selection. */
  paragraphSelection: (root?: HTMLElement | null) => ModelSelection;
}

export function useInsertBreaks(deps: InsertBreaksDeps) {
  const { doc, contentRoot, editable, applyEdit, scrollModelCaretIntoView, getCaret, paragraphSelection } = deps;

  function handleInsertBreak(kind = "") {
    const root = contentRoot.value as HTMLElement | null;
    const model = doc.value;
    if (!root || !model || !editable.value) {
      return;
    }
    const selection = paragraphSelection(root);
    applyEdit(insertBreak(model, selection, asBreakKind(kind)), selection, "insert break");
    root.focus();
    void scrollModelCaretIntoView(getCaret() ?? undefined, true);
  }

  function handleInsertSectionBreak(start = "") {
    const root = contentRoot.value as HTMLElement | null;
    const model = doc.value;
    if (!root || !model || !editable.value) {
      return;
    }
    const selection = paragraphSelection(root);
    applyEdit(
      insertSectionBreak(model, selection, asSectionStart(start)),
      selection,
      "insert section break",
    );
    root.focus();
    void scrollModelCaretIntoView(getCaret() ?? undefined, true);
  }

  return { handleInsertBreak, handleInsertSectionBreak };
}
