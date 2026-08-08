// Format painter (SPEC-020). Copies a selection's formatting and applies it to
// a later selection, in single-use ("once", one click) or repeat mode
// ("repeat", double-click, stays armed until Escape). Owns its armed state
// (the canvas cursor and button's pressed state render from `painterMode`) and
// guards against re-applying to the selection it last acted on.
import { ref } from "vue";
import type { Ref } from "vue";
import {
  applyPainterFormat,
  capturePainterFormat,
  isCollapsed,
  noPainterFormat,
  noSelection,
  type PainterFormat,
} from "../../lib/commands";
import { compareCaret } from "../../lib/textEditing";
import type { ModelSelection } from "../../lib/commands";
import type { DocumentModel } from "../../lib/documentModel";

export interface FormatPainterDeps {
  doc: Ref<DocumentModel | null>;
  contentRoot: Ref<Element | null>;
  editable: Readonly<Ref<boolean>>;
  currentSelection: (root?: HTMLElement | null) => ModelSelection | null;
  applyEdit: (
    result: { model: DocumentModel; selection: ModelSelection } | null,
    beforeSelection: ModelSelection | null,
    label: string,
    groupable?: boolean,
  ) => void;
}

export function useFormatPainter(deps: FormatPainterDeps) {
  const { doc, contentRoot, editable, currentSelection, applyEdit } = deps;

  const painterMode = ref("" as "" | "once" | "repeat");
  let painterFormat: PainterFormat | null = noPainterFormat();
  // The selection maybeApplyPainter last acted on — guards against its own
  // applyEdit's caret-restore echo re-triggering the same application.
  let lastPainterSelection: ModelSelection | null = noSelection();

  /** Disarms the format painter (AC-15..17): drops the copied format and the
   * armed cursor/button state. Idempotent, so Escape can call it unconditionally. */
  function disarmPainter(): void {
    painterMode.value = "";
    painterFormat = noPainterFormat();
    lastPainterSelection = noSelection();
  }

  /** Arms the painter from the current caret/selection (AC-9, AC-10). */
  function armPainter(mode: "once" | "repeat" = "once"): void {
    const root = contentRoot.value as HTMLElement | null;
    const model = doc.value;
    if (!root || !model || !editable.value) {
      return;
    }
    const selection = currentSelection(root);
    if (!selection) {
      return;
    }
    painterFormat = capturePainterFormat(model, selection);
    painterMode.value = mode;
    lastPainterSelection = noSelection();
  }

  /** A single click arms single-use mode, or disarms if already armed. */
  function handlePainterClick(): void {
    if (painterMode.value !== "") {
      disarmPainter();
      return;
    }
    armPainter("once");
  }

  /** A double-click on the painter button arms repeat mode (AC-14). */
  function handlePainterDblClick(): void {
    armPainter("repeat");
  }

  /** Applies the armed painter to the selection just made (AC-12). Single-use
   * disarms after one application; repeat stays armed. A collapsed selection
   * is not "a selection made" — nothing is applied, the painter stays armed. */
  function maybeApplyPainter(): void {
    const root = contentRoot.value as HTMLElement | null;
    const model = doc.value;
    if (!root || !model || painterMode.value === "" || !painterFormat) {
      return;
    }
    const selection = currentSelection(root);
    if (!selection) {
      return;
    }
    const last = lastPainterSelection;
    if (
      last &&
      compareCaret(selection.anchor, last.anchor) === 0 &&
      compareCaret(selection.focus, last.focus) === 0
    ) {
      return;
    }
    lastPainterSelection = selection;
    if (isCollapsed(selection)) {
      return;
    }
    applyEdit(applyPainterFormat(model, selection, painterFormat), selection, "format painter", false);
    if (painterMode.value === "once") {
      disarmPainter();
    }
  }

  return {
    painterMode,
    disarmPainter,
    armPainter,
    handlePainterClick,
    handlePainterDblClick,
    maybeApplyPainter,
  };
}
