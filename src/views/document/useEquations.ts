// SPEC-056 equation editing surface. Insert commands (AC-1/2/17) and the
// in-equation editing session that makes an equation editable rather than a
// read-only atom: after an insert (or when an existing equation is activated)
// an editor panel opens over the live equation, driving the pure edit engine
// (src/lib/math/edit.ts) for typing (AC-4), the template palette (AC-5),
// Tab/Shift+Tab slot navigation (AC-6/7), arrow structural walking (AC-12),
// backspace-unwrap (AC-13), character-picker glyph insertion (AC-14/15) and
// alt-text editing (AC-20). The engine and serialisers live in src/lib; this
// composable only bridges the editor context to them, following useNotes/
// useFields.
import { ref, shallowRef } from "vue";
import {
  caretSelection,
  documentStart,
  type ModelSelection,
} from "../../lib/commands.editing";
import {
  insertDisplayEquation,
  insertInlineEquation,
  insertNumberedDisplayEquation,
} from "../../lib/commands.equations";
import {
  backspace,
  createEditState,
  editLinearForm,
  insertGlyph,
  insertTemplate,
  moveDown,
  moveLeft,
  moveRight,
  moveUp,
  nextSlot,
  prevSlot,
  typeChar,
  type EditState,
} from "../../lib/math/edit";
import {
  cloneContainers,
  type DocumentModel,
  type EquationSpec,
  type MathNode,
  type Paragraph,
  type Run,
} from "../../lib/documentModel";
import type { EditorContext } from "./context";

export type EquationsDeps = EditorContext;

/** The shape of `useEquations`'s return — the object DocumentView hands to
 * EquationEditorHost.vue as its single `api` prop. */
export type EquationsApi = ReturnType<typeof useEquations>;

/** Typed-access seam for EquationEditorHost.vue: the espree constraint bars an
 * inline `as` cast in a `.vue` script, so the host casts its runtime `Object`
 * prop through this helper (same idiom as asChart/asVisibleCommentThreads). */
export function asEquationsApi(api: unknown): EquationsApi {
  return api as EquationsApi;
}

/** Where an equation run lives in the model, so an edit session can write its
 * new body back. `cell` descends into a table cell (row/column). */
interface EquationLocator {
  section: number;
  block: number;
  cell?: { row: number; column: number };
  runIndex: number;
}

export function useEquations(deps: EquationsDeps) {
  const { doc, contentRoot, editable, currentSelection, applyEdit, getCaret } = deps;

  // The open editing session, or null when no equation is being edited.
  const editState = shallowRef<EditState | null>(null);
  const editAltText = ref("");
  const editDisplay = ref<"inline" | "block">("inline");
  let locator: EquationLocator | null = null;

  function selectionNow(): ModelSelection | null {
    const root = contentRoot.value as HTMLElement | null;
    if (!root) return null;
    return currentSelection(root) ?? caretSelection(getCaret() ?? documentStart());
  }

  /** The paragraphs of a block, flattened (a table contributes its cells'). */
  function paragraphByLocator(model: DocumentModel, loc: EquationLocator): Paragraph | null {
    const block = model.sections[loc.section]?.blocks[loc.block];
    if (!block) return null;
    if (loc.cell) {
      if (block.type !== "table") return null;
      const para = block.rows[loc.cell.row]?.cells[loc.cell.column]?.blocks;
      const p = para?.find((b) => b.type === "paragraph");
      return (p as Paragraph) ?? null;
    }
    return block.type === "paragraph" ? block : null;
  }

  /** Finds the locator of the run carrying `spec` in the current model. */
  function locate(model: DocumentModel, spec: EquationSpec): EquationLocator | null {
    for (let s = 0; s < model.sections.length; s++) {
      const blocks = model.sections[s].blocks;
      for (let b = 0; b < blocks.length; b++) {
        const block = blocks[b];
        if (block.type === "paragraph") {
          const i = block.runs.findIndex((r) => r.equation === spec);
          if (i >= 0) return { section: s, block: b, runIndex: i };
        } else if (block.type === "table") {
          for (let r = 0; r < block.rows.length; r++) {
            for (let c = 0; c < block.rows[r].cells.length; c++) {
              const paras = block.rows[r].cells[c].blocks;
              for (const para of paras) {
                if (para.type !== "paragraph") continue;
                const i = para.runs.findIndex((run) => run.equation === spec);
                if (i >= 0) return { section: s, block: b, cell: { row: r, column: c }, runIndex: i };
              }
            }
          }
        }
      }
    }
    return null;
  }

  /** Opens the editor on the equation identified by `spec` in `model`. */
  function openEditor(model: DocumentModel, spec: EquationSpec) {
    const loc = locate(model, spec);
    if (!loc) return;
    locator = loc;
    editState.value = createEditState(spec.body, true);
    editDisplay.value = spec.display;
    editAltText.value = spec.altText ?? "";
  }

  /** Insert an empty inline equation at the caret (AC-1, and the Alt+= body),
   * then open the editor with the caret inside it. */
  function insertEquation() {
    const model = doc.value;
    const selection = selectionNow();
    if (!model || !editable.value || !selection) return;
    const command = insertInlineEquation(model, selection);
    applyEdit(command, selection, "insert equation", { skipCaretRestore: true });
    const spec = newestEquation(command.model);
    if (spec) openEditor(command.model, spec);
  }

  /** Insert a display equation centred on its own paragraph (AC-2). */
  function insertDisplayEquationAtCaret() {
    const model = doc.value;
    const selection = selectionNow();
    if (!model || !editable.value || !selection) return;
    const command = insertDisplayEquation(model, selection);
    applyEdit(command, selection, "insert display equation", { skipCaretRestore: true });
    const spec = newestEquation(command.model);
    if (spec) openEditor(command.model, spec);
  }

  /** Insert a numbered display equation (AC-17). */
  function insertNumberedDisplayEquationAtCaret() {
    const model = doc.value;
    const selection = selectionNow();
    if (!model || !editable.value || !selection) return;
    const command = insertNumberedDisplayEquation(model, selection);
    applyEdit(command, selection, "insert numbered equation", { skipCaretRestore: true });
    const spec = newestEquation(command.model);
    if (spec) openEditor(command.model, spec);
  }

  /** The equation whose body is empty — the one just inserted. Falls back to
   * the last equation in document order. */
  function newestEquation(model: DocumentModel): EquationSpec | null {
    let last: EquationSpec | null = null;
    for (const section of model.sections) {
      for (const block of section.blocks) {
        const paras: Paragraph[] =
          block.type === "paragraph"
            ? [block]
            : block.type === "table"
              ? block.rows.flatMap((row) => row.cells.flatMap((cell) => cell.blocks.filter((b): b is Paragraph => b.type === "paragraph")))
              : [];
        for (const para of paras) {
          for (const run of para.runs) {
            if (run.equation) {
              if (run.equation.body.length === 0) return run.equation;
              last = run.equation;
            }
          }
        }
      }
    }
    return last;
  }

  /** Activate the editor for an existing equation (a click on it). */
  function activateEquation(spec: unknown) {
    const model = doc.value;
    if (!model || !editable.value || !spec) return;
    openEditor(model, spec as EquationSpec);
  }

  /** Whether an equation editing session is open. */
  function isEditing(): boolean {
    return editState.value !== null;
  }

  function apply(next: EditState) {
    editState.value = next;
  }

  // Engine operations the editor panel and key handling call.
  const type = (text: string) => {
    if (!editState.value) return;
    let s = editState.value;
    for (const ch of text) s = typeChar(s, ch);
    apply(s);
  };
  const template = (id: string) => editState.value && apply(insertTemplate(editState.value, id));
  const glyph = (g: string) => editState.value && apply(insertGlyph(editState.value, g));
  const tab = () => editState.value && apply(nextSlot(editState.value));
  const shiftTab = () => editState.value && apply(prevSlot(editState.value));
  const left = () => editState.value && apply(moveLeft(editState.value));
  const right = () => editState.value && apply(moveRight(editState.value));
  const up = () => editState.value && apply(moveUp(editState.value));
  const down = () => editState.value && apply(moveDown(editState.value));
  const back = () => editState.value && apply(backspace(editState.value));

  /** Dispatch an EquationEditor "op" event to the matching engine call — the
   * single entry point the editor panel drives, so DocumentView needs no
   * per-operation wiring. */
  const op = (o: { kind: string; text?: string; id?: string }) => {
    switch (o.kind) {
      case "type": if (o.text) type(o.text); break;
      case "template": if (o.id) template(o.id); break;
      case "glyph": if (o.text) glyph(o.text); break;
      case "tab": tab(); break;
      case "shiftTab": shiftTab(); break;
      case "left": left(); break;
      case "right": right(); break;
      case "up": up(); break;
      case "down": down(); break;
      case "backspace": back(); break;
    }
  };

  /** Insert a glyph from SPEC-026's character picker into the equation
   * (AC-14/15). Returns true when the glyph was consumed by an open session. */
  function insertPickerGlyph(g: string): boolean {
    if (!editState.value) return false;
    glyph(g);
    return true;
  }

  function setAltText(text: string) {
    editAltText.value = text;
  }

  /** Build the edited equation spec from the current session. The default
   * (empty) alt text is stored as `undefined` so it keeps defaulting to the
   * (regenerated) linear form (AC-19). */
  function editedEquation(previous: EquationSpec, body: MathNode[]): EquationSpec {
    const equation: EquationSpec = { display: previous.display, body };
    if (editAltText.value !== "") {
      equation.altText = editAltText.value.trim() === "" ? "" : editAltText.value;
    }
    return equation;
  }

  /** Replace the run carrying the edited equation with a fresh run inside a
   * fresh paragraph/table block, so the block's object identity changes and the
   * canvas's `v-memo` repaints it (AC-21). Every command module replaces blocks
   * rather than mutating them (see commands.editing.ts); doing the same here is
   * what makes a just-edited equation actually appear in the document. */
  function writeEquation(next: DocumentModel, loc: EquationLocator, equation: EquationSpec) {
    const blocks = next.sections[loc.section]?.blocks;
    const block = blocks?.[loc.block];
    if (!blocks || !block) return;
    const replaceRun = (para: Paragraph): Paragraph => ({
      ...para,
      runs: para.runs.map((run, i) => (i === loc.runIndex ? { ...run, equation } : run)),
    });
    if (!loc.cell) {
      if (block.type === "paragraph") blocks[loc.block] = replaceRun(block);
      return;
    }
    if (block.type !== "table") return;
    const { row, column } = loc.cell;
    blocks[loc.block] = {
      ...block,
      rows: block.rows.map((r, ri) =>
        ri !== row
          ? r
          : {
              ...r,
              cells: r.cells.map((c, ci) =>
                ci !== column
                  ? c
                  : {
                      ...c,
                      blocks: c.blocks.map((b) =>
                        b.type === "paragraph" && b.runs[loc.runIndex]?.equation
                          ? replaceRun(b)
                          : b,
                      ),
                    },
              ),
            },
      ),
    };
  }

  /** Commit the session's body and alt text back into the model and close the
   * editor. */
  function commitAndClose() {
    const state = editState.value;
    const loc = locator;
    const model = doc.value;
    if (!state || !loc || !model || !editable.value) {
      close();
      return;
    }
    const before = selectionNow() ?? caretSelection(documentStart());
    const next = cloneContainers(model);
    const paragraph = paragraphByLocator(next, loc);
    const run: Run | undefined = paragraph?.runs[loc.runIndex];
    if (paragraph && run?.equation) {
      writeEquation(next, loc, editedEquation(run.equation, state.body));
      applyEdit({ model: next, selection: before }, before, "edit equation");
    }
    close();
  }

  function close() {
    editState.value = null;
    editAltText.value = "";
    locator = null;
  }

  return {
    insertEquation,
    insertDisplayEquationAtCaret,
    insertNumberedDisplayEquationAtCaret,
    activateEquation,
    isEditing,
    insertPickerGlyph,
    // editor session state + ops (consumed by EquationEditor.vue)
    equationEditState: editState,
    equationAltText: editAltText,
    equationDisplay: editDisplay,
    equationLinearForm: () => (editState.value ? editLinearForm(editState.value) : ""),
    equationOp: op,
    setEquationAltText: setAltText,
    commitEquation: commitAndClose,
    closeEquationEditor: close,
  };
}
