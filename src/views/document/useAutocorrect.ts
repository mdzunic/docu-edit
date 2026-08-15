// Smart-quote/dash autocorrect and the special-character picker (SPEC-026).
// The substitution *decisions* are pure functions in lib/core/autocorrect;
// this composable owns the tracking state those decisions read (which
// substitution was produced at which caret offset) and records each
// substitution with the undo-step shape the spec demands: the plain insert
// first, then the substitution as its own step, so one Ctrl/Cmd+Z restores
// the literal typed character. The picker dialog and the non-breaking-space
// shortcut insert literally — never through the typed-input path, so
// autocorrect never sees them.
import { ref } from "vue";
import {
  dashSubstitution,
  decadeSubstitution,
  isDigit,
  loadAutocorrectSettings,
  saveAutocorrectSettings,
  smartQuote,
  type Replacement,
} from "../../lib/core/autocorrect";
import { caretSelection, deleteRange, documentStart, insertText } from "../../lib/commands";
import { blockText, paragraphAtCaret } from "../../lib/textEditing";
import { breakGroup } from "../../lib/undoHistory";
import type { CommandResult, ModelSelection } from "../../lib/commands";
import type { DocumentModel, PendingFormat } from "../../lib/documentModel";
import type { CaretPosition } from "../../lib/textEditing";
import type { EditorContext } from "./context";

export interface AutocorrectDeps extends EditorContext {
  /** Live read of the view's pending collapsed-caret format (SPEC-012): a
   * special-character insert applies it the same way typed text would. */
  getPendingFormat: () => PendingFormat | null;
}

/** The code point immediately before `index` in `text`, as a string (a full
 * surrogate pair when it is a non-BMP character), or null at the start. */
function codePointBefore(text: string, index: number): string | null {
  if (index <= 0) {
    return null;
  }
  const code = text.charCodeAt(index - 1);
  if (code >= 0xdc00 && code <= 0xdfff && index >= 2) {
    return text.slice(index - 2, index);
  }
  return text[index - 1];
}

/** Replaces [replacement.start, replacement.end) within the caret's paragraph
 * with replacement.text, returning a CommandResult whose caret sits after the
 * inserted text. Used to record a substitution as its own edit. */
function replaceInParagraph(
  model: DocumentModel,
  atCaret: CaretPosition,
  replacement: Replacement,
): CommandResult {
  const deleted = deleteRange(model, {
    anchor: { ...atCaret, offset: replacement.start },
    focus: { ...atCaret, offset: replacement.end },
  });
  return insertText(deleted.model, deleted.selection, replacement.text, null);
}

export function useAutocorrect(deps: AutocorrectDeps) {
  const {
    doc,
    contentRoot,
    editable,
    documentId,
    currentSelection,
    applyEdit,
    getCaret,
    getPendingFormat,
    onSessionClose,
  } = deps;

  // --- special characters and autocorrect (SPEC-026) --------------------------
  // The "Insert → Special characters" picker dialog.
  const specialCharacterDialogOpen = ref(false);
  // Smart-quote/dash autocorrect settings, persisted in localStorage.
  const autocorrectSettings = ref(loadAutocorrectSettings());
  // Tracking state for the follow-up substitutions: the caret offset just after
  // an en dash / opening single quote THIS feature produced (or -1 for none).
  // These are plain lets, not refs — they never drive rendering, only the next
  // keystroke's decision.
  let producedEnDashAt = -1;
  let producedOpenSingleAt = -1;

  /** Invalidates the substitution follow-up context (the produced-at offsets
   * above) — the view calls this when undo/redo rewrites text and moves the
   * caret (SPEC-026 AC-22). */
  function resetSubstitutionContext(): void {
    producedEnDashAt = -1;
    producedOpenSingleAt = -1;
  }

  /** Inserts a literal string at the current caret WITHOUT running autocorrect
   * — the character picker (AC-30) and the non-breaking-space shortcut (AC-11)
   * both go through here, never the typed-input path. */
  function insertSpecialCharacter(text: string): void {
    const root = contentRoot.value as HTMLElement | null;
    const model = doc.value;
    if (!root || !model || !editable.value || !text) {
      return;
    }
    const selection =
      currentSelection(root) ?? caretSelection(getCaret() ?? documentStart());
    applyEdit(
      insertText(model, selection, text, getPendingFormat()),
      selection,
      "insert character",
    );
    // A deliberate insert is not a substitution follow-up context.
    producedEnDashAt = -1;
    producedOpenSingleAt = -1;
  }

  function openSpecialCharacters(): void {
    specialCharacterDialogOpen.value = true;
  }

  function handleSpecialCharacterInsert(text: string): void {
    insertSpecialCharacter(text);
  }

  function toggleAutocorrectSmartQuotes(): void {
    const next = {
      ...autocorrectSettings.value,
      smartQuotes: !autocorrectSettings.value.smartQuotes,
    };
    autocorrectSettings.value = next;
    saveAutocorrectSettings(next);
  }

  function toggleAutocorrectDashes(): void {
    const next = {
      ...autocorrectSettings.value,
      dashes: !autocorrectSettings.value.dashes,
    };
    autocorrectSettings.value = next;
    saveAutocorrectSettings(next);
  }

  /** Runs the smart-quote/dash/decade substitutions after a single character was
   * plain-inserted (SPEC-026). `plainResult` is that plain insert. Returns true
   * when it recorded its own undo entries (the caller must then stop); false to
   * let the caller apply the plain insert normally. Only ever reached from the
   * typed-input branch of the view's handleBeforeInput (AC-27..31). */
  function applyAutocorrect(
    selection: ModelSelection | null,
    typed: string,
    plainResult: CommandResult | null,
  ): boolean {
    if (!selection || !plainResult) {
      return false;
    }
    const settings = autocorrectSettings.value;
    const caretAfter = plainResult.selection.focus;
    const blockAfter = paragraphAtCaret(plainResult.model, caretAfter);
    if (!blockAfter) {
      return false;
    }
    const text = blockText(blockAfter);
    const off = caretAfter.offset;

    // Decade idiom (AC-50): a digit typed right after our opening single quote
    // flips it to a closing quote as part of the digit's OWN edit, so one undo
    // removes the digit and reverts the flip together.
    if (isDigit(typed) && producedOpenSingleAt >= 0) {
      const rep = decadeSubstitution(text, off, typed, producedOpenSingleAt);
      if (rep) {
        const flipped = replaceInParagraph(plainResult.model, caretAfter, rep);
        applyEdit(
          { model: flipped.model, selection: caretSelection(caretAfter) },
          selection,
          "typing",
          { groupable: true },
        );
        producedOpenSingleAt = -1;
        producedEnDashAt = -1;
        return true;
      }
    }

    // Smart quotes (AC-18, AC-19): plain insert first, then the curly quote as a
    // separate undo entry so one Ctrl/Cmd+Z restores the straight quote (AC-21).
    if (settings.smartQuotes && (typed === "'" || typed === '"')) {
      const quote = smartQuote(typed, codePointBefore(text, off - 1));
      if (quote) {
        applyEdit(plainResult, selection, "typing", { groupable: true });
        breakGroup(documentId.value);
        const subResult = replaceInParagraph(plainResult.model, caretAfter, {
          start: off - 1,
          end: off,
          text: quote,
        });
        applyEdit(subResult, caretSelection(caretAfter), "smart quote");
        producedOpenSingleAt =
          quote === "‘" ? subResult.selection.focus.offset : -1;
        producedEnDashAt = -1;
        return true;
      }
    }

    // Dash substitution (AC-16, AC-17): same two-step undo shape.
    if (settings.dashes && typed === "-") {
      const sub = dashSubstitution(
        text,
        off,
        producedEnDashAt >= 0 ? producedEnDashAt : null,
      );
      if (sub) {
        applyEdit(plainResult, selection, "typing", { groupable: true });
        breakGroup(documentId.value);
        const subResult = replaceInParagraph(
          plainResult.model,
          caretAfter,
          sub.replacement,
        );
        applyEdit(
          subResult,
          caretSelection(caretAfter),
          sub.producesEnDash ? "en dash" : "em dash",
        );
        producedEnDashAt = sub.producesEnDash
          ? subResult.selection.focus.offset
          : -1;
        producedOpenSingleAt = -1;
        return true;
      }
    }

    // An ordinary typed character with no substitution invalidates any pending
    // follow-up context.
    producedEnDashAt = -1;
    producedOpenSingleAt = -1;
    return false;
  }

  /** Ends an autocorrect session: dialog closed, follow-up context cleared. */
  function resetAutocorrect(): void {
    specialCharacterDialogOpen.value = false;
    resetSubstitutionContext();
  }

  onSessionClose(resetAutocorrect);

  return {
    autocorrectSettings,
    toggleAutocorrectSmartQuotes,
    toggleAutocorrectDashes,
    applyAutocorrect,
    resetSubstitutionContext,
    insertSpecialCharacter,
    specialCharacterDialogOpen,
    openSpecialCharacters,
    handleSpecialCharacterInsert,
  };
}
