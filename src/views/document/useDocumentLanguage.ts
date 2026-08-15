// The language selector and status bar indicator (SPEC-035 AC-9, AC-9.1,
// AC-15, AC-15.1). Mirrors useWordCount.ts's selection-tracking shape:
// DocumentView's central updateCaretStatus calls setSelection alongside its
// other per-selection state, so the indicator recomputes on every
// caret/selection change without this composable owning any DOM listeners
// of its own.
import { computed, ref } from "vue";
import type { Ref } from "vue";
import { applyFormat, caretSelection, isCollapsed, setDocumentLanguage } from "../../lib/commands";
import type { CommandResult, ModelSelection } from "../../lib/commands";
import { formattingAtOffset, paragraphAt } from "../../lib/textEditing";
import type { CaretPosition } from "../../lib/textEditing";
import type { DocumentModel } from "../../lib/documentModel";
import { DEFAULT_LANGUAGE, effectiveLanguage, languageLabel } from "../../lib/core/language";

export interface DocumentLanguageDeps {
  doc: Ref<DocumentModel | null>;
  applyEdit: (result: CommandResult | null, beforeSelection: ModelSelection | null, label: string) => void;
}

export function useDocumentLanguage(deps: DocumentLanguageDeps) {
  const { doc, applyEdit } = deps;

  const languageDialogOpen = ref(false);
  const selectionActive = ref(false);
  const currentLanguageTag = ref(DEFAULT_LANGUAGE);
  /** Whether the selection's first run is already marked "do not check
   * spelling" (SPEC-035 AC-11) — the popover's checkbox state. Only
   * meaningful while selectionActive; a real selection is required to
   * apply the attribute at all (there is no document-wide "never check"
   * default). */
  const selectionNoProof = ref(false);
  let lastSelection: ModelSelection | null = null;

  /** Called from DocumentView's updateCaretStatus, the single place every
   * caret/selection change already flows through (same convention as
   * useWordCount's setSelection). */
  function setSelection(start: CaretPosition | null, end: CaretPosition | null) {
    const model = doc.value;
    if (!model || !start || !end) {
      selectionActive.value = false;
      lastSelection = null;
      currentLanguageTag.value = model?.defaultStyles.language ?? DEFAULT_LANGUAGE;
      return;
    }
    const selection: ModelSelection = { anchor: start, focus: end };
    lastSelection = selection;
    if (isCollapsed(selection)) {
      // AC-15: the document default, not the caret run's own language —
      // matches the AC's literal "selection language when text is
      // selected, document default otherwise".
      selectionActive.value = false;
      currentLanguageTag.value = model.defaultStyles.language ?? DEFAULT_LANGUAGE;
      return;
    }
    selectionActive.value = true;
    const paragraph = paragraphAt(model, start);
    const formatting = paragraph ? formattingAtOffset(paragraph, start.offset) : undefined;
    currentLanguageTag.value = effectiveLanguage(formatting?.lang, model.defaultStyles.language);
    selectionNoProof.value = formatting?.noProof === true;
  }

  /** The status bar control's text (SPEC-035 AC-15). */
  const statusBarLanguageLabel = computed(() => languageLabel(currentLanguageTag.value));

  function openLanguageDialog() {
    languageDialogOpen.value = true;
  }

  function closeLanguageDialog() {
    languageDialogOpen.value = false;
  }

  /** Applies `tag` to the active selection (AC-9.1), else the whole
   * document's default (AC-9). */
  function setLanguage(tag: string) {
    const model = doc.value;
    if (!model) {
      return;
    }
    const selection = lastSelection ?? caretSelection({ section: 0, block: 0, offset: 0 });
    if (selectionActive.value && !isCollapsed(selection)) {
      applyEdit(applyFormat(model, selection, { lang: tag }), selection, "set language");
    } else {
      applyEdit(setDocumentLanguage(model, selection, tag), selection, "set document language");
    }
    currentLanguageTag.value = tag;
  }

  /** "Do not check spelling" (SPEC-035 AC-11): a selection-only attribute —
   * there is no UI for a document-wide default, since always skipping the
   * whole document would defeat the feature. A no-op without a real
   * selection. */
  function setNoProof(value: boolean) {
    const model = doc.value;
    const selection = lastSelection;
    if (!model || !selection || !selectionActive.value || isCollapsed(selection)) {
      return;
    }
    applyEdit(applyFormat(model, selection, { noProof: value }), selection, "do not check spelling");
    selectionNoProof.value = value;
  }

  return {
    languageDialogOpen,
    currentLanguageTag,
    statusBarLanguageLabel,
    selectionActive,
    selectionNoProof,
    setSelection,
    openLanguageDialog,
    closeLanguageDialog,
    setLanguage,
    setNoProof,
  };
}
