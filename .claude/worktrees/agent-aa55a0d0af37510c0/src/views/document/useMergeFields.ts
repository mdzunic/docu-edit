// Mail-merge fields (SPEC-046): the attached data source, the "Insert →
// Merge field…" picker and merge-run dialogs' open state, the named field
// lists both they and "Table → Repeat row for each…" render, and the
// convert-literal-{{…}}-text command. The merge run itself lives in
// MergeRunDialog + lib/merge; the repeat-row confirm stays with useTables
// (it is a table command).
import { computed, ref, shallowRef } from "vue";
import { insertMergeChip } from "../../lib/commands";
import { namedArrayFieldList, namedFieldList, noParsedDataSource } from "../../lib/merge";
import { convertLiteralChipsToRuns, hasConvertibleLiteralChips } from "../../lib/mergeChips";
import type { ModelSelection } from "../../lib/commands";
import type { EditorContext } from "./context";

export interface MergeFieldsDeps extends EditorContext {
  /** DocumentView's paragraphSelection always resolves to a selection. */
  paragraphSelection: (root?: HTMLElement | null) => ModelSelection;
}

export function useMergeFields(deps: MergeFieldsDeps) {
  const { doc, contentRoot, applyEdit, addToast, paragraphSelection } = deps;

  // --- mail-merge chip insertion (SPEC-046 AC-12, AC-15, AC-30) ---------------
  // The attached data source (if any) is shared between "Insert -> Merge
  // field...", "Table -> Repeat row for each..." and the merge run dialog —
  // attaching it once populates every named field list, and running the merge
  // re-parses the same file rather than trusting stale state across a session.
  const mergeDataSource = shallowRef(noParsedDataSource());
  const mergeFieldPickerOpen = ref(false);
  const mergeRunDialogOpen = ref(false);

  function attachMergeDataSource(parsed = noParsedDataSource()) {
    mergeDataSource.value = parsed;
  }
  const mergeFieldPickerFields = computed(() => (doc.value ? namedFieldList(mergeDataSource.value, doc.value) : []));
  const repeatRowFieldOptions = computed(() => (doc.value ? namedArrayFieldList(mergeDataSource.value, doc.value) : []));

  function insertMergeFieldFromPicker(name = "") {
    mergeFieldPickerOpen.value = false;
    const root = contentRoot.value as HTMLElement | null;
    const model = doc.value;
    if (!name || !root || !model) {
      return;
    }
    const selection = paragraphSelection(root);
    const result = insertMergeChip(model, selection, { kind: "merge", name });
    applyEdit(result, selection, "insert merge field");
  }

  // --- mail-merge: convert literal {{...}} text (SPEC-046) --------------------

  /** Whether "Convert {{…}} text to merge fields" has anything to do — offered
   * when a document holds at least one well-formed `{{…}}` literal outside an
   * existing chip (AC-20, "Literal text is not a chip"). */
  const canConvertLiteralChips = computed(() => (doc.value ? hasConvertibleLiteralChips(doc.value) : false));

  /** Scans the whole document and converts every literal `{{…}}` match into a
   * chip in one undo step — explicit and undoable, never automatic on open
   * (SPEC-046 "Literal text is not a chip"). */
  function convertLiteralChips() {
    const model = doc.value;
    const root = contentRoot.value as HTMLElement | null;
    if (!model || !root) {
      return;
    }
    const selection = paragraphSelection(root);
    const { model: next, converted } = convertLiteralChipsToRuns(model);
    if (converted === 0) {
      return;
    }
    applyEdit({ model: next, selection }, selection, "convert merge fields");
    addToast(converted === 1 ? "Converted 1 merge field" : `Converted ${converted} merge fields`);
  }

  return {
    mergeDataSource,
    mergeFieldPickerOpen,
    mergeRunDialogOpen,
    attachMergeDataSource,
    mergeFieldPickerFields,
    repeatRowFieldOptions,
    insertMergeFieldFromPicker,
    canConvertLiteralChips,
    convertLiteralChips,
  };
}
