// The supported ceiling (SPEC-047): a document past it opens read-only
// rather than the editor pretending it can hold arbitrarily large content.
// Extracted from DocumentView (SPEC-114 needed its wiring lines).
//
// Counted once per document, at the door ("counted after parse and before
// pagination" — the spec's own definition), NOT re-checked on every edit: a
// live re-check would permanently lock a document the moment one paste
// carried it over the limit, with undo — the way back under — disabled by
// the very same flag. A document that grows past the ceiling mid-session
// stays editable until it is next opened; the door check catches it then.
// (Evaluating per-document rather than per-edit also keeps the O(blocks)
// count walk off the keystroke path.)
import { computed, ref, watch, type ShallowRef } from "vue";
import type { DocumentModel } from "../../lib/documentModel";
import { ceilingWarningMessage, checkSupportedCeiling, withinCeiling } from "../../lib/supportedCeiling";

export function useSupportedCeiling(doc: ShallowRef<DocumentModel | null>) {
  const ceilingResult = ref(withinCeiling());
  watch(
    doc,
    (model, prev) => {
      if (!model) {
        ceilingResult.value = withinCeiling();
        return;
      }
      if (prev && prev.id === model.id) {
        return; // the same document being edited — the door check already ran
      }
      ceilingResult.value = checkSupportedCeiling(model);
    },
    { immediate: true },
  );
  return {
    pastSupportedCeiling: computed(() => ceilingResult.value.exceeded),
    ceilingWarning: computed(() => ceilingWarningMessage(ceilingResult.value)),
  };
}
