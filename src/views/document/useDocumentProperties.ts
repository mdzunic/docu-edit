// The "File → Properties" dialog (SPEC-041): owns its open/closed state,
// the async file-size computation the General tab shows (AC-2), and the two
// commit handlers PropertiesDialog.vue's `update-properties`/
// `update-author-identity` events call. `doc` is DocumentView's shallowRef —
// updates mutate its current value in place and `triggerRef` it, the same
// "mutate the stored model directly" contract `handleTitleUpdate` already
// uses, so the change is persisted (via `touchDocument`) exactly like every
// other direct edit.
import { ref, shallowRef, triggerRef } from "vue";
import type { Ref, ShallowRef } from "vue";
import type { DocumentModel, DocumentProperties } from "../../lib/documentModel";
import { writeDocx } from "../../lib/docx/writer";
import { writeOdt } from "../../lib/odf/writer";

export interface DocumentPropertiesDeps {
  doc: ShallowRef<DocumentModel | null>;
  touchDocument: (id: string) => void;
  /** The settable identity (lib/commentIdentity.ts), as DocumentView's own
   * refs for the Comments/Track Changes panes — updating the identity here
   * updates both, so all three surfaces stay in lockstep for the rest of
   * the session (SPEC-041 AC-15). */
  commentAuthorName: Ref<string>;
  trackChangesAuthorName: Ref<string>;
}

export function useDocumentProperties(deps: DocumentPropertiesDeps) {
  const { doc, touchDocument, commentAuthorName, trackChangesAuthorName } = deps;

  const propertiesOpen = ref(false);
  /** Bytes, or null while still computing / for a document with no saved
   * format yet (AC-18a) — PropertiesDialog.vue shows "Calculating…" for the
   * former, a blank field for the latter (the format check below). */
  const propertiesFileSize = shallowRef<number | null>(null);

  /** Serialises the current model the same way a real save would, purely to
   * report its size (AC-2) — nothing here is written anywhere. Recomputed
   * fresh every time the dialog opens so edits made since the last open (or
   * the last real save) are reflected. */
  async function openProperties() {
    propertiesOpen.value = true;
    propertiesFileSize.value = null;
    const model = doc.value;
    if (!model || !model.format) {
      return;
    }
    const result = model.format === "odt" ? await writeOdt(model) : await writeDocx(model);
    // The dialog may have been closed, or a different document opened,
    // while the (async) write was in flight — stale results are dropped.
    if (propertiesOpen.value && doc.value === model) {
      propertiesFileSize.value = result.bytes.length;
    }
  }

  function closeProperties() {
    propertiesOpen.value = false;
  }

  function updateDocumentProperties(next: DocumentProperties) {
    const model = doc.value;
    if (!model) {
      return;
    }
    model.properties = next;
    triggerRef(doc);
    touchDocument(model.id);
  }

  function updateAuthorIdentity(name: string) {
    commentAuthorName.value = name;
    trackChangesAuthorName.value = name;
  }

  return {
    propertiesOpen,
    propertiesFileSize,
    openProperties,
    closeProperties,
    updateDocumentProperties,
    updateAuthorIdentity,
  };
}
