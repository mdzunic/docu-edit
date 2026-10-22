// Version history (SPEC-040): the pane's open state, the snapshot cadence
// and commit-point/restore triggers, and the pane's actions (preview with
// diff highlighting, restore, name/unname, delete, download). The storage
// and diff/patch logic lives in lib/versionHistory.ts; this composable is
// the Vue-facing session state, following useTrackChanges.ts's shape.
import { computed, ref } from "vue";
import {
  deleteVersion as deleteVersionRecord,
  loadIndex,
  nameVersion as nameVersionRecord,
  reconstructVersion,
  recordVersion,
  removeVersionName as removeVersionNameRecord,
  versionListEntries,
  HISTORY_PAUSED_MESSAGE,
  type SnapshotTrigger,
  type VersionListEntry,
  type VersionsIndexRecord,
} from "../../lib/versionHistory";
import { createVersionCadenceTimer, type VersionCadenceTimer } from "../../lib/core/versionCadenceScheduler";
import { usePreferences } from "../../composables/usePreferences";
import { applyTrackChanges } from "../../lib/trackChangesDiff";
import { caretSelection } from "../../lib/commands";
import type { DocumentModel } from "../../lib/documentModel";
import { DOCX_MIME_TYPE, docxFileName, writeDocx } from "../../lib/docx/writer";
import { ODT_MIMETYPE, odtFileName, writeOdt } from "../../lib/odf/writer";
import { triggerBlobDownload } from "../../lib/core/triggerDownload";
import type { EditorContext } from "./context";

/** Typed pass-through for a `VersionListEntry[]` arriving at
 * VersionHistoryPane.vue as a runtime-declared Array prop (espree bars the
 * annotation there — see useTrackChanges.ts's asChangeEntries). */
export function asVersionListEntries(value: unknown): VersionListEntry[] {
  return value as VersionListEntry[];
}

export type UseVersionHistoryDeps = EditorContext;

export function useVersionHistory(deps: UseVersionHistoryDeps) {
  const { doc, documentId, applyEdit, getCaret, onSessionClose } = deps;

  // --- pane open state (per-user, like the other docks) -------------------

  // Open/closed state. Since SPEC-065 this is a tab in the consolidated side
  // panel, which owns the persistence (one key, migrated from the legacy
  // docu-edit:version-history-pane key) and drives this ref through useSidePanel.
  const versionHistoryPaneOpen = ref(false);
  function toggleVersionHistoryPane(): void {
    versionHistoryPaneOpen.value = !versionHistoryPaneOpen.value;
  }

  // --- index (the pane's list; small, index-only per AC-43) --------------

  const indexRecord = ref<VersionsIndexRecord>({ id: "", store: "index", docId: "", versions: [], paused: false });
  const versions = computed<VersionListEntry[]>(() => versionListEntries(indexRecord.value));
  const pausedMessage = computed(() => (indexRecord.value.paused ? HISTORY_PAUSED_MESSAGE : null));
  const namedOnlyFilter = ref(false);
  const visibleVersions = computed(() =>
    namedOnlyFilter.value ? versions.value.filter((v) => !!v.name) : versions.value,
  );

  async function refreshIndex(): Promise<void> {
    if (!documentId.value) {
      return;
    }
    indexRecord.value = await loadIndex(documentId.value);
  }

  // --- cadence + "changed since last version" (AC-1) -----------------------

  let cachedPrev: { id: string; model: DocumentModel } | null = null;
  let changedSinceLastVersion = false;

  // The version author is the SPEC-038 identity, now a centralized preference
  // (SPEC-066) — read live so a rename is reflected without a reload.
  const versionHistoryPreferences = usePreferences();
  function authorName(): string {
    return versionHistoryPreferences.preferences.authorName;
  }

  async function attemptSnapshot(trigger: SnapshotTrigger): Promise<void> {
    const model = doc.value;
    if (!model || !documentId.value || previewActive.value) {
      return;
    }
    // "Keep version history" off stops the periodic recording (SPEC-066);
    // the snapshots that protect a restore still run so undoing a restore
    // never loses the pre-restore state.
    if (
      !versionHistoryPreferences.preferences.keepVersionHistory &&
      (trigger === "cadence" || trigger === "commit")
    ) {
      return;
    }
    const result = await recordVersion({
      docId: documentId.value,
      model,
      author: authorName(),
      trigger,
      cachedPrev,
    });
    indexRecord.value = result.index;
    if (result.newEntry) {
      cachedPrev = result.newEntry;
      changedSinceLastVersion = false;
    }
  }

  /** Called wherever autosave.markDirty() is (see DocumentView.vue's
   * noteEdit) — the same activity signal the cadence timer runs off, one
   * order of magnitude coarser (versionCadenceScheduler.ts's header). */
  function noteChange(): void {
    changedSinceLastVersion = true;
    cadenceTimer.noteChange();
  }

  const cadenceTimer: VersionCadenceTimer = createVersionCadenceTimer({
    onDue: () => {
      if (changedSinceLastVersion) {
        void attemptSnapshot("cadence");
      }
    },
  });

  /** SPEC-010's download commit point (AC-2); DocumentView calls this right
   * after autosave.commitNow(). */
  function noteCommitPoint(): void {
    if (changedSinceLastVersion) {
      void attemptSnapshot("commit");
    }
  }

  // --- preview: select a version, read-only in the canvas with diff
  // highlighting (AC-11..13) --------------------------------------------
  //
  // `doc.value` is swapped directly (not through replaceDocumentModel/
  // applyEdit), so nothing here ever touches the committed record in the
  // store — the live document is simply stashed and restored around the
  // swap, exactly like switching to a different open document.

  const previewVersionId = ref<string | null>(null);
  const previewActive = computed(() => previewVersionId.value !== null);
  let stashedLiveModel: DocumentModel | null = null;

  async function selectVersion(versionId: string): Promise<void> {
    const docId = documentId.value;
    const chronological = indexRecord.value.versions;
    const idx = chronological.findIndex((v) => v.id === versionId);
    if (idx === -1) {
      return;
    }
    const target = await reconstructVersion(docId, versionId);
    if (!target) {
      return;
    }
    const previousEntry = idx > 0 ? chronological[idx - 1] : null;
    const previous = previousEntry ? await reconstructVersion(docId, previousEntry.id) : null;
    const entry = chronological[idx];
    const markedUp = previous
      ? applyTrackChanges(previous, target, entry.author, new Date(entry.timestamp).toISOString())
      : target;
    if (!previewActive.value) {
      stashedLiveModel = doc.value;
    }
    previewVersionId.value = versionId;
    doc.value = { ...markedUp, id: docId };
  }

  function exitPreview(): void {
    if (stashedLiveModel) {
      doc.value = stashedLiveModel;
      stashedLiveModel = null;
    }
    previewVersionId.value = null;
  }

  // --- restore (AC-4, AC-5, AC-14..16) --------------------------------------

  async function restoreVersion(versionId: string): Promise<void> {
    exitPreview();
    const current = doc.value;
    const docId = documentId.value;
    if (!current || !docId) {
      return;
    }
    await attemptSnapshot("pre-restore");
    const target = await reconstructVersion(docId, versionId);
    if (!target) {
      return;
    }
    const restored: DocumentModel = { ...target, id: current.id };
    const at = getCaret() ?? { section: 0, block: 0, offset: 0 };
    const selection = caretSelection(at);
    applyEdit({ model: restored, selection }, selection, "restore version");
    await attemptSnapshot("post-restore");
  }

  // --- name / unname / delete (AC-17, AC-18, AC-38's escape hatches) ------

  async function nameVersion(versionId: string, name: string): Promise<void> {
    indexRecord.value = await nameVersionRecord(documentId.value, versionId, name);
  }

  async function removeVersionName(versionId: string): Promise<void> {
    const model = doc.value;
    if (!model) {
      return;
    }
    indexRecord.value = await removeVersionNameRecord(documentId.value, versionId, model);
  }

  async function deleteVersion(versionId: string): Promise<void> {
    const model = doc.value;
    if (!model) {
      return;
    }
    if (previewVersionId.value === versionId) {
      exitPreview();
    }
    indexRecord.value = await deleteVersionRecord(documentId.value, versionId, model);
  }

  // --- download (AC-19..21) -------------------------------------------------

  async function downloadVersion(versionId: string, format: "docx" | "odt"): Promise<void> {
    const model = await reconstructVersion(documentId.value, versionId);
    if (!model) {
      return;
    }
    const odt = format === "odt";
    const { bytes } = odt ? await writeOdt(model) : await writeDocx(model);
    const fileName = odt ? odtFileName(model.title) : docxFileName(model.title);
    triggerBlobDownload(bytes, fileName, odt ? ODT_MIMETYPE : DOCX_MIME_TYPE);
  }

  // --- session lifecycle -----------------------------------------------------

  function beginVersionHistorySession(): void {
    cachedPrev = null;
    changedSinceLastVersion = false;
    previewVersionId.value = null;
    stashedLiveModel = null;
    namedOnlyFilter.value = false;
    void refreshIndex();
  }

  onSessionClose(() => {
    cachedPrev = null;
    changedSinceLastVersion = false;
    previewVersionId.value = null;
    stashedLiveModel = null;
    cadenceTimer.reset();
    indexRecord.value = { id: "", store: "index", docId: "", versions: [], paused: false };
  });

  return {
    versionHistoryPaneOpen,
    toggleVersionHistoryPane,
    versions: visibleVersions,
    pausedMessage,
    namedOnlyFilter,
    previewActive,
    previewVersionId,
    beginVersionHistorySession,
    noteChange,
    noteCommitPoint,
    selectVersion,
    exitPreview,
    restoreVersion,
    nameVersion,
    removeVersionName,
    deleteVersion,
    downloadVersion,
  };
}

export type UseVersionHistoryReturn = ReturnType<typeof useVersionHistory>;
