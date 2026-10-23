// The deck's version history controller (SPEC-103) — SPEC-040's pane for the
// presentation editor: the panel's open state (the `presentationHistoryOpen`
// preference), the snapshot triggers (createVersionCadenceTimer's five
// minutes of editing activity, the commit points, and the pair bracketing a
// restore), and the panel's actions — preview, restore, name/unname, delete,
// download. Storage and the chain live in lib/presentationVersions.ts;
// useDeckSaving composes this with the deck's autosave and hands the view one
// controller. Nothing here is awaited on an edit's path, and every failure
// stays inside history: the deck's own records are never touched.
import { computed, onBeforeUnmount, ref, shallowRef, type Ref, type ShallowRef } from "vue";
import type { PresentationModel } from "../../lib/presentation/model";
import { replacePresentationModel } from "../../lib/documentStore";
import { emptyVersionsIndex, HISTORY_PAUSED_MESSAGE, type SnapshotTrigger, type VersionsIndexRecord } from "../../lib/core/versionChain";
import { createVersionCadenceTimer } from "../../lib/core/versionCadenceScheduler";
import {
  deckVersionEntries,
  deleteDeckVersion,
  loadDeckIndex,
  nameDeckVersion,
  reconstructDeckVersion,
  recordDeckVersion,
  removeDeckVersionName,
  DECK_HISTORY_TOO_LARGE_MESSAGE,
  type DeckVersionTail,
} from "../../lib/presentationVersions";
import { downloadPresentation } from "../../lib/presentation/download";
import { usePreferences } from "../../composables/usePreferences";

/** The undo label a restore records (SPEC-080's stack). */
export const RESTORE_UNDO_LABEL = "Restore version";

/** Said when a version cannot be read back (a missing blob, a broken chain). */
export const VERSION_UNREADABLE_MESSAGE = "That version couldn't be read";

export interface DeckHistoryDeps {
  model: ShallowRef<PresentationModel | null>;
  currentIndex: Ref<number>;
  /** usePresentationUndo.record — a restore is one undo step. */
  record: (label?: string) => void;
  /** useSlideRender.commit — re-renders the restored deck and reports it. */
  commit: () => void;
  reportError: (message: string) => void;
}

export function useDeckHistory(deps: DeckHistoryDeps) {
  const { preferences, setPreference } = usePreferences();

  // --- the panel ---------------------------------------------------------------
  const open = computed(() => preferences.presentationHistoryOpen);
  function toggle(): void {
    setPreference("presentationHistoryOpen", !open.value);
  }
  /** What a version downloads as — the deck's own format. */
  const format = computed(() => deps.model.value?.format ?? "pptx");

  const index = shallowRef<VersionsIndexRecord>(emptyVersionsIndex(""));
  const versions = computed(() => deckVersionEntries(index.value));
  const namedOnly = ref(false);
  const visibleVersions = computed(() => (namedOnly.value ? versions.value.filter((v) => !!v.name) : versions.value));
  const pausedMessage = computed(() => {
    if (!index.value.paused) {
      return "";
    }
    return index.value.pausedReason === "too-large" ? DECK_HISTORY_TOO_LARGE_MESSAGE : HISTORY_PAUSED_MESSAGE;
  });

  // --- the session ---------------------------------------------------------------
  let docId: string | null = null;
  let cachedPrev: DeckVersionTail | null = null;
  let changedSinceLastVersion = false;

  function begin(id: string): void {
    docId = id;
    cachedPrev = null;
    changedSinceLastVersion = false;
    namedOnly.value = false;
    exitPreview();
    index.value = emptyVersionsIndex(id);
    void loadDeckIndex(id).then((loaded) => {
      if (docId === id) {
        index.value = loaded;
      }
    });
  }

  function end(): void {
    cadence.reset();
    docId = null;
    cachedPrev = null;
    changedSinceLastVersion = false;
    exitPreview();
  }

  // --- snapshots (SPEC-040's cadence and triggers) ----------------------------
  async function snapshot(trigger: SnapshotTrigger): Promise<void> {
    const id = docId;
    const model = deps.model.value;
    if (!id || !model) {
      return;
    }
    // "Keep version history" off stops the periodic and commit-point
    // recording only; a restore's own pair still runs, so undoing a restore
    // can never lose the pre-restore state.
    if (!preferences.keepVersionHistory && (trigger === "cadence" || trigger === "commit")) {
      return;
    }
    const result = await recordDeckVersion({ docId: id, model, author: preferences.authorName, trigger, cachedPrev });
    if (docId !== id) {
      return;
    }
    index.value = result.index;
    if (result.newEntry) {
      cachedPrev = result.newEntry;
      changedSinceLastVersion = false;
    }
  }

  const cadence = createVersionCadenceTimer({
    onDue: () => {
      if (changedSinceLastVersion) {
        void snapshot("cadence");
      }
    },
  });
  onBeforeUnmount(() => cadence.dispose());

  /** An edit happened (useDeckSaving.noteEdit): arms the cadence. */
  function noteChange(): void {
    changedSinceLastVersion = true;
    cadence.noteChange();
  }

  /** A commit point — a download, a clean close: snapshots if the deck
   * changed since its last version. */
  function noteCommitPoint(): void {
    if (changedSinceLastVersion) {
      void snapshot("commit");
    }
  }

  // --- preview: read-only, on its own stage -----------------------------------
  // The reconstructed deck renders in the panel's own SlideStage with its own
  // slide index; the live deck is never swapped, so leaving the preview leaves
  // it exactly as it was. There is no diff highlighting — the text side's
  // applyTrackChanges is DocumentModel-typed.
  const previewModel = shallowRef<PresentationModel | null>(null);
  const previewVersionId = ref<string | null>(null);
  const previewIndex = ref(0);

  async function preview(versionId: string): Promise<void> {
    const id = docId;
    if (!id) {
      return;
    }
    const model = await reconstructDeckVersion(id, versionId, deps.model.value ?? undefined);
    if (docId !== id) {
      return;
    }
    if (!model) {
      deps.reportError(VERSION_UNREADABLE_MESSAGE);
      return;
    }
    previewModel.value = model;
    previewVersionId.value = versionId;
    previewIndex.value = Math.min(deps.currentIndex.value, Math.max(0, model.slides.length - 1));
  }

  function exitPreview(): void {
    previewModel.value = null;
    previewVersionId.value = null;
    previewIndex.value = 0;
  }

  function selectPreviewSlide(slideIndex: number): void {
    const count = previewModel.value?.slides.length ?? 0;
    previewIndex.value = Math.min(Math.max(0, slideIndex), Math.max(0, count - 1));
  }

  // --- restore ---------------------------------------------------------------------
  /** pre-restore snapshot → record("Restore version") → swap → commit() →
   * post-restore snapshot. Ctrl+Z afterwards puts back what the undo
   * snapshot holds (the slides and the deck settings it clones, SPEC-080) —
   * not the package or the title — and the pre-restore version stays in the
   * list, one Restore away. */
  async function restore(versionId: string): Promise<void> {
    exitPreview();
    const id = docId;
    if (!id || !deps.model.value) {
      return;
    }
    await snapshot("pre-restore");
    const target = await reconstructDeckVersion(id, versionId, deps.model.value ?? undefined);
    if (docId !== id || !deps.model.value) {
      return;
    }
    if (!target) {
      deps.reportError(VERSION_UNREADABLE_MESSAGE);
      return;
    }
    deps.record(RESTORE_UNDO_LABEL);
    deps.model.value = replacePresentationModel(id, target);
    deps.commit();
    await snapshot("post-restore");
  }

  // --- name / unname / delete / download ---------------------------------------
  /** Runs one index action, adopting the index it returns; a failure is
   * reported and leaves the list as it was. */
  async function update(action: (id: string, live: PresentationModel) => Promise<VersionsIndexRecord>): Promise<void> {
    const id = docId;
    const live = deps.model.value;
    if (!id || !live) {
      return;
    }
    try {
      const next = await action(id, live);
      if (docId === id) {
        index.value = next;
      }
    } catch {
      deps.reportError("Version history couldn't be updated");
    }
  }

  function nameVersion(versionId: string, name: string): Promise<void> {
    return update((id) => nameDeckVersion(id, versionId, name));
  }

  function removeName(versionId: string): Promise<void> {
    return update((id, live) => removeDeckVersionName(id, versionId, live));
  }

  function deleteVersion(versionId: string): Promise<void> {
    if (previewVersionId.value === versionId) {
      exitPreview();
    }
    return update((id, live) => deleteDeckVersion(id, versionId, live));
  }

  /** Downloads a version in its own format (.pptx/.odp). */
  async function downloadVersion(versionId: string): Promise<void> {
    const id = docId;
    if (!id) {
      return;
    }
    const model = await reconstructDeckVersion(id, versionId, deps.model.value ?? undefined);
    if (!model) {
      deps.reportError(VERSION_UNREADABLE_MESSAGE);
      return;
    }
    await downloadPresentation(model);
  }

  return {
    open,
    toggle,
    format,
    versions,
    visibleVersions,
    namedOnly,
    pausedMessage,
    previewModel,
    previewVersionId,
    previewIndex,
    preview,
    exitPreview,
    selectPreviewSlide,
    restore,
    nameVersion,
    removeName,
    deleteVersion,
    downloadVersion,
    begin,
    end,
    noteChange,
    noteCommitPoint,
  };
}

export type DeckHistory = ReturnType<typeof useDeckHistory>;
