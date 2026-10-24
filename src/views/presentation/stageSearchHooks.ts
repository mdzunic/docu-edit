// The stage hooks a deck search drives (SPEC-098), built from the editing and
// notes controllers so PresentationView wires the search in one call: select
// a hit's frame, open a text session on it (or its table cell) with the match
// selected through slideTextDom's setSelectionRange, open the notes panel and
// select a notes hit in its textarea, end a running session before a replace,
// and hand focus back to the stage when the bar closes.
import { nextTick, type Ref } from "vue";
import type { PresentationModel } from "../../lib/presentation/model";
import type { DeckHit } from "../../lib/presentation/deckSearch";
import type { usePresentationEditing } from "./usePresentationEditing";

export interface StageSearchHooks {
  selectFrame: (frameId: string) => void;
  startEditing: (hit: DeckHit) => void;
  openNotes: (hit: DeckHit) => Promise<void> | void;
  stopEditing: () => void;
  editingFrameId: () => string | null;
  surface: () => HTMLElement | null;
  focusStage: () => void;
}

/** The offset of a notes line inside the textarea's newline-joined value. */
export function notesOffset(lines: readonly string[], para: number, offset: number): number {
  return lines.slice(0, para).reduce((total, line) => total + line.length + 1, 0) + offset;
}

export function stageSearchHooks(
  stageEdit: Pick<ReturnType<typeof usePresentationEditing>, "selection" | "editing" | "surface">,
  notes: { open: { readonly value: boolean }; toggle: () => void },
  model: Ref<PresentationModel | null>,
): StageSearchHooks {
  const { selection, editing, surface } = stageEdit;
  return {
    selectFrame: (frameId) => selection.select(frameId),
    startEditing: (hit) => {
      if (!hit.frameId) return;
      const range = { start: { para: hit.para, offset: hit.start }, end: { para: hit.para, offset: hit.end } };
      // A shape's text has no session (SPEC-081 edits text frames and cells):
      // it stays selected, its match painted by the proofing layer.
      if (hit.scope === "cell") editing.startEditingCell(hit.frameId, hit.row ?? 0, hit.column ?? 0, undefined, range);
      else editing.startEditing(hit.frameId, false, undefined, range);
    },
    openNotes: async (hit) => {
      if (!notes.open.value) notes.toggle();
      await nextTick();
      const area = document.querySelector<HTMLTextAreaElement>('[data-testid="slides-notes-editor"]');
      const lines = model.value?.slides[hit.slideIndex]?.notes ?? [];
      area?.focus();
      area?.setSelectionRange(notesOffset(lines, hit.para, hit.start), notesOffset(lines, hit.para, hit.end));
    },
    stopEditing: () => {
      if (editing.editingFrameId.value) editing.stopEditing();
    },
    editingFrameId: () => editing.editingFrameId.value,
    surface,
    focusStage: () => {
      // A session the bar opened ends here: its host is not where focus goes,
      // and the stage's keys stand aside while a session is open.
      if (editing.editingFrameId.value) editing.stopEditing();
      const selected = document.querySelector<HTMLElement>('[data-testid="slide-frame"].is-selected');
      (selected ?? document.querySelector<HTMLElement>('[data-testid="stage-workspace"]'))?.focus();
    },
  };
}
