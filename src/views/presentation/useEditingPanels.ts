// The below-stage panel controllers, wired together (SPEC-094's size-budget
// extraction — the stage-fit code the spec names was already extracted into
// SlideStageArea by SPEC-089, so this equivalent extraction pays for the
// Design-panel wiring): transitions (SPEC-085), animations (SPEC-086), tables
// (SPEC-091) and charts (SPEC-092) all take the same record/commit trio plus
// the stage selection, and their two cross-hookups (the table panel's editing
// cell anchor, the chart double-click) belong with them, not in the view.
// Speaker notes (SPEC-083), the first below-stage panel, joined them when
// SPEC-103's saving wiring needed the view's lines.
import { watch, type ComputedRef, type Ref } from "vue";
import type { PresentationModel, Slide } from "../../lib/presentation/model";
import { useTransitions } from "./useTransitions";
import { useAnimations } from "./useAnimations";
import { useTables } from "./useTables";
import { useSlideCharts } from "./useSlideCharts";
import { useSlideMedia } from "./useSlideMedia";
import { useNotes } from "./useNotes";
import type { usePresentationEditing } from "./usePresentationEditing";

export interface EditingPanelsDeps {
  model: Ref<PresentationModel | null>;
  currentSlide: ComputedRef<Slide | null>;
  currentIndex: Ref<number>;
  record: (label?: string) => void;
  commit: () => void;
  /** Marks `model.media` changed so `commit` re-slices it (SPEC-095). */
  bumpMedia: () => void;
  reportError: (message: string) => void;
  stageEdit: ReturnType<typeof usePresentationEditing>;
}

/** Wires the notes/transition/animation/table/chart panel controllers and
 * their stage hookups; the view mounts the returned controllers unchanged. */
export function useEditingPanels(deps: EditingPanelsDeps) {
  const { model, currentSlide, currentIndex, record, commit, bumpMedia, reportError, stageEdit } = deps;

  const notes = useNotes({ model, currentSlide, currentIndex, record, commit });

  const transitions = useTransitions({ model, currentSlide, currentIndex, record, commit });

  const media = useSlideMedia({
    model,
    currentSlide,
    currentIndex,
    record,
    commit,
    bumpMedia,
    reportError,
    selectedFrameId: stageEdit.selection.selectedFrameId,
    selectFrame: stageEdit.selection.select,
  });

  const animations = useAnimations({
    model,
    currentSlide,
    currentIndex,
    record,
    commit,
    selectedFrameId: stageEdit.selection.selectedFrameId,
    selectFrame: stageEdit.selection.select,
  });

  const tables = useTables({
    model,
    currentSlide,
    currentIndex,
    record,
    commit,
    selectedFrame: stageEdit.selection.selectedFrame,
    selectFrame: stageEdit.selection.select,
  });
  // The table panel anchors on the cell being edited.
  watch(stageEdit.editing.activeCell, (cell) => cell && tables.setCurrentCell(cell.row, cell.col));

  const charts = useSlideCharts({
    model,
    currentSlide,
    currentIndex,
    record,
    commit,
    selectedFrame: stageEdit.selection.selectedFrame,
    selectFrame: stageEdit.selection.select,
  });
  stageEdit.setChartEditHandler(charts.openDataSheet); // double-click a chart → data sheet

  return { notes, transitions, animations, tables, charts, media };
}
