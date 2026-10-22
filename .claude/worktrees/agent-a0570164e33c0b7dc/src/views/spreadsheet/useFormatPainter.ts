// The sheet's format painter (SPEC-123 "Format painter", SPEC-020's
// vocabulary): a click on `sheet-format-painter` arms it for one application
// from the current selection; a double-click locks it (aria-pressed and
// data-locked="true") until Escape or another click on it. While armed the
// grid carries data-format-painter="once|locked" and the copy cursor. The
// next COMPLETED selection — a pointer release in the grid, or Enter after a
// keyboard selection — receives the source's pattern (formatting.ts's
// capturePattern/applyPattern), ONE "Format painter" undo step per
// application, across sheets if the user switches tabs (style ids are the
// workbook's). A multi-range source disables the painter; a source over
// MAX_FORMAT_MATERIALISED positions is refused with FORMAT_TOO_LARGE_MESSAGE.
// SPEC-124: the source's merges are reproduced in the target, in the same
// step — after useMerges' warning when that clears a value or formula.
import { computed, onScopeDispose, ref, shallowRef, watch, type Ref } from "vue";
import { useToasts } from "../../composables/useToasts";
import { applyPattern, capturePattern, FORMAT_TOO_LARGE_MESSAGE, type StylePattern } from "../../lib/spreadsheet/formatting";
import type { Sheet } from "../../lib/spreadsheet/model";
import { applyPatternMerges, patternMerges, targetsDiscardValues } from "../../lib/spreadsheet/mergeCells";
import type { Formatting } from "./useFormatting";
import type { Merges } from "./useMerges";
import type { SheetSelectionController } from "./useSheetSelection";

export type PainterMode = "" | "once" | "locked";

export interface FormatPainterDeps {
  book: { activeSheet: Readonly<Ref<Sheet | null>> };
  selection: Pick<SheetSelectionController, "selection" | "drag">;
  formatting: Pick<Formatting, "command">;
  /** SPEC-124: the merge warning, before painted merges clear a value. */
  merges?: Pick<Merges, "confirmDiscard">;
  toast?: (message: string) => void;
}

export function useFormatPainter(deps: FormatPainterDeps) {
  const toast = deps.toast ?? useToasts().addToast;
  const mode = ref<PainterMode>("");
  const pattern = shallowRef<StylePattern | null>(null);

  /** The painter cannot take a multi-range source. */
  const disabled = computed(() => deps.selection.selection.value.ranges.length !== 1 || deps.book.activeSheet.value?.kind !== "grid");

  function disarm(): void {
    mode.value = "";
    pattern.value = null;
  }

  function arm(next: "once" | "locked"): void {
    const sheet = deps.book.activeSheet.value;
    const range = deps.selection.selection.value.ranges[0];
    if (!sheet || !range || disabled.value) {
      return;
    }
    const captured = capturePattern(sheet, range);
    if (!captured) {
      toast(FORMAT_TOO_LARGE_MESSAGE);
      return;
    }
    pattern.value = captured;
    mode.value = next;
  }

  /** Paints the pattern onto every range of the selection just made. */
  function paint(): void {
    const source = pattern.value;
    if (mode.value === "" || !source) {
      return;
    }
    if (mode.value === "once") {
      disarm();
    }
    const run = (): void =>
      deps.formatting.command("Format painter", (model, sheetId, ranges) => {
        for (const range of ranges) {
          const result = applyPattern(model, sheetId, range, source);
          if (!result.ok) {
            return result;
          }
          // SPEC-124: the source's merges, reproduced in the target.
          const merged = applyPatternMerges(model, sheetId, range, source);
          if (!merged.ok) {
            return merged;
          }
        }
        return { ok: true };
      });
    const sheet = deps.book.activeSheet.value;
    const discards = source.merges !== undefined && sheet !== null && deps.selection.selection.value.ranges.some((range) => targetsDiscardValues(sheet, patternMerges(range, source)));
    if (deps.merges) {
      deps.merges.confirmDiscard(discards, run);
    } else {
      run();
    }
  }

  // A pointer selection completes when its drag is released.
  watch(
    () => deps.selection.drag.value,
    (now, before) => {
      if (before && !now) {
        paint();
      }
    },
  );

  function onKeydown(event: KeyboardEvent): void {
    if (mode.value === "" || event.defaultPrevented) {
      return;
    }
    if (event.key === "Escape") {
      event.preventDefault();
      disarm();
      return;
    }
    const inGrid = event.target instanceof Element && event.target.closest('[data-testid="spreadsheet-grid"]') !== null;
    if (event.key === "Enter" && inGrid && !event.shiftKey && !event.ctrlKey && !event.metaKey && !event.altKey) {
      // Enter completes a keyboard selection instead of moving on.
      event.preventDefault();
      event.stopPropagation();
      paint();
    }
  }

  if (typeof window !== "undefined") {
    window.addEventListener("keydown", onKeydown, true);
  }
  onScopeDispose(() => {
    if (typeof window !== "undefined") {
      window.removeEventListener("keydown", onKeydown, true);
    }
  });

  return {
    mode,
    disabled,
    /** A click arms one application, or disarms. */
    click: (): void => (mode.value === "" ? arm("once") : disarm()),
    /** A double-click locks it on. */
    dblclick: (): void => arm("locked"),
    disarm,
    paint,
    onKeydown,
  };
}

export type FormatPainter = ReturnType<typeof useFormatPainter>;

/** Narrows a component's runtime-typed `painter` prop. */
export function asFormatPainter(value: unknown): FormatPainter {
  return value as FormatPainter;
}
