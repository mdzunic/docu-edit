// The workbook view's conditional formats (SPEC-129 "The evaluator"): one
// lib evaluator (cfEvaluator.ts) per open sheet, asked only for the cells
// the grid mounts (`resultAt`, which SheetCellLayer's cells read through
// useSheetCells), and invalidated
// - from every recalculation pass's result — the keys it was started on and
//   those whose value it changed (useRecalc's onResult) — so a format's
//   statistics are recomputed only when a changed cell lies in its ranges;
// - with "all" on undo and redo (the snapshot may restore any rules and any
//   values), structural edits (useStructure calls `invalidate("all")`), rule
//   commands (the manager's steps, labelled CONDITIONAL_FORMATTING_LABEL) and
//   sheet switches.
// `revision` bumps on every invalidation; `resultAt` reads it, so the cell
// layer's computed re-renders. The evaluators read the model as stored —
// values in the workbook's settings and the user's number locale — and
// cache nothing in it; a new workbook or locale starts fresh evaluators.
import { ref, watch, type Ref, type ShallowRef } from "vue";
import { DEFAULT_THEME } from "../../lib/core/theme";
import { createCfEvaluator, type CfEvaluator } from "../../lib/spreadsheet/cfEvaluator";
import { CONDITIONAL_FORMATTING_LABEL, sheetFormats, type CfCellResult } from "../../lib/spreadsheet/conditionalFormats";
import type { CellKey } from "../../lib/spreadsheet/formula/dependencies";
import { storedValueContext } from "../../lib/spreadsheet/formula/evaluate";
import type { Sheet, SpreadsheetModel } from "../../lib/spreadsheet/model";
import type { SpreadsheetLocale } from "../../lib/spreadsheet/numberFormat/locale";
import type { Recalc } from "./useRecalc";
import type { SpreadsheetUndo } from "./useSpreadsheetUndo";

export interface ConditionalFormatsDeps {
  book: { workbook: ShallowRef<SpreadsheetModel | null>; activeSheet: Readonly<Ref<Sheet | null>> };
  recalc: Pick<Recalc, "onResult">;
  undo: Pick<SpreadsheetUndo, "onStep">;
  /** The user's number locale (text coerces to numbers in it); absent =
   * en-US. */
  locale?: () => SpreadsheetLocale;
}

export function useConditionalFormats(deps: ConditionalFormatsDeps) {
  /** Bumped by every invalidation; read by `resultAt`. */
  const revision = ref(0);
  let evaluators = new Map<string, CfEvaluator>();
  let builtFor: { model: SpreadsheetModel; locale: SpreadsheetLocale | undefined } | null = null;

  /** The open sheet's evaluator, created on first use. */
  function evaluatorFor(model: SpreadsheetModel, sheetId: string): CfEvaluator {
    const locale = deps.locale?.();
    if (builtFor?.model !== model || builtFor.locale !== locale) {
      evaluators = new Map();
      builtFor = { model, locale };
    }
    let evaluator = evaluators.get(sheetId);
    if (!evaluator) {
      evaluator = createCfEvaluator(model, sheetId, {
        base: storedValueContext(model, locale ? { locale } : {}),
        theme: model.theme ?? DEFAULT_THEME,
      });
      evaluators.set(sheetId, evaluator);
    }
    return evaluator;
  }

  /** Invalidates every open sheet's evaluator; "all" also forgets the
   * evaluators of sheets the workbook no longer has. */
  function invalidate(changed: CellKey[] | "all"): void {
    if (changed === "all") {
      const ids = new Set(deps.book.workbook.value?.sheets.map((sheet) => sheet.id) ?? []);
      for (const id of evaluators.keys()) {
        if (!ids.has(id)) {
          evaluators.delete(id);
        }
      }
    }
    for (const evaluator of evaluators.values()) {
      evaluator.invalidate(changed);
    }
    revision.value += 1;
  }

  deps.recalc.onResult((changed) => invalidate(changed));
  deps.undo.onStep((kind, label) => {
    if (kind !== "commit" || label === CONDITIONAL_FORMATTING_LABEL) {
      invalidate("all");
    }
  });
  watch(
    () => deps.book.activeSheet.value?.id,
    () => invalidate("all"),
  );

  /** The conditional-format result of a cell on the sheet on show. */
  function resultAt(row: number, col: number): CfCellResult | undefined {
    void revision.value;
    const model = deps.book.workbook.value;
    const sheet = deps.book.activeSheet.value;
    if (!model || !sheet || sheetFormats(sheet).length === 0) {
      return undefined;
    }
    return evaluatorFor(model, sheet.id).resultFor(row, col);
  }

  return { resultAt, invalidate, revision };
}

export type ConditionalFormatsView = ReturnType<typeof useConditionalFormats>;
