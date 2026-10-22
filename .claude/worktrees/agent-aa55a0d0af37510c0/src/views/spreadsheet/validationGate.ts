// SPEC-130's entry gate, framework-free (useDataValidation holds its state):
// which typed entries the data-validation rules refuse, and what the alert
// for the first refusal says.
//
// - firstValidationFailure: the first of the targets — in the order given,
//   which for Ctrl/Cmd+Enter's fill is the selection's (validatedFillTargets,
//   sparse: only the cells a rule covers) — whose rule refuses the candidate
//   (validation.ts's validateValue, in the user's locale, over stored
//   values), skipping rules whose `error.show` is false (the entry is then
//   committed silently, as in Excel). A formula candidate is checked by its
//   result evaluated at each target (shifted from `origin` there, as the
//   fill writes it); one that cannot be computed is accepted. A plain value
//   under a position-independent rule is checked once per rule.
// - The alert: titled `error.title` or "Invalid entry", reading `error.text`
//   or the default message; its buttons' verdicts are ALERT_VERDICTS — stop:
//   Retry → "retry", Cancel → "cancel"; warning: Yes → "commit", No →
//   "retry", Cancel → "cancel"; information: OK → "commit", Cancel →
//   "cancel" (Escape is Cancel).
//
// Nothing else validates (validation.ts's header): paste, fill, drag-move,
// clearing, undo/redo, sort and CSV import never reach this gate.
import type { CellRef, GridRange } from "../../lib/spreadsheet/address";
import { fillTargets as cellFillTargets } from "../../lib/spreadsheet/cells";
import { storedValueContext } from "../../lib/spreadsheet/formula/evaluate";
import { shiftFormula } from "../../lib/spreadsheet/formula/references";
import { tokenize } from "../../lib/spreadsheet/formula/tokens";
import { buildMergeIndex, mergeAt } from "../../lib/spreadsheet/merges";
import { MAX_COLS, type CellValue, type Sheet, type SpreadsheetModel } from "../../lib/spreadsheet/model";
import type { SpreadsheetLocale } from "../../lib/spreadsheet/numberFormat/locale";
import { workbookFormatContext } from "../../lib/spreadsheet/numberFormat/render";
import {
  formulaCandidateValue,
  sheetValidations,
  validateValue,
  validationAt,
  type DataValidation,
  type ValidationDeps,
  type ValidationErrorStyle,
} from "../../lib/spreadsheet/validation";

/** What the gate checks: the entry's value, or its formula (canonical, as
 * written at the edited cell — `origin`). */
export interface EntryCandidate {
  value: CellValue;
  formula?: string;
}

/** The user's answer to the alert (or "commit" when nothing failed). */
export type EntryVerdict = "commit" | "retry" | "cancel";

/** The first target a candidate fails at, and the rule it breaks. */
export interface ValidationFailure {
  at: CellRef;
  rule: DataValidation;
}

/** The alert's buttons. */
export type ValidationAlertChoice = "retry" | "cancel" | "yes" | "no" | "ok";

/** What `validation-alert` shows. */
export interface ValidationAlertState {
  style: ValidationErrorStyle;
  title: string;
  text: string;
}

export const DEFAULT_ALERT_TITLE = "Invalid entry";
export const DEFAULT_ALERT_TEXT = "This value doesn't match the data validation rules for this cell.";

/** The verdict each alert button gives (see the header). */
export const ALERT_VERDICTS: Readonly<Record<ValidationAlertChoice, EntryVerdict>> = {
  retry: "retry",
  cancel: "cancel",
  yes: "commit",
  no: "retry",
  ok: "commit",
};

/** SPEC-130's entry gate as useCellEditing's commit path uses it. */
export interface EntryGate {
  /** The first of `targets` (on `sheet`) refusing the candidate under a rule
   * that shows its alert, or null — synchronous, so a valid entry commits at
   * once. `origin`: where a formula candidate is written (each target
   * evaluates it shifted from there). */
  firstFailure(sheet: Sheet, targets: Iterable<CellRef>, candidate: EntryCandidate, origin?: CellRef): ValidationFailure | null;
  /** The cells of a Ctrl/Cmd+Enter fill a rule covers, in fill order. */
  fillTargets(sheet: Sheet, ranges: readonly GridRange[], origin: CellRef): Iterable<CellRef>;
  /** Opens the alert for a failure; resolves with the user's answer. */
  alert(failure: ValidationFailure): Promise<EntryVerdict>;
}

/** The deps validateValue and listItems evaluate with: stored values, the
 * user's locale. */
export function validationDepsFor(model: SpreadsheetModel, locale: SpreadsheetLocale): ValidationDeps {
  return { base: storedValueContext(model, { locale }), ctx: workbookFormatContext(model, locale) };
}

/** Whether a rule's verdict is the same at every cell it covers for a plain
 * value: no `custom` overlay, and its formulas hold only constants and
 * absolute references (no relative part, name or function — ROW() and a
 * name may depend on the cell). A fill checks such a rule once. */
export function positionIndependent(rule: DataValidation): boolean {
  if (rule.type === "custom") {
    return false;
  }
  return [rule.formula1, rule.formula2].every(
    (formula) =>
      formula === undefined ||
      tokenize(formula).every((token) => {
        if (token.kind === "ref") {
          const ref = token.ref;
          const parts = ref ? [ref.start, ...(ref.end ? [ref.end] : [])] : [];
          return parts.every((part) => (part.row === undefined || part.rowAbs) && (part.col === undefined || part.colAbs));
        }
        return token.kind !== "name" && token.kind !== "function" && token.kind !== "structured" && token.kind !== "external" && token.kind !== "unknown";
      }),
  );
}

/** The cells of a Ctrl/Cmd+Enter fill a rule covers (sparse: only the
 * rules' rectangles intersected with what the fill writes — lib's
 * fillTargets — never every cell of the selection), in fill order: range by
 * range, rows top to bottom, columns left to right; a merge's covered cells
 * and a cell already listed skipped, as the fill skips them. */
export function* validatedFillTargets(sheet: Sheet, ranges: readonly GridRange[], origin: CellRef): Generator<CellRef> {
  const rules = sheetValidations(sheet);
  if (rules.length === 0) {
    return;
  }
  const merges = buildMergeIndex(sheet.merges);
  const seen = new Set<number>();
  for (const range of cellFillTargets(sheet, ranges, origin)) {
    const pieces: GridRange[] = [];
    for (const rule of rules) {
      for (const covered of rule.ranges) {
        const top = Math.max(range.top, covered.top);
        const bottom = Math.min(range.bottom, covered.bottom);
        const left = Math.max(range.left, covered.left);
        const right = Math.min(range.right, covered.right);
        if (top <= bottom && left <= right) {
          pieces.push({ top, left, bottom, right });
        }
      }
    }
    const byTop = pieces.toSorted((a, b) => a.top - b.top);
    let row = byTop[0]?.top ?? 0;
    while (byTop.length > 0) {
      const spans = byTop.filter((piece) => piece.top <= row && row <= piece.bottom).toSorted((a, b) => a.left - b.left);
      if (spans.length === 0) {
        const next = byTop.find((piece) => piece.top > row);
        if (!next) {
          break;
        }
        row = next.top;
        continue;
      }
      let from = 0;
      for (const span of spans) {
        for (let col = Math.max(from, span.left); col <= span.right; col += 1) {
          const key = row * MAX_COLS + col;
          const merge = mergeAt(merges, row, col);
          if (seen.has(key) || (merge && (merge.top !== row || merge.left !== col))) {
            continue;
          }
          seen.add(key);
          yield { row, col };
        }
        from = Math.max(from, span.right + 1);
      }
      row += 1;
    }
  }
}

/** The first of `targets` refusing `candidate` under a rule that shows its
 * alert, or null (see the module header). */
export function firstValidationFailure(
  model: SpreadsheetModel,
  sheet: Sheet,
  targets: Iterable<CellRef>,
  candidate: EntryCandidate,
  deps: ValidationDeps,
  origin?: CellRef,
): ValidationFailure | null {
  if (sheetValidations(sheet).length === 0) {
    return null;
  }
  // A plain value under a position-independent rule is checked once per rule.
  const settled = new Map<string, boolean>();
  for (const at of targets) {
    const rule = validationAt(sheet, at);
    if (!rule || rule.type === "any" || !rule.error.show) {
      continue;
    }
    const memo = candidate.formula === undefined && positionIndependent(rule);
    let ok = memo ? settled.get(rule.id) : undefined;
    if (ok === undefined) {
      let value: CellValue | undefined = candidate.value;
      if (candidate.formula !== undefined) {
        const formula = origin ? shiftFormula(candidate.formula, at.row - origin.row, at.col - origin.col) : candidate.formula;
        value = formulaCandidateValue(sheet, at, formula, deps);
      }
      ok = value === undefined || validateValue(model, sheet, at, value, deps).ok;
      if (memo) {
        settled.set(rule.id, ok);
      }
    }
    if (!ok) {
      return { at: { ...at }, rule };
    }
  }
  return null;
}

/** The alert a failing rule shows. */
export function alertFor(rule: DataValidation): ValidationAlertState {
  return {
    style: rule.error.style,
    title: rule.error.title.trim() === "" ? DEFAULT_ALERT_TITLE : rule.error.title,
    text: rule.error.text.trim() === "" ? DEFAULT_ALERT_TEXT : rule.error.text,
  };
}
