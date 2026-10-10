// The Format Cells dialog's state (SPEC-123 "Format Cells"): it opens on the
// ACTIVE cell's effective style, tracks which fields the user changed, shows
// a live preview of the active cell's value with the pending style
// (`format-cells-preview`) and the Number tab's sample (`format-cells-sample`,
// through SPEC-118's formatValue as the user types), and on OK hands
// useFormatting ONE list of patch entries — only the changed fields
// (lib/spreadsheet/formatCells.ts), so a dialog opened on a mixed selection
// never flattens it. A Custom code SPEC-118's parseFormatCode rejects shows
// CUSTOM_CODE_ERROR and disables OK.
import { computed, ref, shallowRef } from "vue";
import type { DocumentTheme, ModelColor } from "../../lib/core/theme";
import type { GridRange } from "../../lib/spreadsheet/address";
import { cellCss, type CssDeclarations } from "../../lib/spreadsheet/cellRender";
import { borderEdgeEntries, fieldsOf, fieldsPatch, type DialogEdge, type FormatCellsFields } from "../../lib/spreadsheet/formatCells";
import { patchStyle, type StylePatch, type StylePatchEntry } from "../../lib/spreadsheet/formatting";
import type { CellBorderEdge, CellStyle, CellValue, SheetBorderStyle } from "../../lib/spreadsheet/model";
import { buildNumberCode, describeNumberCode, isValidFormatCode, numberFormatChoiceOf, type NumberCodeSpec } from "../../lib/spreadsheet/numberCodes";
import { BUILTIN_NUMBER_FORMATS } from "../../lib/spreadsheet/numberFormat/builtin";
import { cellFontName } from "../../lib/spreadsheet/styles";
import type { SpreadsheetLocale } from "../../lib/spreadsheet/numberFormat/locale";
import { formatValue, plainText, type FormatContext } from "../../lib/spreadsheet/numberFormat/render";
import type { FormatCellsTab } from "./sheetToolbar";

export const CUSTOM_CODE_ERROR = "This number format isn't valid";

/** What the dialog opens on (useFormatting's active cell and workbook). */
export interface FormatCellsTarget {
  style: CellStyle;
  value: CellValue;
  locale: SpreadsheetLocale;
  format: FormatContext;
  theme: DocumentTheme;
  documentFonts: string[];
  /** The workbook's styles — their custom codes fill the Custom list. */
  styles: readonly CellStyle[];
  ranges: GridRange[];
  recentColors: string[];
}

export interface FormatCellsDialogDeps {
  target: () => FormatCellsTarget | null;
  /** Runs the entries as ONE "Format cells" step. */
  apply: (entries: StylePatchEntry[]) => void;
}

const EDGES: readonly DialogEdge[] = ["top", "insideHorizontal", "bottom", "left", "insideVertical", "right", "diagonalUp", "diagonalDown"];

export function useFormatCellsDialog(deps: FormatCellsDialogDeps) {
  const open = ref(false);
  const tab = ref<FormatCellsTab>("number");
  const target = shallowRef<FormatCellsTarget | null>(null);
  // Shallow, and replaced rather than mutated: a picked colour or border edge
  // must reach patchStyle as the plain object the picker emitted — a deep ref
  // hands back a reactive Proxy, which no structured clone (patchStyle's,
  // IndexedDB's) accepts.
  const fields = shallowRef<FormatCellsFields>(fieldsOf({}, "", 11));
  const touched = shallowRef(new Set<keyof FormatCellsFields>());
  const number = shallowRef<NumberCodeSpec>({ category: "general" });
  const numberTouched = ref(false);
  const customCode = ref("");
  const edges = shallowRef<Partial<Record<DialogEdge, CellBorderEdge | null>>>({});
  const lineStyle = ref<SheetBorderStyle>("thin");
  const lineColor = shallowRef<ModelColor | null>(null);

  /** Opens on the active cell's effective style (see the module header). */
  function show(at: FormatCellsTab = "number"): void {
    const t = deps.target();
    if (!t) {
      return;
    }
    target.value = t;
    const normal = t.styles[0]?.font;
    const family = cellFontName(t.style.font?.family ?? normal?.family, t.theme);
    fields.value = fieldsOf(t.style, family, normal?.sizePt ?? 11);
    touched.value = new Set();
    number.value = describeNumberCode(t.style.numberFormat, t.locale);
    customCode.value = number.value.category === "custom" ? (number.value.type ?? "") : (t.style.numberFormat ?? "General");
    numberTouched.value = false;
    edges.value = {};
    tab.value = at;
    open.value = true;
  }

  /** A field changed (the template's choices are the field's own values). */
  function set(key: keyof FormatCellsFields, value: unknown): void {
    fields.value = { ...fields.value, [key]: value };
    touched.value = new Set([...touched.value, key]);
  }

  /** A Number-tab option changed (a category resets its options). */
  function setNumber(change: Record<string, unknown>): void {
    const spec = change as Partial<NumberCodeSpec>;
    const next: NumberCodeSpec = spec.category !== undefined && spec.category !== number.value.category ? { category: spec.category } : { ...number.value, ...spec };
    if (next.category === "custom" && next.type === undefined) {
      next.type = customCode.value;
    }
    number.value = next;
    numberTouched.value = true;
  }

  function setCustomCode(code: string): void {
    customCode.value = code;
    number.value = { category: "custom", type: code };
    numberTouched.value = true;
  }

  const locale = (): SpreadsheetLocale | undefined => target.value?.locale;
  /** The pending code (undefined = General). */
  const code = computed(() => {
    const l = locale();
    if (number.value.category === "custom") {
      return customCode.value.trim() === "" ? undefined : customCode.value;
    }
    return l ? buildNumberCode(number.value, l) : undefined;
  });
  const codeError = computed(() => (number.value.category === "custom" && code.value !== undefined && !isValidFormatCode(code.value) ? CUSTOM_CODE_ERROR : ""));

  /** Every built-in code, then every custom code the workbook holds. */
  const customList = computed(() => {
    const codes = new Set(Object.values(BUILTIN_NUMBER_FORMATS));
    for (const style of target.value?.styles ?? []) {
      if (style.numberFormat !== undefined) {
        codes.add(style.numberFormat);
      }
    }
    return [...codes];
  });

  const sample = computed(() => {
    const t = target.value;
    if (!t || codeError.value !== "") {
      return "";
    }
    return plainText(formatValue(t.value, code.value, t.format));
  });

  /** An edge as the Border tab shows it: the user's change, else the active
   * cell's own (inside edges start off). */
  function edgeShown(edge: DialogEdge | string): CellBorderEdge | null {
    const changed = edges.value[edge as DialogEdge];
    if (changed !== undefined) {
      return changed;
    }
    const borders = target.value?.style.borders ?? {};
    switch (edge) {
      case "top":
      case "bottom":
      case "left":
      case "right":
        return borders[edge] ?? null;
      case "diagonalUp":
        return borders.diagonal?.up ? borders.diagonal : null;
      case "diagonalDown":
        return borders.diagonal?.down ? borders.diagonal : null;
      default:
        return null;
    }
  }

  function line(): CellBorderEdge {
    return lineColor.value === null ? { style: lineStyle.value } : { style: lineStyle.value, color: lineColor.value };
  }

  function setEdges(changes: Partial<Record<DialogEdge, CellBorderEdge | null>>): void {
    const next = { ...edges.value, ...changes };
    // The diagonals are one model edge: pin the other at what it shows.
    if (changes.diagonalUp !== undefined && next.diagonalDown === undefined) {
      next.diagonalDown = edgeShown("diagonalDown");
    }
    if (changes.diagonalDown !== undefined && next.diagonalUp === undefined) {
      next.diagonalUp = edgeShown("diagonalUp");
    }
    edges.value = next;
  }

  const multiCell = computed(() => (target.value?.ranges ?? []).some((range) => range.top !== range.bottom || range.left !== range.right));

  /** Clicking an edge: on in the current line style and colour, or off. */
  function toggleEdge(edge: DialogEdge | string): void {
    setEdges({ [edge]: edgeShown(edge) ? null : line() });
  }

  function preset(kind: "none" | "outline" | "inside" | string): void {
    if (kind === "none") {
      setEdges(Object.fromEntries(EDGES.map((edge) => [edge, null])));
    } else if (kind === "outline") {
      setEdges({ top: line(), bottom: line(), left: line(), right: line() });
    } else if (multiCell.value) {
      setEdges({ insideHorizontal: line(), insideVertical: line() });
    }
  }

  /** The changed fields as one patch (the borders apart). */
  function pendingPatch(): StylePatch {
    const patch = fieldsPatch(fields.value, touched.value);
    if (numberTouched.value) {
      const choice = numberFormatChoiceOf(code.value);
      patch.numberFormat = choice.numberFormat;
      patch.builtinFormatId = choice.builtinFormatId ?? null;
    }
    return patch;
  }

  /** The active cell's value with the pending style. */
  const preview = computed((): { text: string; box: CssDeclarations; text_: CssDeclarations } => {
    const t = target.value;
    if (!t) {
      return { text: "", box: {}, text_: {} };
    }
    const borders = Object.fromEntries((["top", "bottom", "left", "right"] as const).map((side) => [side, edgeShown(side)]));
    const style = patchStyle(t.style, { ...pendingPatch(), borders });
    const css = cellCss(style, t.theme, typeof t.value === "number" ? "number" : "string");
    return { text: sample.value, box: css.box, text_: css.text };
  });

  function ok(): void {
    const t = target.value;
    if (!t || codeError.value !== "") {
      return;
    }
    const patch = pendingPatch();
    const entries: StylePatchEntry[] = [];
    for (const range of t.ranges) {
      if (Object.keys(patch).length > 0) {
        entries.push({ range, patch });
      }
      entries.push(...borderEdgeEntries(range, edges.value));
    }
    open.value = false;
    if (entries.length > 0) {
      deps.apply(entries);
    }
  }

  return {
    open,
    tab,
    target,
    fields,
    set,
    number,
    setNumber,
    customCode,
    setCustomCode,
    code,
    codeError,
    customList,
    sample,
    preview,
    edgeShown,
    toggleEdge,
    preset,
    lineStyle,
    lineColor,
    multiCell,
    show,
    ok,
    cancel: (): void => {
      open.value = false;
    },
    showTab: (next: string): void => {
      tab.value = next as FormatCellsTab;
    },
    setLineStyle: (style: string): void => {
      lineStyle.value = style as SheetBorderStyle;
    },
  };
}

export type FormatCellsDialogState = ReturnType<typeof useFormatCellsDialog>;

/** Narrows a component's runtime-typed `dialog` prop (the .vue parser bars
 * casts). */
export function asFormatCellsDialog(value: unknown): FormatCellsDialogState {
  return value as FormatCellsDialogState;
}
