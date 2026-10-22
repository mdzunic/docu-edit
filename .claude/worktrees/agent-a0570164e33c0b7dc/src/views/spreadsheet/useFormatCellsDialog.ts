// The Format Cells dialog's state (SPEC-123 "Format Cells"): it opens on the
// ACTIVE cell's effective style, tracks which fields the user changed, shows
// a live preview of the active cell's value with the pending style
// (`format-cells-preview`) and the Number tab's sample (`format-cells-sample`,
// through SPEC-118's formatValue as the user types), and on OK hands
// useFormatting ONE list of patch entries — only the changed fields
// (lib/spreadsheet/formatCells.ts), so a dialog opened on a mixed selection
// never flattens it. A Custom code SPEC-118's parseFormatCode rejects shows
// CUSTOM_CODE_ERROR and disables OK. The preview box draws the pending outer
// edges too; the Border tab's "None" line style removes the edges it is
// applied to.
//
// SPEC-129 "The rule editor" adds a second, DIFFERENTIAL mode: showDifferential
// opens the same tabs over a Dxf instead of a cell's style — Number, Font
// (colour/bold/italic/underline/strike; ConditionalFormatEditor.vue hides
// family and size, which a Dxf has no room for), Border (top/right/bottom/
// left only — no inside or diagonal, which a single differential cell has no
// use for) and Fill; Alignment and Protection stay hidden. `ok()` branches on
// `mode`: a differential OK reads the fields back through dxfFields.ts's
// fieldsAsDxf and hands the result to the callback showDifferential was
// given, instead of useFormatting's `apply`.
import { computed, ref, shallowRef } from "vue";
import type { DocumentTheme, ModelColor } from "../../lib/core/theme";
import type { GridRange } from "../../lib/spreadsheet/address";
import { AUTOMATIC_COLOR, borderShorthand, cellCss, type CssDeclarations } from "../../lib/spreadsheet/cellRender";
import type { Dxf } from "../../lib/spreadsheet/conditionalFormats";
import { borderEdgeEntries, fieldsOf, fieldsPatch, type DialogEdge, type FormatCellsFields } from "../../lib/spreadsheet/formatCells";
import { patchStyle, type StylePatch, type StylePatchEntry } from "../../lib/spreadsheet/formatting";
import type { CellBorderEdge, CellStyle, CellValue, SheetBorderStyle } from "../../lib/spreadsheet/model";
import { buildNumberCode, describeNumberCode, isValidFormatCode, numberFormatChoiceOf, type NumberCodeSpec } from "../../lib/spreadsheet/numberCodes";
import { BUILTIN_NUMBER_FORMATS } from "../../lib/spreadsheet/numberFormat/builtin";
import { cellColorHex, cellFontName } from "../../lib/spreadsheet/styles";
import type { SpreadsheetLocale } from "../../lib/spreadsheet/numberFormat/locale";
import { formatValue, plainText, type FormatContext } from "../../lib/spreadsheet/numberFormat/render";
import { dxfAsStyle, dxfTouchedFields, fieldsAsDxf } from "./dxfFields";
import type { FormatCellsTab } from "./sheetToolbar";

export const CUSTOM_CODE_ERROR = "This number format isn't valid";

/** SPEC-129: what showDifferential opens on — a Dxf rather than a cell. */
export interface DifferentialContext {
  theme: DocumentTheme;
  locale: SpreadsheetLocale;
  documentFonts?: string[];
}

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

/** A representative value the Number tab's sample formats (SPEC-129's
 * differential target has no real cell behind it). */
const DIFFERENTIAL_SAMPLE_VALUE = 1234.56;

export function useFormatCellsDialog(deps: FormatCellsDialogDeps) {
  const open = ref(false);
  const tab = ref<FormatCellsTab>("number");
  const mode = ref<"cells" | "differential">("cells");
  // Not reactive: set immediately before showDifferential opens the dialog
  // and consumed (then cleared) by its own ok()/cancel() — nothing reads it
  // between, so a plain closure variable is simpler than a ref nothing
  // templates.
  let differentialSave: ((dxf: Dxf) => void) | null = null;
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
  const lineStyle = ref<SheetBorderStyle | "none">("thin");
  const lineColor = shallowRef<ModelColor | null>(null);

  /** Opens on the active cell's effective style (see the module header). */
  function show(at: FormatCellsTab = "number"): void {
    const t = deps.target();
    if (!t) {
      return;
    }
    mode.value = "cells";
    differentialSave = null;
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

  /** SPEC-129: opens on `initial` (a conditional format's Dxf, `{}` for a
   * brand-new rule) instead of a cell's style; OK hands the rebuilt Dxf to
   * `onSave` rather than useFormatting's `apply` (see the module header). */
  function showDifferential(initial: Dxf, ctx: DifferentialContext, onSave: (dxf: Dxf) => void): void {
    mode.value = "differential";
    differentialSave = onSave;
    const style = dxfAsStyle(initial);
    target.value = {
      style,
      value: DIFFERENTIAL_SAMPLE_VALUE,
      locale: ctx.locale,
      format: { locale: ctx.locale, dateSystem: "1900" },
      theme: ctx.theme,
      documentFonts: ctx.documentFonts ?? [],
      styles: [],
      ranges: [{ top: 0, left: 0, bottom: 0, right: 0 }],
      recentColors: [],
    };
    fields.value = fieldsOf(style, "", 11);
    touched.value = dxfTouchedFields(initial);
    number.value = describeNumberCode(initial.numberFormat, ctx.locale);
    customCode.value = number.value.category === "custom" ? (number.value.type ?? "") : (initial.numberFormat ?? "General");
    numberTouched.value = initial.numberFormat !== undefined;
    edges.value = {};
    tab.value = "number";
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

  /** The edge the current line style and colour draw (null: None). */
  function line(): CellBorderEdge | null {
    if (lineStyle.value === "none") {
      return null;
    }
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

  /** Clicking an edge: on in the current line style and colour, or off
   * (always off in None). */
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

  /** A model colour as a picker's bar shows it ("" = automatic). */
  function colorHex(color: ModelColor | null | undefined): string {
    const t = target.value;
    return t ? cellColorHex(color, t.theme) : "";
  }

  /** The active cell's value with the pending style, in a box drawing the
   * pending outer edges. */
  const preview = computed((): { text: string; box: CssDeclarations; textCss: CssDeclarations } => {
    const t = target.value;
    if (!t) {
      return { text: "", box: {}, textCss: {} };
    }
    const sides = ["top", "bottom", "left", "right"] as const;
    const style = patchStyle(t.style, { ...pendingPatch(), borders: Object.fromEntries(sides.map((side) => [side, edgeShown(side)])) });
    const css = cellCss(style, t.theme, typeof t.value === "number" ? "number" : "string");
    const box = { ...css.box };
    for (const side of sides) {
      const edge = style.borders?.[side];
      if (edge) {
        box[`border-${side}`] = borderShorthand(edge.style, colorHex(edge.color) || AUTOMATIC_COLOR);
      }
    }
    return { text: sample.value, box, textCss: css.text };
  });

  function ok(): void {
    const t = target.value;
    if (!t || codeError.value !== "") {
      return;
    }
    if (mode.value === "differential") {
      const dxf = fieldsAsDxf(fields.value, touched.value, edgeShown, code.value, numberTouched.value);
      const onSave = differentialSave;
      open.value = false;
      differentialSave = null;
      onSave?.(dxf);
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
    mode,
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
    colorHex,
    multiCell,
    show,
    showDifferential,
    ok,
    cancel: (): void => {
      open.value = false;
      differentialSave = null;
    },
    showTab: (next: string): void => {
      tab.value = next as FormatCellsTab;
    },
    setLineStyle: (style: string): void => {
      lineStyle.value = style as SheetBorderStyle | "none";
    },
  };
}

export type FormatCellsDialogState = ReturnType<typeof useFormatCellsDialog>;

/** Narrows a component's runtime-typed `dialog` prop (the .vue parser bars
 * casts). */
export function asFormatCellsDialog(value: unknown): FormatCellsDialogState {
  return value as FormatCellsDialogState;
}
