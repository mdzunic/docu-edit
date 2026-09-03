// Table style banding (SPEC-076): the four-level resolution that turns a
// table style's base formatting, its banding, its edge regions and a cell's
// own overrides into the one border/shading set a renderer paints — plus the
// five built-in gallery looks.
//
// Framework-free by contract (AC-24, AC-78): no Vue import, no DOM. Both
// renderers (the DOM `DocTableBlock.vue` and the canvas page renderer) call
// `resolveCellFormatting`, so a banded table looks the same on screen, in
// print preview and in the PDF export (AC-29).
//
// Nothing here ever writes into a cell. That is the whole point: banding is
// resolved at paint time, so toggling `look.bandedRows` off restores the
// cell's own shading by construction rather than by a cleanup pass (AC-26).
// SPEC-042's separate bake of a style's *base* borders/shading onto cells
// when a style is applied from the Styles pane is untouched (AC-30) — it
// simply means levels 1 and 4 agree for such a table, and the resolver stays
// correct for a table that got its style from `w:tblStyle` on import and was
// never baked.
import type { NamedStyle, Table, TableConditionalFormat, TableLook, TableStyleBanding } from "./documentModel";
import type { CellBorders } from "./core/tables";

/** A cell's effective paint, resolved from the table, its style and the
 * cell's own overrides. Every field is absent when nothing at any level sets
 * it — an absent border side falls through to the canvas's default hairline
 * gridline, exactly as an unset cell border does today. */
export interface ResolvedCellFormatting {
  borders?: CellBorders;
  shadingFill?: string;
  /** From a conditional region's `bold` — the header/accent looks' emphasis.
   * Absent when no enabled region sets it. */
  bold?: boolean;
}

/** How many rows the table's leading contiguous `isHeader` run covers — the
 * run that repeats across page breaks (AC-1) and that row banding starts
 * counting after (AC-27). A table whose first row is not `isHeader` has a run
 * of zero, even when a later row carries the flag (AC-5). */
export function leadingHeaderRowCount(table: Table): number {
  let count = 0;
  while (count < table.rows.length && table.rows[count].isHeader) {
    count += 1;
  }
  return count;
}

/** The band size a banding entry declares, floored at 1 — a size of 0 or a
 * negative from a malformed source would otherwise divide by zero. */
function bandSize(size: number | undefined): number {
  return size !== undefined && size >= 1 ? Math.floor(size) : 1;
}

/** Whether the index at `position` falls in an odd-numbered band (AC-27):
 * bands are blocks of `size` counted from `start`, and the first, third,
 * fifth… take the shading. Positions before `start` or at/after `end` are in
 * no band at all — they belong to the header run, the first/last row or the
 * first/last column, which level 3 paints instead. */
function inOddBand(position: number, start: number, end: number, size: number): boolean {
  if (position < start || position >= end) {
    return false;
  }
  return Math.floor((position - start) / size) % 2 === 0;
}

/** `base` with every side `over` sets replaced — per side, so a region that
 * sets only `top` leaves the other three showing the level below it. */
function mergeBorders(base: CellBorders | undefined, over: CellBorders | undefined): CellBorders | undefined {
  if (!over) {
    return base;
  }
  if (!base) {
    return { ...over };
  }
  return { ...base, ...over };
}

/** Applies one conditional region over the accumulating result. */
function applyConditional(into: ResolvedCellFormatting, format: TableConditionalFormat | undefined): void {
  if (!format) {
    return;
  }
  if (format.shading !== undefined) {
    into.shadingFill = format.shading;
  }
  if (format.bold !== undefined) {
    into.bold = format.bold;
  }
  into.borders = mergeBorders(into.borders, format.borders);
}

/** A cell's effective borders and shading (AC-24), applying the four-level
 * resolution order, lowest first:
 *
 * 1. the table style's base `cellBorders` / `cellShadingFill`;
 * 2. row/column banding from the style, where the table's `look` enables it;
 * 3. the first/last row and column conditional formats, where enabled;
 * 4. the cell's own `borders` / `shadingFill`.
 *
 * Within level 2 the row band wins over the column band, and within level 3
 * the row regions win over the column regions and `firstRow` wins over
 * `lastRow` — Word's own `w:tblStylePr` precedence, so a table imported from
 * a real document paints the way its author saw it.
 *
 * `style` is the table's resolved catalogue entry, or undefined for a table
 * with no style: banding then contributes nothing and only the cell's own
 * formatting survives, which is exactly today's rendering (AC-23/D-10).
 */
export function resolveCellFormatting(
  table: Table,
  style: NamedStyle | undefined,
  rowIndex: number,
  colIndex: number,
): ResolvedCellFormatting {
  const cell = table.rows[rowIndex]?.cells[colIndex];
  const resolved: ResolvedCellFormatting = {};

  // Level 1 — the style's base.
  if (style?.cellShadingFill !== undefined) {
    resolved.shadingFill = style.cellShadingFill;
  }
  resolved.borders = mergeBorders(resolved.borders, style?.cellBorders);

  const banding = style?.banding;
  const look: TableLook = table.look ?? {};
  if (banding) {
    const rowCount = table.rows.length;
    const colCount = table.columnWidthsCm.length;
    const rowStart = Math.max(leadingHeaderRowCount(table), look.firstRow ? 1 : 0);
    const rowEnd = look.lastRow ? Math.max(rowStart, rowCount - 1) : rowCount;
    const colStart = look.firstColumn ? 1 : 0;
    const colEnd = look.lastColumn ? Math.max(colStart, colCount - 1) : colCount;

    // Level 2 — banding, columns first so the row band wins where both apply.
    if (look.bandedColumns && banding.bandedColumns) {
      if (inOddBand(colIndex, colStart, colEnd, bandSize(banding.bandedColumns.size))) {
        resolved.shadingFill = banding.bandedColumns.shading;
      }
    }
    if (look.bandedRows && banding.bandedRows) {
      if (inOddBand(rowIndex, rowStart, rowEnd, bandSize(banding.bandedRows.size))) {
        resolved.shadingFill = banding.bandedRows.shading;
      }
    }

    // Level 3 — the edge regions, lowest precedence first.
    if (look.lastColumn && colIndex === colCount - 1) {
      applyConditional(resolved, banding.lastColumn);
    }
    if (look.firstColumn && colIndex === 0) {
      applyConditional(resolved, banding.firstColumn);
    }
    if (look.lastRow && rowIndex === rowCount - 1) {
      applyConditional(resolved, banding.lastRow);
    }
    if (look.firstRow && rowIndex === 0) {
      applyConditional(resolved, banding.firstRow);
    }
  }

  // Level 4 — the cell's own formatting always wins (SPEC-023's rule that
  // cell-level formatting keeps precedence over everything a style says).
  if (cell?.shadingFill !== undefined) {
    resolved.shadingFill = cell.shadingFill;
  }
  resolved.borders = mergeBorders(resolved.borders, cell?.borders);

  if (resolved.borders === undefined) {
    delete resolved.borders;
  }
  return resolved;
}

/** The catalogue entry `table.styleId` names, or undefined — the lookup both
 * renderers use before calling `resolveCellFormatting`. */
export function tableStyleFor(styles: NamedStyle[] | undefined, table: Table): NamedStyle | undefined {
  return table.styleId === undefined ? undefined : styles?.find((style) => style.id === table.styleId);
}

/** Whether a style carries any banding at all — the Markdown drop check
 * (AC-74) and the gallery's "is this look applied" test. */
export function styleHasBanding(style: NamedStyle | undefined): boolean {
  const banding = style?.banding;
  return (
    banding !== undefined &&
    (banding.bandedRows !== undefined ||
      banding.bandedColumns !== undefined ||
      banding.firstRow !== undefined ||
      banding.lastRow !== undefined ||
      banding.firstColumn !== undefined ||
      banding.lastColumn !== undefined)
  );
}

// --- the built-in gallery ---------------------------------------------------

/** A single 0.5pt black line — the "grid borders" four of the five built-ins
 * share. A fresh object per side per call: the catalogue entries are copied
 * into a document's own `styles`, and two styles must never alias one
 * border object. */
function gridBorders(): CellBorders {
  const side = { style: "single" as const, widthPt: 0.5, color: "#000000" };
  return { top: { ...side }, bottom: { ...side }, left: { ...side }, right: { ...side } };
}

/** One entry of the Table styles gallery (AC-31): a built-in style
 * definition plus the `look` applying it turns on (AC-33). */
export interface BuiltInTableStyle {
  id: string;
  name: string;
  /** The `look` a click applies alongside the style id. */
  look: TableLook;
}

/** The five built-in looks, in gallery order. Ids double as the `styleId`
 * written into `Table.styleId` and as the catalogue entry's `id`, so an
 * applied look round-trips as a real `w:style` / `style:style` definition
 * rather than as a dangling reference (AC-33). */
export const BUILT_IN_TABLE_STYLES: BuiltInTableStyle[] = [
  { id: "TableStylePlain", name: "Plain", look: {} },
  { id: "TableStyleGrid", name: "Grid", look: {} },
  { id: "TableStyleBandedRows", name: "Banded rows", look: { bandedRows: true } },
  { id: "TableStyleHeader", name: "Header emphasis", look: { firstRow: true } },
  { id: "TableStyleAccent", name: "Accent", look: { firstRow: true, bandedRows: true } },
];

/** One gallery row: the five built-ins preceded by the "None" entry that
 * clears both `styleId` and `look` (AC-34). `id` null IS "None" — a distinct
 * type from a built-in's string id, which is why the list is built here
 * rather than spelled out in the espree-parsed dialog. */
export interface TableGalleryEntry {
  id: string | null;
  name: string;
  look: TableLook;
}

export function tableStyleGalleryEntries(): TableGalleryEntry[] {
  return [{ id: null, name: "None", look: {} }, ...BUILT_IN_TABLE_STYLES];
}

/** The catalogue entry for a built-in id, freshly built so the caller can
 * push it straight into `DocumentModel.styles` without aliasing. Returns
 * undefined for an id that is not a built-in. */
export function builtInTableStyle(id: string): NamedStyle | undefined {
  switch (id) {
    case "TableStylePlain":
      return { id, name: "Plain", type: "table" };
    case "TableStyleGrid":
      return { id, name: "Grid", type: "table", cellBorders: gridBorders() };
    case "TableStyleBandedRows":
      return {
        id,
        name: "Banded rows",
        type: "table",
        cellBorders: gridBorders(),
        banding: { bandedRows: { size: 1, shading: "#f2f2f2" } },
      };
    case "TableStyleHeader":
      return {
        id,
        name: "Header emphasis",
        type: "table",
        cellBorders: gridBorders(),
        banding: { firstRow: { shading: "#e6e6e6", bold: true } },
      };
    case "TableStyleAccent":
      return {
        id,
        name: "Accent",
        type: "table",
        cellBorders: gridBorders(),
        banding: {
          firstRow: { shading: "#2f5496", bold: true },
          bandedRows: { size: 1, shading: "#d9e2f3" },
        },
      };
    default:
      return undefined;
  }
}

/** The six look flags in the order the gallery's checkboxes list them
 * (AC-35). Shared by the dialog and by both writers' attribute order. */
export const TABLE_LOOK_FLAGS: { key: keyof TableLook; label: string }[] = [
  { key: "firstRow", label: "Header row" },
  { key: "lastRow", label: "Total row" },
  { key: "bandedRows", label: "Banded rows" },
  { key: "firstColumn", label: "First column" },
  { key: "lastColumn", label: "Last column" },
  { key: "bandedColumns", label: "Banded columns" },
];

/** Typed pass-through for a look flag key arriving from a `.vue` template's
 * plain string (espree bars the annotation there — same idiom as
 * core/tables.ts's asBorderLineStyle). */
export function asTableLookKey(value: string): keyof TableLook {
  return value as keyof TableLook;
}

// --- the style editor's banding draft (AC-37) -------------------------------

/** One row of the style editor's banding grid: a region, and the three things
 * the dialog lets an author set on it. Strings and booleans only, because the
 * espree-parsed `.vue` script binds them straight to inputs. */
export interface TableBandingRegionDraft {
  key: string;
  label: string;
  shading: string;
  bold: boolean;
  borderStyle: string;
}

/** The whole banding draft: the two band sizes plus the six regions. */
export interface TableBandingDraft {
  rowBandSize: number;
  columnBandSize: number;
  regions: TableBandingRegionDraft[];
}

/** The six regions in the order the dialog lists them, with the labels it
 * shows. The keys are `TableStyleBanding`'s own field names. */
const BANDING_REGIONS: { key: keyof TableStyleBanding; label: string }[] = [
  { key: "firstRow", label: "Header row" },
  { key: "lastRow", label: "Total row" },
  { key: "firstColumn", label: "First column" },
  { key: "lastColumn", label: "Last column" },
  { key: "bandedRows", label: "Banded rows" },
  { key: "bandedColumns", label: "Banded columns" },
];

/** A style's banding as the editor's draft (AC-37). An absent region reads
 * as an empty row, which `bandingFromDraft` turns back into absence. */
export function bandingDraftFrom(banding: TableStyleBanding | undefined): TableBandingDraft {
  return {
    rowBandSize: banding?.bandedRows?.size ?? 1,
    columnBandSize: banding?.bandedColumns?.size ?? 1,
    regions: BANDING_REGIONS.map((region) => {
      const value = banding?.[region.key];
      const format: TableConditionalFormat | undefined =
        region.key === "bandedRows" || region.key === "bandedColumns"
          ? value
            ? { shading: (value as { shading: string }).shading }
            : undefined
          : (value as TableConditionalFormat | undefined);
      return {
        key: region.key,
        label: region.label,
        shading: format?.shading ?? "",
        bold: format?.bold ?? false,
        // One border style per region; the width and colour come from the
        // style's own border controls, so a style's borders are described in
        // one place rather than six.
        borderStyle: format?.borders?.top?.style ?? "",
      };
    }),
  };
}

/** The draft back as banding, or undefined when no region sets anything —
 * so a table style with an untouched banding section carries no `banding`
 * field at all rather than six empty objects (AC-22's "independently
 * absent"). */
export function bandingFromDraft(
  draft: TableBandingDraft,
  borderWidthPt: number,
  borderColor: string,
): TableStyleBanding | undefined {
  const banding: TableStyleBanding = {};
  for (const region of draft.regions) {
    const shading = region.shading.trim();
    if (region.key === "bandedRows" || region.key === "bandedColumns") {
      if (shading.length === 0) {
        continue;
      }
      const size = region.key === "bandedRows" ? draft.rowBandSize : draft.columnBandSize;
      banding[region.key] = { size: size >= 1 ? Math.floor(size) : 1, shading };
      continue;
    }
    const format: TableConditionalFormat = {};
    if (shading.length > 0) {
      format.shading = shading;
    }
    if (region.bold) {
      format.bold = true;
    }
    if (region.borderStyle) {
      const side = {
        style: region.borderStyle as NonNullable<CellBorders["top"]>["style"],
        widthPt: borderWidthPt,
        color: borderColor,
      };
      format.borders = { top: { ...side }, bottom: { ...side }, left: { ...side }, right: { ...side } };
    }
    if (Object.keys(format).length > 0) {
      banding[region.key as "firstRow" | "lastRow" | "firstColumn" | "lastColumn"] = format;
    }
  }
  return Object.keys(banding).length > 0 ? banding : undefined;
}

/** Typed pass-through for the draft arriving as a `.vue` template's plain
 * `Object` prop (espree bars the annotation there). */
export function asTableBandingDraft(value: unknown): TableBandingDraft {
  return (value ?? bandingDraftFrom(undefined)) as TableBandingDraft;
}
