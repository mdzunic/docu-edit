// Prop narrowing for the workbook's components (the .vue parser bars casts —
// see AppShell.vue on the espree constraint): each helper coerces a
// runtime-typed prop into the shape the composable that built it declares.
import type { BandItem } from "./useSheetCells";

/** SheetCellLayer's `bands` (useSheetCells' BandItem[]). */
export function asBandItems(value: unknown): BandItem[] {
  return Array.isArray(value) ? (value as BandItem[]) : [];
}
