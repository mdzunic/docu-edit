// The read-only selection model (SPEC-117 "Selection"). Placeholder types —
// the navigation functions are implemented below them.
import type { CellRef, GridRange } from "./address";
import type { MergeIndex } from "./merges";
import type { Sheet } from "./model";

export interface SheetSelection {
  active: CellRef;
  ranges: GridRange[];
  activeRangeIndex: number;
}

export type Direction = "up" | "down" | "left" | "right";

export interface NavContext {
  sheet: Sheet;
  merges: MergeIndex;
  isRowHidden(r: number): boolean;
  isColHidden(c: number): boolean;
  pageRows: number; // visible rows/cols for PageUp/PageDown
  pageCols: number;
}
