// The render items useSheetCells assembles and SheetCellLayer binds (SPEC-117
// "Rendering"): a band of rows, each row's cells split into the frozen- and
// scrolled-column regions, and the resolved border elements. Plain data —
// no Vue — so the composable that builds them and the helpers that narrow
// them for the component (sheetLayerProps.ts) share one definition.
import type { CellText, CssDeclarations, diagonalLines } from "../../lib/spreadsheet/cellRender";
import type { CfCellResult } from "../../lib/spreadsheet/conditionalFormats";

export interface CellItem {
  key: string;
  ref: string;
  row: number;
  col: number;
  left: number; // relative to the region
  top: number; // relative to the host row
  width: number;
  height: number;
  text: string;
  runs: { text: string; style: CssDeclarations }[] | null;
  /** A number format's `*x` fill (SPEC-118, cellText). */
  fill: CellText["fill"];
  boxStyle: CssDeclarations;
  textStyle: CssDeclarations;
  contentStyle: CssDeclarations;
  kindClass: "is-number" | "is-text" | "is-centred" | "is-empty";
  rowSpan: number;
  colSpan: number;
  merged: boolean;
  /** A merge's paint in a further frozen pane; its element is elsewhere. */
  paintOnly: boolean;
  raised: boolean;
  selected: boolean;
  diagonal: ReturnType<typeof diagonalLines>;
  /** SPEC-129: the cell's conditional-format result — its data bar and
   * icon, and whether it hides its value; absent = none. */
  cf?: CfCellResult;
}

export interface RowItem {
  index: number;
  top: number;
  height: number;
  frozen: CellItem[];
  scrolled: CellItem[];
}

export interface BorderItem {
  key: string;
  style: CssDeclarations;
}

export interface BandItem {
  kind: "frozen" | "main";
  top: number;
  height: number;
  rows: RowItem[];
  gridRows: number[]; // y of each horizontal gridline, band-relative
  gridColsFrozen: number[]; // x of each vertical gridline, region-relative
  gridColsScrolled: number[];
  bordersFrozen: BorderItem[];
  bordersScrolled: BorderItem[];
}
