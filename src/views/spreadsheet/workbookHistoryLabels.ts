// The workbook history panel's words (SPEC-138): a version's counts — "3
// sheets · 1,204 cells · +12 cells" — its download's label (the workbook's own
// format and package variant), the restore note and the two notices. Pure, so
// useWorkbookHistory's `statLabel` and `downloadLabels` and their spec share
// one source.
import type { VersionIndexEntry } from "../../lib/core/versionChain";
import { spreadsheetFormatLabel } from "../../lib/format";
import { spreadsheetFileExtension } from "../../lib/spreadsheet/fileTypes";
import type { PackageVariant, SpreadsheetFormat } from "../../lib/spreadsheet/model";

/** Said under the panel's title: a restore is a replace, and its undo is the
 * version history wrote just before it. */
export const WORKBOOK_RESTORE_NOTE =
  "Restoring replaces the workbook and clears undo — the version saved just before is at the top of the list";

/** Said when a version's chain or one of its blobs is gone. */
export const VERSION_MISSING_MESSAGE = "This version can't be opened — part of it is missing";

/** The restore toast and its action. */
export const VERSION_RESTORED_MESSAGE = "Version restored";
export const UNDO_RESTORE_LABEL = "Undo restore";

function count(n: number, one: string, many: string): string {
  return `${n.toLocaleString("en-US")} ${n === 1 ? one : many}`;
}

/** "+12 cells", "−3 cells", "no change in cell count". */
export function cellDeltaLabel(delta: number): string {
  if (delta === 0) {
    return "no change in cell count";
  }
  return `${delta > 0 ? "+" : "−"}${count(Math.abs(delta), "cell", "cells")}`;
}

/** "3 sheets · 1,204 cells · +12 cells" — a version as the list shows it. */
export function workbookStatLabel(entry: VersionIndexEntry & { cellDelta?: number }): string {
  return `${count(entry.sheets ?? 0, "sheet", "sheets")} · ${count(entry.cells ?? 0, "cell", "cells")} · ${cellDeltaLabel(entry.cellDelta ?? 0)}`;
}

/** The version download's button and title: "Download .xlsx" — the
 * workbook's own format and variant. */
export function workbookDownloadLabels(format: SpreadsheetFormat, variant?: PackageVariant): { button: string; title: string } {
  return {
    button: `Download .${spreadsheetFileExtension(format, variant)}`,
    title: `Download this version as ${spreadsheetFormatLabel(format, variant)}`,
  };
}
