// SpreadsheetView.vue is the workbook view's wiring hub (SPEC-116) — the same
// architecture-as-test pattern as DocumentView.size.spec.ts,
// HomeView.size.spec.ts and MediaView.size.spec.ts: feature logic belongs in
// a src/views/spreadsheet/ composable or a components/spreadsheet/ component;
// the view keeps only wiring.
//
// The rules (mirroring the others):
// - When this test fails, extract code out of the view. Raising the budget
//   to make a change pass is weakening a test, which the testing rules
//   forbid.
// - When an extraction lands, ratchet DOWN: set the budget to the new line
//   count plus modest wiring headroom, and append a row to the ledger.
//
// Ratchet ledger:
//   2026-10-09  107  SPEC-116 opening entry: the ledger opens at the count the
//                    view lands on, measured as this test measures it — 106
//                    lines plus the final newline — with no headroom
//   2026-10-09   99  SPEC-121: the not-found state moved into
//                    WorkbookNotFound.vue and the editing wiring arrived as
//                    one useWorkbookEditing call (106 -> 98 lines)
//   2026-10-11   97  SPEC-126: the app bar's wiring moved into
//                    WorkbookAppBar.vue as the CSV import bar and dialog
//                    arrived (one useCsvImport call, two hookups) — 98 -> 96
//                    lines, measured 97 with the final newline, no headroom
//   2026-10-11   93  SPEC-128: the status bar's wiring moved into
//                    WorkbookStatusBar.vue as sort and filter arrived (one
//                    useSortFilter call; its surfaces ride the grid context)
//                    — 96 -> 92 lines, measured 93 with the final newline, no
//                    headroom
import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";

const LINE_BUDGET = 93;

describe("SpreadsheetView size budget", () => {
  it(`SpreadsheetView.vue stays within its ${LINE_BUDGET}-line budget`, () => {
    const viewPath = join(dirname(fileURLToPath(import.meta.url)), "SpreadsheetView.vue");
    const lineCount = readFileSync(viewPath, "utf8").split("\n").length;
    expect(
      lineCount,
      `SpreadsheetView.vue is ${lineCount} lines, over its ${LINE_BUDGET}-line budget. ` +
        "Extract logic into a src/views/spreadsheet/ composable or a " +
        "components/spreadsheet/ component — do not raise the budget.",
    ).toBeLessThanOrEqual(LINE_BUDGET);
  });
});
