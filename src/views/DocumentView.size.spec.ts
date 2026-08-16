// DocumentView.vue is the editor's backbone and wiring hub — and the file
// the factory historically grew fastest: after the 2026-08-08 composable
// extraction cut it from 7,757 to 4,707 lines, twelve specs regrew it to
// 7,395 by 2026-08-15 (~380 lines/day), because nothing pushed back. This
// budget is that pushback — the same architecture-as-test pattern as
// lib/core's boundary.spec.ts.
//
// The rules (also in CLAUDE.md):
// - Feature code belongs in a src/views/document/ composable (its Deps
//   extending EditorContext from src/views/document/context.ts) and in
//   components; DocumentView gains only a `use*({ ...editorContext,
//   ...extras })` wiring call and its template hookup. The headroom under
//   the budget exists for exactly that wiring — a few dozen lines per
//   feature, not a feature's implementation.
// - When this test fails, extract code out of the view. Raising the budget
//   to make a change pass is weakening a test, which the testing rules
//   forbid.
// - When an extraction lands, ratchet DOWN: set the budget to the new line
//   count plus ~100 of wiring headroom, and append a row to the ledger.
//
// Ratchet ledger:
//   2026-08-15  7350  Phase 0 — shared EditorContext, applyEdit options
//   2026-08-15  7180  Phase 1a — useAutocorrect extracted
//   2026-08-15  6925  Phase 1b — useStyleManager extracted
//   2026-08-15  6655  Phase 1c — useContextMenus extracted
//   2026-08-15  6545  Phase 1d — useFileActions extracted
//   2026-08-15  6285  Phase 1e — page-setup/outline/merge trio extracted
//   2026-08-15  6190  Phase 1f — useViewportVirtualization extracted
//   2026-08-15  5685  Phase 1g — useFormattingSurfaces (triad collapse)
//   2026-08-15  5185  Phase 1h — DocumentMenuBar.vue extracted
//   2026-08-15  5025  Phase 1i — TableHandlesOverlay.vue extracted
import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";

const LINE_BUDGET = 5025;

describe("DocumentView size budget", () => {
  it(`DocumentView.vue stays within its ${LINE_BUDGET}-line budget`, () => {
    const viewPath = join(dirname(fileURLToPath(import.meta.url)), "DocumentView.vue");
    const lineCount = readFileSync(viewPath, "utf8").split("\n").length;
    expect(
      lineCount,
      `DocumentView.vue is ${lineCount} lines, over its ${LINE_BUDGET}-line budget. ` +
        "Extract feature code into a src/views/document/ composable " +
        "(see this spec's header comment) — do not raise the budget.",
    ).toBeLessThanOrEqual(LINE_BUDGET);
  });
});
