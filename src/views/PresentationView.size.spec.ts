// PresentationView.vue is the presentation editor's wiring hub — the same
// architecture-as-test pattern as DocumentView.size.spec.ts and
// HomeView.size.spec.ts. SPEC-080 introduces the presentation-side composable
// directory (src/views/presentation/*: usePresentationUndo, useSlideCommands)
// and the SlideFilmstrip management UI, so feature logic lives there and in
// components, not in the view.
//
// The rules (mirroring DocumentView's):
// - Feature code belongs in a src/views/presentation/ composable or a
//   src/components/presentation/ component; PresentationView keeps only wiring
//   (load/upgrade, current-slide state, fit scaling, command/undo/download
//   hookup, template).
// - When this test fails, extract code out of the view — do not raise the
//   budget (that weakens the test, which the testing rules forbid).
// - When an extraction lands, ratchet DOWN: set the budget to the new line
//   count plus a little wiring headroom, and append a row to the ledger.
//
// Ratchet ledger:
//   2026-08-25  560  SPEC-080 — first budget (view at 542 lines)
//   2026-08-26  550  SPEC-086 — extracted useSlideRender (commit/refresh) to make
//                    room for the Animations panel wiring (view at 543 lines)
import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";

const LINE_BUDGET = 550;

describe("PresentationView size budget", () => {
  it(`PresentationView.vue stays within its ${LINE_BUDGET}-line budget`, () => {
    const viewPath = join(dirname(fileURLToPath(import.meta.url)), "PresentationView.vue");
    const lineCount = readFileSync(viewPath, "utf8").split("\n").length;
    expect(
      lineCount,
      `PresentationView.vue is ${lineCount} lines, over its ${LINE_BUDGET}-line budget. ` +
        "Extract feature code into a src/views/presentation/ composable or a " +
        "src/components/presentation/ component (see this spec's header comment) — do not raise the budget.",
    ).toBeLessThanOrEqual(LINE_BUDGET);
  });
});
