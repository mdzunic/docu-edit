// PresenterView.vue is the presenter view's wiring hub (SPEC-100) — the same
// architecture-as-test pattern as PresentationView.size.spec.ts,
// DocumentView.size.spec.ts and HomeView.size.spec.ts. The presenter window's
// logic lives in src/views/presentation/ composables (usePresenterView, which
// builds usePresenterSession, useAudienceWindow and usePresenterTools) and its
// UI in src/components/presentation/ (PresenterSlidePane, PresenterNotes,
// PresenterGrid), not in the view.
//
// The rules (mirroring PresentationView's):
// - Feature code belongs in a src/views/presentation/ composable or a
//   src/components/presentation/ component; PresenterView keeps only wiring
//   (the usePresenterView call, the template and its layout styles).
// - When this test fails, extract code out of the view — do not raise the
//   budget (that weakens the test, which the testing rules forbid).
// - When an extraction lands, ratchet DOWN: set the budget to the new line
//   count plus a little wiring headroom, and append a row to the ledger.
//
// Ratchet ledger:
//   2026-10-23  400  SPEC-100 — first budget (view at 379 lines)
import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";

const LINE_BUDGET = 400;

describe("PresenterView size budget", () => {
  it(`PresenterView.vue stays within its ${LINE_BUDGET}-line budget`, () => {
    const viewPath = join(dirname(fileURLToPath(import.meta.url)), "PresenterView.vue");
    const lineCount = readFileSync(viewPath, "utf8").split("\n").length;
    expect(
      lineCount,
      `PresenterView.vue is ${lineCount} lines, over its ${LINE_BUDGET}-line budget. ` +
        "Extract feature code into a src/views/presentation/ composable or a " +
        "src/components/presentation/ component (see this spec's header comment) — do not raise the budget.",
    ).toBeLessThanOrEqual(LINE_BUDGET);
  });
});
