// HomeView.vue is the landing page's wiring hub — extracted to the library
// module (src/views/library/* + useLibrary + lib/libraryView) during the
// presentations preparation, the same architecture-as-test pattern as
// DocumentView.size.spec.ts and lib/core's boundary.spec.ts.
//
// The rules (mirroring DocumentView's):
// - Library UI belongs in a src/views/library/ component (product-neutral:
//   it renders entries and emits intents — it never knows what a "docx" or
//   a "presentation" is beyond the entry's fields); shared logic in
//   useLibrary/lib. HomeView keeps only product wiring: create/open
//   actions, kind routing, dialogs, toasts and the keyboard handler.
// - When this test fails, extract code out of the view. Raising the budget
//   to make a change pass is weakening a test, which the testing rules
//   forbid.
// - When an extraction lands, ratchet DOWN: set the budget to the new line
//   count plus ~100 of wiring headroom, and append a row to the ledger.
//
// Ratchet ledger:
//   2026-08-24  751  Library module extraction (1,476 -> 651)
import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";

const LINE_BUDGET = 751;

describe("HomeView size budget", () => {
  it(`HomeView.vue stays within its ${LINE_BUDGET}-line budget`, () => {
    const viewPath = join(dirname(fileURLToPath(import.meta.url)), "HomeView.vue");
    const lineCount = readFileSync(viewPath, "utf8").split("\n").length;
    expect(
      lineCount,
      `HomeView.vue is ${lineCount} lines, over its ${LINE_BUDGET}-line budget. ` +
        "Extract library UI into a src/views/library/ component " +
        "(see this spec's header comment) — do not raise the budget.",
    ).toBeLessThanOrEqual(LINE_BUDGET);
  });
});
