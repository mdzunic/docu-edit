// MediaView.vue is the media viewer's wiring hub (SPEC-107) — the same
// architecture-as-test pattern as DocumentView.size.spec.ts and
// HomeView.size.spec.ts: feature logic belongs in a src/views/media/
// composable or a components/media/ component; the view keeps only wiring.
//
// The rules (mirroring the other two):
// - When this test fails, extract code out of the view. Raising the budget
//   to make a change pass is weakening a test, which the testing rules
//   forbid.
// - When an extraction lands, ratchet DOWN: set the budget to the new line
//   count plus modest wiring headroom, and append a row to the ledger.
//
// Ratchet ledger:
//   2026-08-30  170  Opening entry: the view lands at 124 lines
import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";

const LINE_BUDGET = 170;

describe("MediaView size budget", () => {
  it(`MediaView.vue stays within its ${LINE_BUDGET}-line budget`, () => {
    const viewPath = join(dirname(fileURLToPath(import.meta.url)), "MediaView.vue");
    const lineCount = readFileSync(viewPath, "utf8").split("\n").length;
    expect(
      lineCount,
      `MediaView.vue is ${lineCount} lines, over its ${LINE_BUDGET}-line budget. ` +
        "Extract logic into a src/views/media/ composable or a " +
        "components/media/ component — do not raise the budget.",
    ).toBeLessThanOrEqual(LINE_BUDGET);
  });
});
