// PresentationPlayView.vue is the slideshow's wiring hub (SPEC-084), given a
// budget by SPEC-100 — the same architecture-as-test pattern as
// PresentationView.size.spec.ts, DocumentView.size.spec.ts and
// HomeView.size.spec.ts. Playback logic lives in src/views/presentation/
// composables (usePlayback, useBuildPreview, useShowingMedia, usePlayInput,
// useReducedMotion, usePresenterSession) and the lib/presentation modules they
// call, not in the view.
//
// The rules (mirroring PresentationView's):
// - Feature code belongs in a src/views/presentation/ composable or a
//   src/components/presentation/ component; PresentationPlayView keeps only
//   wiring (load/upgrade, the step and transition pipeline, URL sync, fit
//   scaling, exit, template).
// - When this test fails, extract code out of the view — do not raise the
//   budget (that weakens the test, which the testing rules forbid).
// - When an extraction lands, ratchet DOWN: set the budget to the new line
//   count plus a little wiring headroom, and append a row to the ledger.
//
// Ratchet ledger:
//   2026-10-23  600  SPEC-100 — first budget: the keyboard, pointer and idle-HUD
//                    handling (with the full-screen request it sits beside),
//                    the reduced-motion query and the build step watcher moved
//                    out into usePlayInput, useReducedMotion and
//                    watchBuildSteps, paying for the usePresenterSession
//                    follower call and the blank / full-screen-prompt markup
//                    (view at 590 lines, from 658)
import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";

const LINE_BUDGET = 600;

describe("PresentationPlayView size budget", () => {
  it(`PresentationPlayView.vue stays within its ${LINE_BUDGET}-line budget`, () => {
    const viewPath = join(dirname(fileURLToPath(import.meta.url)), "PresentationPlayView.vue");
    const lineCount = readFileSync(viewPath, "utf8").split("\n").length;
    expect(
      lineCount,
      `PresentationPlayView.vue is ${lineCount} lines, over its ${LINE_BUDGET}-line budget. ` +
        "Extract feature code into a src/views/presentation/ composable or a " +
        "src/components/presentation/ component (see this spec's header comment) — do not raise the budget.",
    ).toBeLessThanOrEqual(LINE_BUDGET);
  });
});
