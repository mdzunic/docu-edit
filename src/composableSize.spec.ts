// The composable size ratchet (ARCH-2), the sibling of the four
// *View.size.spec.ts budgets.
//
// Those four guard the views, and they work: DocumentView came down from
// 6,655 lines to 4,330 under theirs. But they only ever asked "is the VIEW
// small?", and the answer to that is yes whether the extracted code became a
// well-shaped composable or a second dumping ground. Nothing bounded the
// destination, so the pressure moved rather than dispersed:
// useShapes.ts reached 2,703 lines — larger than every view except
// DocumentView, larger than documentModel.ts, larger than the .pptx reader —
// and every line of it satisfied the rule it was extracted to satisfy.
//
// So the budget follows the code. Every `use*.ts` under src/views and
// src/composables is bounded here.
//
// The shape is a norm plus an exception list, rather than 95 individual
// budgets. DEFAULT_BUDGET is what a composable should be: at the time of
// writing the median is 208 lines and three quarters sit under 362, so 400 is
// generous rather than aspirational. OVER_BUDGET is then not a config table
// but a debt register — the 23 files that already exceed the norm, each
// pinned at the length it had the day this landed. It should only ever get
// shorter.
//
// The rules, matching the view budgets:
// - When this fails on growth, extract — splitting a composable, or moving
//   framework-free logic down into src/lib where it can be unit-tested
//   without a component. Raising a number to make a change pass is weakening
//   a test, which the testing rules forbid.
// - When an extraction lands, ratchet DOWN. That one is enforced here rather
//   than merely asked for: an entry left far above its file's real size is
//   headroom the next feature will silently spend, which is exactly how
//   useShapes grew. See the third test.
// - A new composable gets DEFAULT_BUDGET and no entry. Needing one is a
//   signal worth stopping on: it means the thing was born over the norm.
import { readdirSync, readFileSync, statSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";

/** What a composable should fit in. Not a target to grow into. */
const DEFAULT_BUDGET = 400;

/** How far an entry may sit above its file before the ratchet demands to be
 * tightened. Loose enough that ordinary edits don't churn the register,
 * tight enough that a real extraction has to be banked. */
const RATCHET_SLACK = 50;

/** The debt register: composables already over the norm when this landed,
 * pinned at that day's length. Paths are relative to src/. Entries come off
 * this list by shrinking, never by growing. */
const OVER_BUDGET: Record<string, number> = {
  "views/document/useShapes.ts": 2703,
  "views/document/usePaginationView.ts": 1347,
  "composables/usePagination.ts": 935,
  "views/document/useConnectors.ts": 853,
  "views/document/useFormattingSurfaces.ts": 688,
  "views/document/useTables.ts": 668,
  "views/presentation/useFrameManipulation.ts": 659,
  "views/document/useFindReplace.ts": 606,
  "views/presentation/useSlideTextEditing.ts": 601,
  "views/document/useCanvasGestures.ts": 591,
  "views/document/useSpellCheck.ts": 535,
  "views/document/useComments.ts": 533,
  "views/document/useLinks.ts": 526,
  "views/document/useHeadersFooters.ts": 495,
  "composables/useLibrary.ts": 489,
  "views/document/useCharts.ts": 489,
  "views/document/useCaptions.ts": 483,
  "views/document/useContextMenus.ts": 446,
  "views/document/useStyleManager.ts": 444,
  "views/document/useTableTools.ts": 425,
  "views/document/useNotes.ts": 415,
  "views/document/useFields.ts": 408,
  "composables/useAutosave.ts": 396,
};

const srcDir = dirname(fileURLToPath(import.meta.url));

/** Every composable under src/views and src/composables, as paths relative
 * to src/. A composable is a `use*.ts` — the naming convention CLAUDE.md
 * states, and the same one the imports rely on. */
function composables(): string[] {
  const found: string[] = [];
  const walk = (dir: string) => {
    for (const name of readdirSync(dir)) {
      const path = join(dir, name);
      if (statSync(path).isDirectory()) {
        walk(path);
      } else if (/^use[A-Z].*\.ts$/.test(name) && !name.endsWith(".spec.ts")) {
        found.push(path.slice(srcDir.length + 1));
      }
    }
  };
  walk(join(srcDir, "views"));
  walk(join(srcDir, "composables"));
  return found.toSorted();
}

/** Counted the way the view budgets count, so the two are comparable. */
function lineCount(relativePath: string): number {
  return readFileSync(join(srcDir, relativePath), "utf8").split("\n").length;
}

function budgetFor(relativePath: string): number {
  return OVER_BUDGET[relativePath] ?? DEFAULT_BUDGET;
}

describe("composable size budgets", () => {
  it("every composable stays within its budget", () => {
    const over = composables()
      .map((path) => ({ path, lines: lineCount(path), budget: budgetFor(path) }))
      .filter((f) => f.lines > f.budget)
      .map((f) =>
        OVER_BUDGET[f.path] === undefined
          ? `${f.path} is ${f.lines} lines, over the ${DEFAULT_BUDGET}-line default`
          : `${f.path} is ${f.lines} lines, over its ${f.budget}-line entry`,
      );
    expect(
      over,
      "Composables over budget. Extract — split the composable, or move " +
        "framework-free logic down into src/lib where it is unit-testable " +
        "without mounting a component. Do not raise a number to make a " +
        "change pass; see this spec's header.",
    ).toEqual([]);
  });

  it("the debt register has no entry for a file that no longer exists", () => {
    // A stale entry is a budget nobody is spending and nobody can see is
    // gone — and it silently pre-approves the size if the name comes back.
    const present = new Set(composables());
    const stale = Object.keys(OVER_BUDGET).filter((path) => !present.has(path));
    expect(
      stale,
      "OVER_BUDGET names files that are no longer composables (renamed, " +
        "moved or deleted). Remove these entries.",
    ).toEqual([]);
  });

  it("no entry sits far above the file it bounds", () => {
    // This is what makes it a ratchet rather than a ceiling. Slack left
    // after an extraction is exactly the headroom the next feature spends
    // without anyone deciding to spend it.
    const slack = Object.keys(OVER_BUDGET)
      .filter((path) => composables().includes(path))
      .map((path) => ({ path, lines: lineCount(path), budget: OVER_BUDGET[path] }))
      .filter((f) => f.budget - f.lines > RATCHET_SLACK)
      .map(
        (f) =>
          `${f.path} is ${f.lines} lines but its entry says ${f.budget} — ` +
          `lower it to ${f.lines}` +
          (f.lines <= DEFAULT_BUDGET ? " (or drop the entry: it is under the default now)" : ""),
      );
    expect(
      slack,
      "These composables have shrunk. Bank it: lower the entry to the new " +
        "count so the space cannot be refilled without a decision.",
    ).toEqual([]);
  });
});
