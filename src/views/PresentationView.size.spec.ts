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
//   2026-08-28  525  SPEC-089 — extracted SlideStageArea (.stage-workspace) to
//                    pay for the layout gallery wiring (view at 511 lines)
//   2026-08-28  455  SPEC-090 — extracted DeckHeader (.deck-header) as the deck-
//                    setup dialog + useDeckSetup wiring landed (view at 443 lines)
//   2026-08-29  450  SPEC-091 — extracted SlidePanelStack (the four below-stage
//                    panels) to pay for the useTables wiring (view at 448 lines)
//   2026-08-29  445  SPEC-092 — extracted PresentationDialogs (the trailing modal
//                    dialogs) and useFilmstripCommands (the filmstrip intent
//                    dispatch) to pay for the useSlideCharts + SlideChartLayer
//                    chart wiring (view at 442 lines)
//   2026-08-31  430  SPEC-093 — extracted DeckBody (the filmstrip + stage +
//                    panel-stack workspace body) to pay for the Header & Footer
//                    wiring (useHeaderFooter + HeaderFooterDialog + toolbar
//                    button + the deckFields resolver) (view at 426 lines)
//   2026-09-01  420  SPEC-094 — extracted useEditingPanels (the transition/
//                    animation/table/chart controller wiring; the stage-fit
//                    code the spec names for extraction already lives in
//                    SlideStageArea since SPEC-089, so this is its sanctioned
//                    equivalent) to pay for the Design-panel wiring
//                    (useDeckTheme + ThemesPanel + toolbar toggle)
//                    (view at 399 lines)
//   2026-09-01  390  SPEC-096 — extracted the .stage-column block (which had
//                    moved into DeckBody with SPEC-093) into SlideStageColumn
//                    and the deck load/upgrade flow into useDeckLoad, paying
//                    for the Align/Group/Position insert-bar hookups; the
//                    spec mandates a drop of at least 30 (view at 377 lines)
//   2026-09-01  380  SPEC-095 — extracted the window-keydown onMounted/
//                    onBeforeUnmount block into useWindowKeys, paying for the
//                    Media panel wiring (useSlideMedia via useEditingPanels, the
//                    <MediaPanel> mount, the slides-media toolbar toggle and the
//                    insert-media hookup); the spec mandates a drop of at least
//                    10 (view at 374 lines)
//   2026-10-08  365  SPEC-097 — extracted the ribbon strips' PresentationToolbar
//                    and SlideInsertBar prop/emit blocks into SlideToolbarHost
//                    (one `<SlideToolbarHost strip=… />` per ribbon slot),
//                    paying for the useSlideTextDepth wiring call, the
//                    <TextPanel> mount and the Text toggle prop pair; the spec
//                    mandates a drop of at least 10 (view at 360 lines)
//   2026-10-08  335  SPEC-099 — extracted the deck-navigation block (slide
//                    commands, derived facts, Add slide, the filmstrip
//                    command/reorder dispatch) into useDeckNavigation and the
//                    "not found" state into PresentationNotFound, paying for
//                    the useSlideSorter wiring call, the <SlideSorter> /
//                    <ReuseSlidesDialog> mounts and the Sorter toggle prop
//                    pair; the spec mandates a drop of at least 15
//                    (view at 329 lines)
//   2026-10-22  295  SPEC-104 — extracted the ribbon block (SlideRibbon's
//                    toggles, the menu bar and both SlideToolbarHost strips)
//                    into PresentationRibbon — the stage-fit code the spec
//                    names already lives in SlideStageArea (see SPEC-094's
//                    row), so this is its sanctioned equivalent — paying for
//                    the useDeckExport wiring call, the Ctrl+P shortcut and
//                    the deck-export prop pair (view at 289 lines)
import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";

const LINE_BUDGET = 295;

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
