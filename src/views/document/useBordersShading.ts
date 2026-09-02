// Format → Borders and shading… (SPEC-074 AC-6..AC-11): opens the dialog
// seeded from the first paragraph the selection touches, and applies what it
// returns to every paragraph the selection touches — SPEC-014's rule, which
// the applyParagraphBorders command implements.
//
// DocumentView keeps only the `useBordersShading({ ...editorContext, ... })`
// wiring call and the template hookup; every decision lives here or in
// lib/paragraphBorders.ts.
import { ref } from "vue";
import { applyParagraphBorders, type ModelSelection } from "../../lib/commands";
import { firstTouchedFormatting } from "../../lib/commands.paragraph";
import type { ParagraphBorders } from "../../lib/documentModel";
import type { EditorContext } from "./context";

/** What the dialog emits on OK — the `null`-removes shape the command takes. */
export interface BordersShadingResult {
  borders: ParagraphBorders | null;
  shadingFill: string | null;
}

export interface BordersShadingDeps extends EditorContext {
  /** DocumentView's paragraphSelection always resolves to a selection. */
  paragraphSelection: (root?: HTMLElement | null) => ModelSelection;
}

export function useBordersShading(deps: BordersShadingDeps) {
  const { doc, contentRoot, editable, applyEdit } = deps;

  const bordersDialogOpen = ref(false);

  /** The dialog's seed: the borders and shading of the first paragraph the
   * selection touches, so reopening it shows what is already applied.
   *
   * A plain ref written by openBordersDialog, NOT a computed: the seed is
   * derived from the DOM selection (`paragraphSelection` reads
   * `document.getSelection()` and the view's caret), which Vue cannot track,
   * so a computed would only re-run when `doc` or `contentRoot` changed.
   * Moving the caret between two opens without editing left it cached on
   * the previously inspected paragraph — and OK then stamped THAT
   * paragraph's borders onto the new selection. Same convention as
   * usePageSetupDialogs.openPageSetup, which seeds its prefill on open. */
  const bordersDialogInitial = ref<{ borders?: ParagraphBorders; shadingFill?: string }>({});

  /** The seed for the selection as it is right now. */
  function currentInitial(): { borders?: ParagraphBorders; shadingFill?: string } {
    const model = doc.value;
    if (!model) {
      return {};
    }
    const root = contentRoot.value as HTMLElement | null;
    const formatting = firstTouchedFormatting(model, deps.paragraphSelection(root));
    return { borders: formatting.borders, shadingFill: formatting.shadingFill };
  }

  function openBordersDialog(): void {
    if (!doc.value || !editable.value) {
      return;
    }
    // Seed at the moment of opening, from the live selection (see above).
    bordersDialogInitial.value = currentInitial();
    bordersDialogOpen.value = true;
  }

  function closeBordersDialog(): void {
    bordersDialogOpen.value = false;
  }

  function applyBordersShading(result = { borders: null, shadingFill: null } as BordersShadingResult): void {
    const model = doc.value;
    const root = contentRoot.value as HTMLElement | null;
    bordersDialogOpen.value = false;
    if (!model || !editable.value) {
      return;
    }
    const selection = deps.paragraphSelection(root);
    applyEdit(
      applyParagraphBorders(model, selection, {
        borders: result.borders,
        shadingFill: result.shadingFill,
      }),
      selection,
      "borders and shading",
    );
    root?.focus();
  }

  return { bordersDialogOpen, bordersDialogInitial, openBordersDialog, closeBordersDialog, applyBordersShading };
}
