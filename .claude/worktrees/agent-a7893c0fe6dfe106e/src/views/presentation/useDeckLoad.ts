// Loading (and legacy-upgrading) the deck for PresentationView (SPEC-096
// extraction — the view stays a thin wiring hub under its ratcheting line
// budget): owns the `presentation` shallowRef, the text-id redirect, the
// preparation-milestone upgrade on open, and the id-change reload.
import { shallowRef, watch, type ComputedRef, type Ref, type ShallowRef } from "vue";
import type { Router } from "vue-router";
import type { PresentationModel } from "../../lib/presentation/model";
import { documentKindOf, getPresentation, initLibrary, storePresentation } from "../../lib/documentStore";
import { isLegacyPresentationModel } from "../../lib/presentation/model";
import { upgradePresentationModel } from "../../lib/presentation/open";

export interface DeckLoadDeps {
  presentationId: ComputedRef<string>;
  router: Router;
  /** The slide the view should land on after a load (SPEC-084 `?return`). */
  returnIndex: () => number;
  currentIndex: Ref<number>;
  /** Resets the undo stack — per-open-view (SPEC-080 AC-45). */
  resetUndo: () => void;
}

export function useDeckLoad(deps: DeckLoadDeps): {
  presentation: ShallowRef<PresentationModel | null>;
  start: () => void;
} {
  // A legacy (preparation-milestone) model still holds outline-shaped slides,
  // which the frame renderer cannot draw — hide it until loadPresentation has
  // upgraded it, so the filmstrip/stage only ever see the frame model.
  function initialModel(): PresentationModel | null {
    const model = getPresentation(deps.presentationId.value) ?? null;
    return model && isLegacyPresentationModel(model) ? null : model;
  }

  const presentation = shallowRef<PresentationModel | null>(initialModel());

  async function loadPresentation(): Promise<void> {
    // A text document's id belongs to the word processor — hand it over.
    if (documentKindOf(deps.presentationId.value) === "text") {
      void deps.router.replace(`/doc/${deps.presentationId.value}`);
      return;
    }
    let model = getPresentation(deps.presentationId.value) ?? null;
    // Preparation-milestone records upgrade on open (SPEC-079 AC-54..56).
    if (model && isLegacyPresentationModel(model)) {
      try {
        const upgraded = await upgradePresentationModel(model);
        storePresentation(upgraded);
        model = getPresentation(deps.presentationId.value) ?? upgraded;
      } catch {
        // Reading stays tolerant: if the retained package cannot be re-parsed,
        // keep the record as-is rather than blanking the view.
      }
    }
    presentation.value = model;
    // Reselect the slide the play view was showing on exit, else the first.
    deps.currentIndex.value = deps.returnIndex();
    deps.resetUndo();
  }

  /** Kicks off the initial load; main.ts awaits initLibrary() before mounting,
   * so re-syncing after it resolves matters only when the view is mounted
   * directly (tests). Call AFTER the return-selection reader exists. */
  function start(): void {
    void loadPresentation();
    void initLibrary().then(loadPresentation);
    watch(deps.presentationId, () => {
      void loadPresentation();
    });
  }

  return { presentation, start };
}
