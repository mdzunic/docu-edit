// The deck header's actions, lifted out of PresentationView so the view stays
// within its size budget (the same extraction rule as the document editor's
// composables). Two groups: the download pair (own format + "Download a copy
// as…" the other, SPEC-080) and playback entry (SPEC-084) — the Present button,
// its Cmd/Ctrl+Shift+Return shortcut, and the return-selection handoff the play
// view writes back via a `?return` query. Framework state (the model, the
// current index) comes in as refs; nothing here renders.
import { computed, type Ref } from "vue";
import type { Router, RouteLocationNormalizedLoaded } from "vue-router";
import { downloadPresentation } from "../../lib/presentation/download";
import { presentationFormatLabel } from "../../lib/format";
import type { PresentationModel, PresentationFormat } from "../../lib/presentation/model";

export interface UseDeckActionsDeps {
  model: Ref<PresentationModel | null>;
  presentationId: Ref<string>;
  currentIndex: Ref<number>;
  slideTotal: Ref<number>;
  router: Router;
  route: RouteLocationNormalizedLoaded;
  reportError: (message: string) => void;
}

export function useDeckActions(deps: UseDeckActionsDeps) {
  const otherFormat = computed<PresentationFormat>(() => (deps.model.value?.format === "odp" ? "pptx" : "odp"));
  const otherFormatLabel = computed(() => presentationFormatLabel(otherFormat.value));

  function download(): void {
    if (deps.model.value) {
      void downloadPresentation(deps.model.value);
    }
  }

  async function downloadCopy(): Promise<void> {
    const model = deps.model.value;
    if (!model) return;
    const { droppedObjects } = await downloadPresentation(model, otherFormat.value);
    if (droppedObjects > 0) {
      deps.reportError(`${droppedObjects} object${droppedObjects === 1 ? "" : "s"} couldn't be converted`);
    }
  }

  function close(): void {
    // Mirrors the editor's close (useFileActions): the library scrolls to and
    // highlights the row that was just closed (SPEC-060).
    deps.router.push({ path: "/", state: { closedDocId: deps.presentationId.value } });
  }

  // --- playback entry (SPEC-084) --------------------------------------------
  const canPresent = computed(() => deps.slideTotal.value > 0);

  /** Start playback at the current slide (1-based `n` in the route). */
  function present(): void {
    if (!canPresent.value) return;
    void deps.router.push(`/slides/${deps.presentationId.value}/present/${deps.currentIndex.value + 1}`);
  }

  /** Cmd/Ctrl+Shift+Return starts playback (AC-1). */
  function onPresentShortcut(event: KeyboardEvent): void {
    if ((event.metaKey || event.ctrlKey) && event.shiftKey && event.key === "Enter") {
      event.preventDefault();
      present();
    }
  }

  /** The 0-based slide the editor should reselect on open, from the play view's
   * `?return` handoff (AC-35); defaults to the first slide. Clamped to range. */
  function returnIndex(): number {
    const raw = Number(deps.route.query.return ?? "");
    if (!Number.isFinite(raw) || raw < 1) return 0;
    return Math.min(Math.trunc(raw) - 1, Math.max(0, deps.slideTotal.value - 1));
  }

  return {
    otherFormatLabel,
    download,
    downloadCopy,
    close,
    canPresent,
    present,
    onPresentShortcut,
    returnIndex,
  };
}
