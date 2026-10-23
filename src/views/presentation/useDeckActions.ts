// The deck header's actions, lifted out of PresentationView so the view stays
// within its size budget (the same extraction rule as the document editor's
// composables). Two groups: the download pair (own format + "Download a copy
// as…" the other, SPEC-080) and playback entry (SPEC-084) — the Present button,
// its Cmd/Ctrl+Shift+Return shortcut, and the return-selection handoff the play
// view writes back via a `?return` query. SPEC-100 adds the presenter view to
// the same entry: its Present menu item and Cmd/Ctrl+Alt+Shift+Return. Framework state (the model, the
// current index) comes in as refs; nothing here renders.
import { computed, ref, type Ref } from "vue";
import type { Router, RouteLocationNormalizedLoaded } from "vue-router";
import { downloadPresentation } from "../../lib/presentation/download";
import { exportPresentationHtml, htmlFileName, mediaExportSummary } from "../../lib/presentation/htmlExport";
import { triggerBlobDownload } from "../../lib/core/triggerDownload";
import { presentationFormatLabel } from "../../lib/format";
import type { PresentationModel, PresentationFormat } from "../../lib/presentation/model";

const HTML_MIME_TYPE = "text/html;charset=utf-8";

/** How Present starts (SPEC-100): the full-screen show in this tab, or the
 * presenter view here driving an audience window. */
export type PresentMode = "show" | "presenter";

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

  // --- HTML export (SPEC-088) -----------------------------------------------
  // The pre-export options dialog opens on the menu item; the actual serialise
  // + download happens on confirm, from the model + media alone.
  const htmlDialogOpen = ref(false);

  function openHtmlExport(): void {
    if (deps.model.value && deps.model.value.slides.length > 0) {
      htmlDialogOpen.value = true;
    }
  }

  /** SPEC-095 AC-49: the dialog's media note — the count and reason when the
   * deck's clips exceed the 8 MB budget and will export as poster plates; "" when
   * nothing is omitted. */
  const htmlMediaNote = computed(() => {
    const model = deps.model.value;
    if (!model) {
      return "";
    }
    const summary = mediaExportSummary(model);
    if (!summary.omitted || summary.count === 0) {
      return "";
    }
    const clips = summary.count === 1 ? "1 media clip" : `${summary.count} media clips`;
    return `${clips} will be omitted from this export (over the 8 MB media budget) and shown as poster plates.`;
  });

  function cancelHtmlExport(): void {
    htmlDialogOpen.value = false;
  }

  function confirmHtmlExport(options: { includeNotes: boolean }): void {
    htmlDialogOpen.value = false;
    const model = deps.model.value;
    if (!model || model.slides.length === 0) return;
    const html = exportPresentationHtml(model, { includeNotes: options.includeNotes });
    triggerBlobDownload(new TextEncoder().encode(html), htmlFileName(model.title), HTML_MIME_TYPE);
  }

  function close(): void {
    // Mirrors the editor's close (useFileActions): the library scrolls to and
    // highlights the row that was just closed (SPEC-060).
    deps.router.push({ path: "/", state: { closedDocId: deps.presentationId.value } });
  }

  // --- playback entry (SPEC-084) --------------------------------------------
  const canPresent = computed(() => deps.slideTotal.value > 0);

  /** Start playback at the current slide (1-based `n` in the route): the
   * show (SPEC-084), or presenter view (SPEC-100). */
  function present(mode: PresentMode = "show"): void {
    if (!canPresent.value) return;
    const route = mode === "presenter" ? "presenter" : "present";
    void deps.router.push(`/slides/${deps.presentationId.value}/${route}/${deps.currentIndex.value + 1}`);
  }

  /** Cmd/Ctrl+Shift+Return starts playback (AC-1); adding Alt starts
   * presenter view instead (SPEC-100). */
  function onPresentShortcut(event: KeyboardEvent): void {
    if ((event.metaKey || event.ctrlKey) && event.shiftKey && event.key === "Enter") {
      event.preventDefault();
      present(event.altKey ? "presenter" : "show");
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
    htmlDialogOpen,
    htmlMediaNote,
    openHtmlExport,
    cancelHtmlExport,
    confirmHtmlExport,
    close,
    canPresent,
    present,
    onPresentShortcut,
    returnIndex,
  };
}
