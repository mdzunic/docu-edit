// The file-level actions on the open document: the app bar's format badge
// and title rename (SPEC-003), the download flow (SPEC-006/007's writers,
// SPEC-050's markdown, and SPEC-034 AC-14's PDF branch), and the File menu's
// Save as template (SPEC-046) and Delete document (SPEC-003). A download is
// the app's explicit save: it stamps the created/modified properties
// (SPEC-041 AC-2), promotes the autosave state to committed and notes a
// version-history commit point (SPEC-040 AC-2), then reports content the
// save could not preserve (SPEC-008, widened by SPEC-052).
import { computed, ref, triggerRef } from "vue";
import { onBeforeRouteLeave, useRouter } from "vue-router";
import { STORAGE_FULL_REASON } from "../../composables/useAutosave";
import { usePreferences } from "../../composables/usePreferences";
import { DEFAULT_TITLE } from "../../lib/documentModel";
import { isStorageAvailable, touchDocument, trashDocument } from "../../lib/documentStore";
import { DOCX_MIME_TYPE, docxFileName, writeDocx } from "../../lib/docx/writer";
import { MD_MIME_TYPE, mdFileName, writeMd } from "../../lib/md/writer";
import { ODT_MIMETYPE, odtFileName, writeOdt } from "../../lib/odf/writer";
import { noConversionReport } from "../../lib/opaque";
import { saveAsTemplate } from "../../lib/templateStore";
import type { EditorContext } from "./context";

export interface FileActionsDeps extends EditorContext {
  /** usePdfExport's exportPdf — "Download as PDF" runs its own rendered-
   * output pipeline (SPEC-034 AC-14), never the writers below. */
  exportPdf: () => Promise<void>;
  /** useAutosave's commitNow — a download promotes the session state to
   * committed (SPEC-003's explicit save). */
  commitNow: () => Promise<void>;
  /** useVersionHistory's noteCommitPoint — the download is SPEC-040 AC-2's
   * other commit-point version-history trigger. */
  noteVersionHistoryCommitPoint: () => void;
  /** useAutosave's flushForClose — closing a document is a commit boundary
   * (SPEC-060): cancel the pending autosave debounce, await its write, and
   * only navigate when it resolves. Returns false on a write failure so the
   * close stays put. */
  flushForClose: () => Promise<boolean>;
}

export function useFileActions(deps: FileActionsDeps) {
  const { doc, addToast, exportPdf, commitNow, noteVersionHistoryCommitPoint } = deps;
  // Deleting the open document navigates back to the library (SPEC-003).
  const router = useRouter();
  // The "Default download format" preference (SPEC-066): "source" keeps the
  // pre-spec behaviour (mirror the document's own format); a fixed choice
  // makes the one-click "Download a copy" escape hatches (the save-error link,
  // the close-without-saving dialog) save in that format.
  const { preferences } = usePreferences();

  const formatBadge = computed(() => doc.value?.format?.toUpperCase() ?? "");

  const downloadFormat = computed(() => {
    const preferred = preferences.defaultDownloadFormat;
    if (preferred === "docx" || preferred === "odt") {
      return preferred;
    }
    if (!doc.value) {
      return "docx";
    }
    if (doc.value.format === "odt") {
      return "odt";
    }
    if (doc.value.format === "md") {
      return "md";
    }
    return "docx";
  });

  function handleTitleUpdate(newTitle = "") {
    if (!doc.value) {
      return;
    }
    const trimmed = newTitle.trim();
    doc.value.title = trimmed.length > 0 ? trimmed : DEFAULT_TITLE;
    triggerRef(doc);
    // The assignment above already mutated the stored model (doc.value IS the
    // raw stored model); touching records the modification time and persists.
    touchDocument(doc.value.id);
  }

  async function handleDownload(format = "") {
    // PDF export (SPEC-034 AC-14): a rendered-output snapshot, not a save —
    // it runs its own pipeline and never touches autosave or the model.
    if (format === "pdf") {
      void exportPdf();
      return;
    }
    if (!doc.value || (format !== "docx" && format !== "odt" && format !== "md")) {
      return;
    }
    // The dcterms:created/dc:date (SPEC-041 AC-2) stamp: `created` is set only
    // the first time (kept stable on every later save, matching what General
    // tab shows once it has a value), `modified` refreshes every save. Both
    // .docx/.odt writers read `model.properties.created`/`modified` — see
    // docProps.ts/meta.ts's buildCorePropertiesXml/buildMetaXml — so this
    // must be resolved into the model *before* writing, then mirrored back so
    // the in-memory model (and so the Properties dialog, without a reopen)
    // agrees with what was just written. Markdown never reads `properties`
    // (SPEC-050: front matter is carried verbatim, never interpreted), but the
    // stamp is still recorded so a later save as .docx/.odt keeps it.
    const modified = new Date();
    const properties = {
      ...doc.value.properties,
      created: doc.value.properties?.created ?? modified.toISOString(),
      modified: modified.toISOString(),
    };
    doc.value.properties = properties;
    const { bytes, dropped } =
      format === "odt"
        ? await writeOdt(doc.value, { modified })
        : format === "md"
          ? await writeMd(doc.value)
          : await writeDocx(doc.value, { modified });
    triggerRef(doc);
    touchDocument(doc.value.id);
    const fileName =
      format === "odt" ? odtFileName(doc.value.title) : format === "md" ? mdFileName(doc.value.title) : docxFileName(doc.value.title);
    // A transient anchor pointed at an object URL is the download mechanism
    // that needs no new browser permissions; revocation is deferred so the
    // browser has started streaming the blob before the URL disappears.
    const mimeType = format === "odt" ? ODT_MIMETYPE : format === "md" ? MD_MIME_TYPE : DOCX_MIME_TYPE;
    const url = URL.createObjectURL(new Blob([bytes], { type: mimeType }));
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = fileName;
    anchor.click();
    setTimeout(() => URL.revokeObjectURL(url), 10_000);
    // A download is SPEC-003's explicit save: promote the state to committed,
    // and SPEC-040's other commit-point version-history trigger (AC-2).
    void commitNow().then(() => noteVersionHistoryCommitPoint());
    // SPEC-008: content that could not be preserved by a cross-format save is
    // reported once, after the download has been triggered, unless the user
    // opted out for this document. SPEC-052 widens this to a same-format save
    // too, but only ever for an edited chart's lost properties (chartFeature)
    // — "Conversion report" is the wrong heading there, since nothing
    // converted; conversionReportSameFormat drives the dialog's own variant.
    // Gated on both the app-level "Show the conversion report" preference
    // (SPEC-066) and the per-document "don't show again" (SPEC-008).
    if (dropped.length > 0 && preferences.showConversionReport && !doc.value.suppressConversionReport) {
      conversionReportSameFormat.value = doc.value.format === format;
      conversionReport.value = dropped;
    }
  }

  const conversionReport = ref(noConversionReport());
  const conversionReportSameFormat = ref(false);

  function closeConversionReport(dontShowAgain = false) {
    conversionReport.value = null;
    if (dontShowAgain && doc.value) {
      doc.value.suppressConversionReport = true;
      touchDocument(doc.value.id);
    }
  }

  const confirmingDelete = ref(false);

  const savingAsTemplate = ref(false);

  function saveCurrentAsTemplate(name = "") {
    savingAsTemplate.value = false;
    if (!doc.value) {
      return;
    }
    const record = saveAsTemplate(doc.value, name);
    addToast(`Saved template "${record.title}"`);
  }

  // Closing the document (SPEC-060). Every exit from the editor — the app-bar
  // Close control, File → Close, the product-name home link and the browser
  // Back button — is the same commit boundary, and it runs through ONE flush
  // implementation: flushAndDecide. The control and the menu command call it
  // then navigate; the route guard calls it for the exits that navigate on
  // their own (Back, the home link), so the deliberate exit and the accidental
  // one are equally safe.
  const closing = ref(false);
  const storageWarningOpen = ref(false);
  // Set just before a navigation that has already been decided: the Close
  // control (which flushed before pushing), deletion (record already removed)
  // and "Close anyway" (loss deliberately accepted) all skip the guard's flush.
  let bypassCloseGuard = false;

  /** The single flush gate: cancels the pending autosave debounce, awaits its
   * committed write and reports whether it is safe to leave. Storage being
   * unavailable is the one case that cannot flush — it raises the confirmation
   * dialog and refuses instead (AC-16/17). A write failure keeps the document
   * open with an error toast and the save indicator's download escape hatch
   * (AC-12..15). While in flight the Close control is busy (AC-9). */
  async function flushAndDecide(): Promise<boolean> {
    if (!doc.value) {
      return true;
    }
    if (!isStorageAvailable()) {
      storageWarningOpen.value = true;
      return false;
    }
    closing.value = true;
    try {
      const flushed = await deps.flushForClose();
      if (!flushed) {
        addToast(
          `Couldn't close "${doc.value.title}" — ${STORAGE_FULL_REASON}. ` +
            "Your document is still open; retry, or download a copy.",
        );
        return false;
      }
      return true;
    } finally {
      closing.value = false;
    }
  }

  // The route guard: navigations that leave /doc/:id on their own — browser
  // Back and the product-name home link — take the same flush path. Returning
  // false cancels the navigation. Navigations already decided upstream (the
  // Close control, deletion, "Close anyway") set the bypass and skip it.
  onBeforeRouteLeave(() => {
    if (bypassCloseGuard) {
      bypassCloseGuard = false;
      return true;
    }
    return flushAndDecide();
  });

  /** The Close control and File → Close: flush first, and only on success
   * return to the library — carrying the closed id in router state (not the
   * URL, where a document id would be a needless privacy leak) so the library
   * can scroll to and highlight the row. */
  async function closeOpenDocument(): Promise<void> {
    if (!doc.value || closing.value) {
      return;
    }
    const id = doc.value.id;
    if (!(await flushAndDecide())) {
      return;
    }
    // Already flushed; the guard need not repeat it for this navigation.
    bypassCloseGuard = true;
    router.push({ path: "/", state: { closedDocId: id } });
  }

  /** The storage-unavailable dialog's "Close anyway": leave despite the loss,
   * skipping the flush that cannot succeed. */
  function confirmCloseAnyway() {
    const id = doc.value?.id;
    storageWarningOpen.value = false;
    bypassCloseGuard = true;
    router.push(id ? { path: "/", state: { closedDocId: id } } : "/");
  }

  function deleteOpenDocument() {
    if (!doc.value) {
      return;
    }
    const { id, title } = doc.value;
    confirmingDelete.value = false;
    // SPEC-067: the library's delete moves to the trash rather than erasing;
    // File → Delete follows suit, keeping its toast and its navigation. The
    // record is kept intact (id, model, version history) so Trash can restore
    // it; only "Delete permanently" and "Empty trash" call deleteDocument.
    trashDocument(id);
    addToast(`Deleted ${title}`);
    // The record is trashed, not saved; skip the leave-guard's autosave flush.
    bypassCloseGuard = true;
    router.push("/");
  }

  // The File -> Export command ids the menu bar and command palette dispatch
  // through (SPEC-006/007/050/034) — literal tuples, one per format (the same
  // list as the menu's EXPORT_FORMATS), because commandCatalogue.spec.ts reads
  // this block from source to hold the catalogue and the registry together.
  const commands: Array<[string, () => void]> = [
    ["file.export-docx", () => { void handleDownload("docx"); }],
    ["file.export-odt", () => { void handleDownload("odt"); }],
    ["file.export-md", () => { void handleDownload("md"); }],
    ["file.export-pdf", () => { void handleDownload("pdf"); }],
  ];

  return {
    commands,
    closing,
    storageWarningOpen,
    closeOpenDocument,
    confirmCloseAnyway,
    formatBadge,
    downloadFormat,
    conversionReport,
    conversionReportSameFormat,
    confirmingDelete,
    savingAsTemplate,
    handleTitleUpdate,
    handleDownload,
    closeConversionReport,
    saveCurrentAsTemplate,
    deleteOpenDocument,
  };
}
