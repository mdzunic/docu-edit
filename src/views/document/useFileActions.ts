// The file-level actions on the open document: the app bar's format badge
// and title rename (SPEC-003), the download flow (SPEC-006/007's writers,
// SPEC-050's markdown, and SPEC-034 AC-14's PDF branch), and the File menu's
// Save as template (SPEC-046) and Delete document (SPEC-003). A download is
// the app's explicit save: it stamps the created/modified properties
// (SPEC-041 AC-2), promotes the autosave state to committed and notes a
// version-history commit point (SPEC-040 AC-2), then reports content the
// save could not preserve (SPEC-008, widened by SPEC-052).
import { computed, ref, triggerRef } from "vue";
import { useRouter } from "vue-router";
import { DEFAULT_TITLE } from "../../lib/documentModel";
import { deleteDocument, touchDocument } from "../../lib/documentStore";
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
}

export function useFileActions(deps: FileActionsDeps) {
  const { doc, addToast, exportPdf, commitNow, noteVersionHistoryCommitPoint } = deps;
  // Deleting the open document navigates back to the library (SPEC-003).
  const router = useRouter();

  const formatBadge = computed(() => doc.value?.format?.toUpperCase() ?? "");

  const downloadFormat = computed(() => {
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
    if (dropped.length > 0 && !doc.value.suppressConversionReport) {
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

  function deleteOpenDocument() {
    if (!doc.value) {
      return;
    }
    const { id, title } = doc.value;
    confirmingDelete.value = false;
    deleteDocument(id);
    addToast(`Deleted ${title}`);
    router.push("/");
  }

  return {
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
