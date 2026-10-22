// Download a stored document as a file (SPEC-067 bulk Download, AC-27). Reuses
// the SPEC-006/007 writers the editor's own download uses, so a document
// downloaded from the library round-trips identically to one downloaded from
// the editor. Kept framework-free (no Vue) so the library composable and any
// headless caller can trigger a download without mounting a component.
import type { DocumentModel } from "./documentModel";
import { triggerBlobDownload } from "./core/triggerDownload";
import { DOCX_MIME_TYPE, docxFileName, writeDocx } from "./docx/writer";
import { MD_MIME_TYPE, mdFileName, writeMd } from "./md/writer";
import { ODT_MIMETYPE, odtFileName, writeOdt } from "./odf/writer";

/** The concrete file format a document downloads as: its own saved format,
 * defaulting to .docx for a document created in the app ("new"). */
function downloadFormatOf(model: DocumentModel): "docx" | "odt" | "md" {
  if (model.format === "odt") {
    return "odt";
  }
  if (model.format === "md") {
    return "md";
  }
  return "docx";
}

/** Serialises a document and triggers a browser download of it in its native
 * format. Awaits the writer but not the download itself (a transient anchor
 * pointed at an object URL, the same mechanism useFileActions uses). */
export async function downloadDocument(model: DocumentModel): Promise<void> {
  const format = downloadFormatOf(model);
  const modified = new Date();
  const { bytes } =
    format === "odt"
      ? await writeOdt(model, { modified })
      : format === "md"
        ? await writeMd(model)
        : await writeDocx(model, { modified });
  const fileName =
    format === "odt"
      ? odtFileName(model.title)
      : format === "md"
        ? mdFileName(model.title)
        : docxFileName(model.title);
  const mimeType =
    format === "odt"
      ? ODT_MIMETYPE
      : format === "md"
        ? MD_MIME_TYPE
        : DOCX_MIME_TYPE;
  triggerBlobDownload(bytes, fileName, mimeType);
}
