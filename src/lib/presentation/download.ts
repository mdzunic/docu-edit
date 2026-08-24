// Download a presentation as a file — the presentation twin of
// lib/downloadDocument.ts, kept framework-free so the library composable
// and the presentation view trigger it without mounting anything. The
// output is the retained source package re-emitted (see the writers), so a
// downloaded presentation carries exactly what was opened.
import { triggerBlobDownload } from "../core/triggerDownload";
import type { PresentationModel } from "./model";
import { pptxFileName, PPTX_MIME_TYPE, writePptx } from "./pptx/writer";
import { ODP_MIMETYPE, odpFileName, writeOdp } from "./odp/writer";

/** Serialises a presentation and triggers a browser download of it in its
 * native format. */
export async function downloadPresentation(model: PresentationModel): Promise<void> {
  const { bytes } = model.format === "odp" ? await writeOdp(model) : await writePptx(model);
  const fileName = model.format === "odp" ? odpFileName(model.title) : pptxFileName(model.title);
  const mimeType = model.format === "odp" ? ODP_MIMETYPE : PPTX_MIME_TYPE;
  triggerBlobDownload(bytes, fileName, mimeType);
}
