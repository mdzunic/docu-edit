// Download a presentation as a file — the presentation twin of
// lib/downloadDocument.ts, kept framework-free so the presentation view
// triggers it without mounting anything. In its own format the output is the
// retained source package re-emitted (or a from-scratch build for a deck with
// no package). "Download a copy as…" the other format converts through the
// model: `stripRetention` drops every retained-package field so the writers
// rebuild from the model alone — a cross-format save never leaks source-format
// parts (SPEC-080 AC-15). Frames the model can carry but the target format
// cannot are dropped and counted into the returned notice.
import { triggerBlobDownload } from "../core/triggerDownload";
import type { PresentationFormat, PresentationModel } from "./model";
import { stripRetention, unconvertibleFrameCount } from "./slides";
import { pptxFileName, PPTX_MIME_TYPE, writePptx } from "./pptx/writer";
import { ODP_MIMETYPE, odpFileName, writeOdp } from "./odp/writer";

/** Serialises a presentation and triggers a browser download. With no
 * `targetFormat` (or the deck's own format) it downloads natively; a
 * different `targetFormat` downloads a converted copy and does NOT mutate the
 * open model. Returns the number of objects that could not be converted (0
 * for a native download, or a deck of placeholders). */
export async function downloadPresentation(
  model: PresentationModel,
  targetFormat: PresentationFormat = model.format,
): Promise<{ droppedObjects: number }> {
  const crossFormat = targetFormat !== model.format;
  const source = crossFormat ? stripRetention({ ...model, format: targetFormat }) : model;
  const droppedObjects = crossFormat ? unconvertibleFrameCount(model) : 0;

  const { bytes } = targetFormat === "odp" ? await writeOdp(source) : await writePptx(source);
  const fileName = targetFormat === "odp" ? odpFileName(model.title) : pptxFileName(model.title);
  const mimeType = targetFormat === "odp" ? ODP_MIMETYPE : PPTX_MIME_TYPE;
  triggerBlobDownload(bytes, fileName, mimeType);
  return { droppedObjects };
}
