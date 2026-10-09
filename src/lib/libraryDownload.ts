// The library's download dispatch (SPEC-067's row and bulk Download): each
// stored document downloads through its own product's writer, by the kind
// the registry gives its record (documentKinds.ts) — a deck's and a
// workbook's codecs loaded on demand, off the library page's module graph.
// It lives here, in the app layer where the products compose, so the
// library composable keeps no knowledge of what any kind is (SPEC-116 added
// the workbook branch).
import { documentKindOf, getDocument, getMedia, getPresentation, getSpreadsheet } from "./documentStore";
import { downloadDocument } from "./downloadDocument";
import { downloadMedia } from "./media/download";

/** Downloads the stored document `id` in its own format; an id that names
 * nothing downloads nothing. */
export async function downloadStoredDocument(id: string): Promise<void> {
  const kind = documentKindOf(id);
  if (kind === "presentation") {
    const presentation = getPresentation(id);
    if (presentation) {
      const { downloadPresentation } = await import("./presentation/download");
      await downloadPresentation(presentation);
    }
    return;
  }
  if (kind === "media") {
    const media = getMedia(id);
    if (media) {
      downloadMedia(media);
    }
    return;
  }
  if (kind === "spreadsheet") {
    const workbook = getSpreadsheet(id);
    if (workbook) {
      const { downloadSpreadsheet } = await import("./spreadsheet/download");
      await downloadSpreadsheet(workbook);
    }
    return;
  }
  const model = getDocument(id);
  if (model) {
    await downloadDocument(model);
  }
}
