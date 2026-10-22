// The presentation product's package signatures: the declared types that
// make an already-detected package a `.pptx` or an `.odp`, and nothing else.
// They live apart from the readers because the app's open funnel
// (openDocument.ts) asks this question of every file it is handed, text
// documents included, and answering it must not load a presentation codec:
// a static import of the readers put the whole pptx/odp read-write stack on
// the library page and the editor's module graphs, inside the window
// SPEC-047's firstPageMs times. The readers and open.ts re-export these, so
// their importers are unchanged.
import type { PackageInfo } from "../core/packageDetect";
import type { PresentationFormat } from "./model";

/** The main part's declared content type that makes an OPC package a
 * presentation (ECMA-376 PresentationML). */
export const PPTX_MAIN_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml";

/** The package mimetype identifying an OpenDocument presentation. */
export const ODP_MIMETYPE = "application/vnd.oasis.opendocument.presentation";
/** An `.otp` template's mimetype — same package shape, opened the same way. */
export const OTP_MIMETYPE = "application/vnd.oasis.opendocument.presentation-template";

/** Which presentation format a detected package is, or null when it is not
 * a presentation at all (the caller then tries the text formats). */
export function presentationFormatOf(info: PackageInfo): PresentationFormat | null {
  if (info.container === "odf" && (info.mimetype === ODP_MIMETYPE || info.mimetype === OTP_MIMETYPE)) {
    return "odp";
  }
  if (info.container === "opc" && info.mainContentType === PPTX_MAIN_CONTENT_TYPE) {
    return "pptx";
  }
  return null;
}
