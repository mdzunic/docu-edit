// Ties presentation file bytes to the presentation model: classify an
// already-detected package as `.pptx`/`.odp`, parse it, and assemble a
// PresentationModel ready to store and navigate to — the presentation twin
// of openDocument.ts's assembleDocumentModel tail. The app-level funnel
// (openDocument.ts) owns choosing between products; this module owns
// everything from "it's a presentation" down.
import type { PackageInfo } from "../core/packageDetect";
import {
  DEFAULT_PRESENTATION_TITLE,
  type ParsedPresentation,
  type PresentationFormat,
  type PresentationModel,
} from "./model";
import { parsePptx, PPTX_MAIN_CONTENT_TYPE, PptxOpenError } from "./pptx/reader";
import { ODP_MIMETYPE, OdpOpenError, OTP_MIMETYPE, parseOdp } from "./odp/reader";

export { PptxOpenError, OdpOpenError };

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

/** Assembles a parsed package into a fresh PresentationModel: the shared
 * tail of every presentation open path, minting a new id the way
 * assembleDocumentModel does. Empty optional fields are omitted, matching
 * the text model's convention. */
export function assemblePresentation(
  title: string,
  format: PresentationFormat,
  parsed: ParsedPresentation,
): PresentationModel {
  const model: PresentationModel = {
    kind: "presentation",
    id: crypto.randomUUID(),
    title,
    format,
    slides: parsed.slides,
    opaque: parsed.opaque,
  };
  if (Object.keys(parsed.properties).length > 0) {
    model.properties = parsed.properties;
  }
  return model;
}

/** Reads and parses presentation bytes directly into a new
 * PresentationModel, with no `File`/FileReader involved — the Node-friendly
 * counterpart of the app's open path (the same split as readDocxBytes /
 * readOdtBytes, SPEC-045 AC-17). */
export async function readPresentationBytes(
  bytes: Uint8Array | ArrayBuffer,
  format: PresentationFormat,
  title: string = DEFAULT_PRESENTATION_TITLE,
): Promise<PresentationModel> {
  const parsed = format === "pptx" ? await parsePptx(bytes) : await parseOdp(bytes);
  return assemblePresentation(title, format, parsed);
}
