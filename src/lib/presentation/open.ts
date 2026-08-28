// Ties presentation file bytes to the presentation model: classify an
// already-detected package as `.pptx`/`.odp`, parse it, and assemble a
// PresentationModel ready to store and navigate to — the presentation twin
// of openDocument.ts's assembleDocumentModel tail. The app-level funnel
// (openDocument.ts) owns choosing between products; this module owns
// everything from "it's a presentation" down, plus the SPEC-079 migration of
// preparation-milestone records.
import type { PackageInfo } from "../core/packageDetect";
import {
  DEFAULT_PRESENTATION_TITLE,
  type ParsedPresentation,
  type PresentationFormat,
  type PresentationModel,
} from "./model";
import { parsePptx, PPTX_MAIN_CONTENT_TYPE, PptxOpenError } from "./pptx/reader";
import { ODP_MIMETYPE, OdpOpenError, OTP_MIMETYPE, parseOdp } from "./odp/reader";
import { writePptx } from "./pptx/writer";
import { writeOdp } from "./odp/writer";

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
    slideWidthCm: parsed.slideWidthCm,
    slideHeightCm: parsed.slideHeightCm,
    layouts: parsed.layouts,
    opaque: parsed.opaque,
  };
  if (parsed.media.length > 0) {
    model.media = parsed.media;
  }
  if (Object.keys(parsed.properties).length > 0) {
    model.properties = parsed.properties;
  }
  if (parsed.odpContentEnvelope) {
    model.odpContentEnvelope = parsed.odpContentEnvelope;
  }
  return model;
}

/** Re-parses a package from a format's bytes. */
async function parseBytes(bytes: Uint8Array | ArrayBuffer, format: PresentationFormat): Promise<ParsedPresentation> {
  return format === "pptx" ? parsePptx(bytes) : parseOdp(bytes);
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
  const parsed = await parseBytes(bytes, format);
  return assemblePresentation(title, format, parsed);
}

/** Upgrades a preparation-milestone record (its model still holding the
 * outline-shaped slides, no `slideWidthCm`) to the SPEC-079 slide model
 * (AC-54..56): re-emit its retained package via the format's writer (pure
 * part re-emission — nothing about the writers changes), re-parse via the new
 * reader, and keep the same id/title so the store overwrites in place. No
 * data is lost — the retained `opaque` is the source of truth. */
export async function upgradePresentationModel(model: PresentationModel): Promise<PresentationModel> {
  const { bytes } = model.format === "odp" ? await writeOdp(model) : await writePptx(model);
  const parsed = await parseBytes(bytes, model.format);
  const upgraded = assemblePresentation(model.title, model.format, parsed);
  upgraded.id = model.id;
  return upgraded;
}
