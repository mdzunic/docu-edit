// The `.pptx` writer (preparation milestone): nothing in a presentation is
// editable yet, so a save is the retained package re-emitted — every part
// byte-identical, in package order. The ZIP container itself is rebuilt
// (compression may differ from the source's), which is why round-trip
// tests compare unzipped parts, not container bytes.
import { decodedPartBytes } from "../../core/packageParts";
import { zip } from "../../core/zipWriter";
import { DEFAULT_PRESENTATION_TITLE, type PresentationModel } from "../model";

export { PPTX_MIME_TYPE } from "./reader";

export function pptxFileName(title: string): string {
  const trimmed = title.trim();
  return `${trimmed.length > 0 ? trimmed : DEFAULT_PRESENTATION_TITLE}.pptx`;
}

/** Serialises a presentation back to `.pptx` bytes: the retained parts,
 * verbatim and in order. */
export async function writePptx(model: PresentationModel): Promise<{ bytes: Uint8Array }> {
  const parts = decodedPartBytes(model.opaque);
  const bytes = await zip([...parts].map(([name, data]) => ({ name, data })));
  return { bytes };
}
