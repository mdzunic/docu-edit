// The `.odp` writer (preparation milestone): the retained package re-emitted
// — every part byte-identical, with ODF's container rule honoured: the
// `mimetype` entry comes first and is STOREd (uncompressed) so format
// sniffers can read it from the raw bytes, the same rule odf/writer.ts
// applies to `.odt`.
import { decodedPartBytes } from "../../core/packageParts";
import { zip, type ZipEntryInput } from "../../core/zipWriter";
import { DEFAULT_PRESENTATION_TITLE, type PresentationModel } from "../model";
import { ODP_MIMETYPE } from "./reader";

export { ODP_MIMETYPE };

export function odpFileName(title: string): string {
  const trimmed = title.trim();
  return `${trimmed.length > 0 ? trimmed : DEFAULT_PRESENTATION_TITLE}.odp`;
}

/** Serialises a presentation back to `.odp` bytes: the retained parts,
 * verbatim, `mimetype` first + STOREd. A package missing its `mimetype`
 * part (impossible via parseOdp, which requires it) gets a fresh one so the
 * output is always a sniffable ODF package. */
export async function writeOdp(model: PresentationModel): Promise<{ bytes: Uint8Array }> {
  const parts = decodedPartBytes(model.opaque);
  const mimetype = parts.get("mimetype") ?? new TextEncoder().encode(ODP_MIMETYPE);
  const entries: ZipEntryInput[] = [{ name: "mimetype", data: mimetype, store: true }];
  for (const [name, data] of parts) {
    if (name !== "mimetype") {
      entries.push({ name, data });
    }
  }
  const bytes = await zip(entries);
  return { bytes };
}
