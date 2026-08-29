// Plans the package additions a .odt save needs for the model's images
// (SPEC-024): media parts under Pictures/ and their META-INF/manifest.xml
// file-entries. Every save (same-format or not) mints a fresh set of media
// part names from `DocumentModel.media`, mirroring docx/media.ts's own
// "never trust the source package's original naming" approach — see that
// module's header comment for why. Unlike .docx, ODF has no SVG problem:
// `draw:image` stores any media type as-is, so every referenced media part
// is written, none dropped.
import type { DocumentModel } from "../documentModel";
import { ensureMediaEncoded, extensionForContentType } from "../core/media";
import { referencedMediaIds } from "../mediaRefs";
import { base64ToBytes } from "../opaque";
import type { ZipEntryInput } from "../core/zipWriter";

export interface OdtManifestEntry {
  path: string;
  mediaType: string;
}

export interface OdtMediaWritePlan {
  /** The `xlink:href` an Image's `draw:image` should point at — "" for a
   * mediaId with no matching `DocumentModel.media` entry (defensive; the
   * caller omits the frame entirely). */
  hrefFor(mediaId: string): string;
  /** Pictures/imageN.ext parts to add to the package. */
  zipEntries: ZipEntryInput[];
  /** META-INF/manifest.xml file-entries the new parts need. */
  manifestEntries: OdtManifestEntry[];
}

const EMPTY_PLAN: OdtMediaWritePlan = {
  hrefFor: () => "",
  zipEntries: [],
  manifestEntries: [],
};

/** Builds an OdtMediaWritePlan for every image the model's blocks still
 * reference. `existingPartNames` seeds the naming so a same-format save
 * (appending to a preserved package) never collides with a Pictures/ path
 * that package already uses for something else. */
export function planOdtMediaWrite(
  model: DocumentModel,
  existingPartNames: Set<string> = new Set(),
): OdtMediaWritePlan {
  const referenced = referencedMediaIds(model);
  if (referenced.length === 0) {
    return EMPTY_PLAN;
  }
  const byId = new Map((model.media ?? []).map((part) => [part.mediaId, part]));
  const hrefOf = new Map<string, string>();
  const zipEntries: ZipEntryInput[] = [];
  const manifestEntries: OdtManifestEntry[] = [];
  let counter = 1;

  for (const mediaId of referenced) {
    const part = byId.get(mediaId);
    if (!part) {
      continue;
    }
    // SPEC-047: force-encode a still-deferred part (a docx opened with
    // deferred media, saved As .odt) before reading its bytes — a no-op once
    // encoded.
    ensureMediaEncoded(part);
    const ext = extensionForContentType(part.contentType);
    let path = `Pictures/image${counter}.${ext}`;
    while (existingPartNames.has(path)) {
      counter++;
      path = `Pictures/image${counter}.${ext}`;
    }
    counter++;
    hrefOf.set(mediaId, path);
    zipEntries.push({ name: path, data: base64ToBytes(part.data) });
    manifestEntries.push({ path, mediaType: part.contentType });
  }

  return {
    hrefFor: (mediaId: string) => hrefOf.get(mediaId) ?? "",
    zipEntries,
    manifestEntries,
  };
}
