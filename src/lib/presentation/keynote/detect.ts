// The Keynote importer's package signatures and its three user-facing
// rejection strings (SPEC-087), apart from the decode stack in reader.ts:
// the app's open funnel (openDocument.ts) checks every ZIP it is handed for
// an iWork archive, and rejects some with these strings, before it knows
// whether a Keynote import is needed at all — so it must be able to ask
// without loading the importer (see ../detect.ts). reader.ts re-exports all
// of this, so its importers are unchanged.

// The three user-facing rejection strings (AC-8..AC-11) — every failure
// path in the importer throws one of them as a KeynoteOpenError (AC-12).
export const KEYNOTE_LEGACY_MESSAGE =
  "This is a Keynote '09 file — open it in Keynote and save it again, or export it as .pptx, then try that file";
export const NOT_KEYNOTE_MESSAGE = "This iWork file isn't a Keynote presentation we can open";
export const KEYNOTE_UNREADABLE_MESSAGE =
  "This Keynote file can't be opened — export it as .pptx from Keynote and try that file";

/** Whether an already-unzipped package is an iWork Archive: it holds
 * `Index/Document.iwa`, or an `Index.zip` (the older nesting — its entries
 * are checked when parsing) (AC-5/AC-6). Content-based; the extension never
 * decides. */
export function isIwaPackage(parts: Map<string, Uint8Array>): boolean {
  return parts.has("Index/Document.iwa") || parts.has("Index.zip");
}

/** Whether a ZIP is an iWork '05–'09 file (`index.apxl`/`index.apxl.gz`),
 * which uses a different, unsupported format (AC-8). */
export function isLegacyKeynote(parts: Map<string, Uint8Array>): boolean {
  return parts.has("index.apxl") || parts.has("index.apxl.gz");
}
