// Byte-based reading (SPEC-045 AC-17): a generated Buffer/Uint8Array can be
// passed straight to readDocx/readOdt with no disk round trip and no
// File/FileReader involved. These wrap openDocument.ts's readDocxBytes/
// readOdtBytes (the same assembly logic openDocumentFile uses) with the
// Node DOMParser polyfill the underlying readers need outside a browser —
// see nodeDom.ts.
import { readDocxBytes, readOdtBytes } from "../openDocument";
import type { DocumentModel } from "../documentModel";
import { ensureNodeXmlGlobals } from "./nodeDom";

/** Reads `.docx` bytes into a `DocumentModel`, in the browser or in Node
 * (AC-17) — no file, no disk. `title` defaults the same way a blank
 * document does; pass one when the bytes came from a named source. */
export async function readDocx(bytes: Uint8Array | ArrayBuffer, title?: string): Promise<DocumentModel> {
  await ensureNodeXmlGlobals();
  return readDocxBytes(bytes, title);
}

/** The `.odt` counterpart of `readDocx` (AC-17). */
export async function readOdt(bytes: Uint8Array | ArrayBuffer, title?: string): Promise<DocumentModel> {
  await ensureNodeXmlGlobals();
  return readOdtBytes(bytes, title);
}
