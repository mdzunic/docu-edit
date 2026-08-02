// Session-only in-memory document store. Documents created or opened while
// the app is running live here; nothing survives a reload — that persistence
// layer is SPEC-003 (IndexedDB) and out of scope for this spec.
import { createDocument, type DocumentModel } from "./documentModel";

const documents = new Map<string, DocumentModel>();

/** Returns the stored document for an id, or undefined if none exists. */
export function getDocument(id: string): DocumentModel | undefined {
  return documents.get(id);
}

/** Returns the stored document for an id, creating and storing a fresh one
 * if none exists yet. */
export function getOrCreateDocument(id: string): DocumentModel {
  const existing = documents.get(id);
  if (existing) {
    return existing;
  }
  const doc = createDocument(id);
  documents.set(id, doc);
  return doc;
}
