// docu-edit's IndexedDB schema and typed store accessors, built on the
// generic promise layer in core/idb.ts. Kept as its own module so
// documentStore can be unit-tested with this seam mocked, and so nothing
// else touches IndexedDB directly.
//
// Every function rejects when IndexedDB is missing, blocked, or over quota —
// callers (documentStore) treat any rejection as "storage unavailable" and
// keep working in-memory.
import { openIdbDatabase } from "./core/idb";
import type { PersistedRecord } from "./documentStore";
import type { VersionStoreRecord } from "./versionHistory";
import type { StoredTemplateRecord } from "./templateStore";

const DB_NAME = "docu-edit";
const DB_VERSION = 3;
const STORE_NAME = "documents";
// SPEC-040: version history's own store, kept separate from `documents` so
// rendering the library never reads 50 MB of version payloads per document
// into memory at startup (see versionHistory.ts's header comment).
const VERSIONS_STORE_NAME = "versions";
// SPEC-046: user templates get their own store too, for the same reason —
// the gallery reads every template's metadata without touching the
// `documents` store's much larger population.
const TEMPLATES_STORE_NAME = "templates";

const database = openIdbDatabase(DB_NAME, DB_VERSION, [
  { name: STORE_NAME, keyPath: "id" },
  { name: VERSIONS_STORE_NAME, keyPath: "id" },
  { name: TEMPLATES_STORE_NAME, keyPath: "id" },
]);

/** Reads every persisted record — committed documents and autosave drafts
 * alike; callers split them by key prefix. */
export async function loadAllRecords(): Promise<PersistedRecord[]> {
  return database.getAll<PersistedRecord>(STORE_NAME);
}

/** Reads one record by key, resolving undefined when absent. */
export async function getRecord(
  id: string,
): Promise<PersistedRecord | undefined> {
  return database.get<PersistedRecord>(STORE_NAME, id);
}

/** Writes (inserts or replaces) one record. */
export async function putRecord(record: PersistedRecord): Promise<void> {
  await database.put(STORE_NAME, record);
}

/** Removes one document record by id. Resolves even if the id is absent. */
export async function deleteRecord(id: string): Promise<void> {
  await database.delete(STORE_NAME, id);
}

// --- version history (SPEC-040) --------------------------------------------
// A separate object store, and separate helpers, so a version-history write
// or read is always its own IndexedDB transaction on its own store — a
// quota failure here can never abort a `documents` store transaction (AC-40).

/** Reads one version-history record (index, payload or blob) by key. */
export async function getVersionsRecord(id: string): Promise<VersionStoreRecord | undefined> {
  return database.get<VersionStoreRecord>(VERSIONS_STORE_NAME, id);
}

/** Writes (inserts or replaces) one version-history record. */
export async function putVersionsRecord(record: VersionStoreRecord): Promise<void> {
  await database.put(VERSIONS_STORE_NAME, record);
}

/** Removes one version-history record by id. Resolves even if absent. */
export async function deleteVersionsRecord(id: string): Promise<void> {
  await database.delete(VERSIONS_STORE_NAME, id);
}

// --- user templates (SPEC-046) ----------------------------------------------

/** Reads every persisted user template. */
export async function loadAllTemplateRecords(): Promise<StoredTemplateRecord[]> {
  return database.getAll<StoredTemplateRecord>(TEMPLATES_STORE_NAME);
}

/** Writes (inserts or replaces) one user template record. */
export async function putTemplateRecord(record: StoredTemplateRecord): Promise<void> {
  await database.put(TEMPLATES_STORE_NAME, record);
}

/** Removes one user template record by id. Resolves even if absent. */
export async function deleteTemplateRecord(id: string): Promise<void> {
  await database.delete(TEMPLATES_STORE_NAME, id);
}
