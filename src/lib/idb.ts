// docu-edit's IndexedDB schema and typed store accessors, built on the
// generic promise layer in core/idb.ts. Kept as its own module so
// documentStore can be unit-tested with this seam mocked, and so nothing
// else touches IndexedDB directly.
//
// Every function rejects when IndexedDB is missing, blocked, or over quota —
// callers (documentStore) treat any rejection as "storage unavailable" and
// keep working in-memory.
import { openIdbDatabase, type WriteBatchOp } from "./core/idb";
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

export type { WriteBatchOp };

/** The key prefixes of the records the start-up read skips (SPEC-137 "The
 * storage layout"): a workbook's row chunks (`sheetrows:`) and its source
 * record (`sheetsource:`) are read only when the workbook is loaded
 * (documentStore.loadSpreadsheet), so the landing page reads no cell and no
 * retained package. SPEC-138's draft chunk and draft source prefixes join
 * this list. The same strings as documentStore's SHEET_ROWS_KEY_PREFIX and
 * SHEET_SOURCE_KEY_PREFIX — spelled out here because documentStore imports
 * this module (documentStore.sheets.spec.ts pins the two equal). */
export const LAZY_RECORD_KEY_PREFIXES: readonly string[] = ["sheetrows:", "sheetsource:"];

/** Above every key that starts with a prefix: no key holds U+FFFF (document
 * ids are UUIDs, sheet ids `[A-Za-z0-9_-]`). */
const PREFIX_END = "\uffff";

/** The key ranges outside every prefix of `excluded`, in key order: below
 * the first, between each prefix's end and the next prefix, above the last
 * prefix's end — each bound open. */
export function rangesExcluding(excluded: readonly string[]): { lower?: string; upper?: string }[] {
  const sorted = [...new Set(excluded)].toSorted();
  const ranges: { lower?: string; upper?: string }[] = [];
  let lower: string | undefined;
  for (const prefix of sorted) {
    ranges.push({ lower, upper: prefix });
    lower = prefix + PREFIX_END;
  }
  ranges.push({ lower });
  return ranges;
}

/** Reads every persisted record the library needs at start-up — committed
 * shells, autosave drafts, media — through the key ranges that exclude
 * LAZY_RECORD_KEY_PREFIXES (SPEC-137: three ranges today); callers split
 * them by key prefix. */
export async function loadAllRecords(): Promise<PersistedRecord[]> {
  const parts = await Promise.all(
    rangesExcluding(LAZY_RECORD_KEY_PREFIXES).map(({ lower, upper }) =>
      database.getAllInRange<PersistedRecord>(STORE_NAME, lower, upper, { lowerOpen: true, upperOpen: true }),
    ),
  );
  return parts.flat();
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

/** Reads every record whose key lies in [lower, upper] (SPEC-137). */
export async function getRecordsInRange(lower: string, upper: string): Promise<PersistedRecord[]> {
  return database.getAllInRange<PersistedRecord>(STORE_NAME, lower, upper);
}

/** Removes every record whose key lies in [lower, upper] (SPEC-137). */
export async function deleteRecordRange(lower: string, upper: string): Promise<void> {
  await database.deleteRange(STORE_NAME, lower, upper);
}

/** Applies puts and deletes to the documents store in ONE readwrite
 * transaction, one operation per success callback (SPEC-137: a workbook's
 * changed chunks, deletions, source record and shell commit together). */
export async function writeBatch(ops: readonly WriteBatchOp[]): Promise<void> {
  await database.writeBatch(STORE_NAME, ops);
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
