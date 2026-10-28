// SPEC-040's version-chain arithmetic, made product-neutral by SPEC-103 so the
// word processor's history (lib/versionHistory.ts) and the deck's
// (lib/presentationVersions.ts) chain, budget and prune by the same rules: the
// record shapes both write into the `versions` object store, their key
// layout, the four baseline triggers, chain gathering, prune-candidate
// selection, the byte measure and the content digest. Nothing here reads or
// writes a store or knows a model's shape — each product keeps its own payload
// encoding and store IO (core imports only core siblings).

// --- constants ---------------------------------------------------------------

/** The per-document history budget — the same ceiling SPEC-009 puts on a
 * single file the app will open (`MAX_FILE_SIZE_BYTES` in openDocument.ts,
 * which core cannot import, so it is restated as a literal here;
 * versionHistory.spec.ts pins the two equal). */
export const HISTORY_BUDGET_BYTES = 50 * 1024 * 1024;

/** Baseline trigger 2: the reconstruction bound (at most 51 records read and
 * applied for any retained version). */
export const CHAIN_DELTA_CAP = 50;

/** Baseline trigger 4: a delta at or above this fraction of its chain's
 * baseline size stores as a fresh baseline instead. */
export const WHOLESALE_REWRITE_RATIO = 0.5;

/** Pruning never removes the newest this-many versions. */
export const RETAINED_RECENT_COUNT = 10;

/** The exact notice text SPEC-040 AC-38 requires. */
export const HISTORY_PAUSED_MESSAGE =
  "Version history paused — 50 MB used. Delete a version, or remove a name, to resume";

// --- record shapes -----------------------------------------------------------

export type VersionChainKind = "baseline" | "delta";

export type SnapshotTrigger = "cadence" | "commit" | "pre-restore" | "post-restore";

export interface VersionIndexEntry {
  id: string;
  timestamp: number;
  author: string;
  words: number;
  name?: string;
  kind: VersionChainKind;
  parentId: string | null;
  /** Encoded (JSON, UTF-8) byte size of this version's payload record —
   * measured, not estimated (SPEC-040 AC-44), and what the budget/pruning
   * arithmetic and the four baseline triggers are computed from. */
  size: number;
  /** A deck version's slide count (SPEC-103). Additive, beside `words`
   * rather than in place of it: renaming that field would orphan every
   * stored `versions:<docId>` record. Absent on text versions. */
  slides?: number;
  /** A workbook version's sheet count and stored-cell count (SPEC-138),
   * additive beside `slides` for the same reason. Absent on every other
   * product's versions. */
  sheets?: number;
  cells?: number;
  /** What took the snapshot (SPEC-103: the deck's panel names its restore
   * pair). Absent on text versions and on any recorded before it. */
  trigger?: SnapshotTrigger;
}

export interface VersionsIndexRecord {
  id: string;
  store: "index";
  docId: string;
  /** Chronological, oldest first — the panes' "newest first" is a display
   * concern, applied where the list renders. */
  versions: VersionIndexEntry[];
  /** True once a snapshot attempt found nothing prunable and the budget was
   * still exceeded (SPEC-040 AC-37/38); cleared the next time usage drops
   * back under budget, whether from a new prune or a manual delete/unname. */
  paused: boolean;
  /** Set beside `paused` when the latch is not the budget's exhaustion but a
   * deck whose blob set plus a single baseline can never fit it (SPEC-103's
   * over-budget degrade) — or a workbook past SPEC-138's cell cap or budget.
   * Absent on every text index. */
  pausedReason?: "too-large";
}

export interface VersionBlobRecord {
  id: string;
  store: "blob";
  docId: string;
  hash: string;
  contentType?: string;
  /** Base64, the same convention as MediaPart/OpaquePart. */
  data: string;
  size: number;
  refCount: number;
}

// --- key layout (one `versions` object store, split by key prefix) ----------

/** `versions:<docId>` — the VersionsIndexRecord, one per document. */
export function versionsIndexKey(docId: string): string {
  return `versions:${docId}`;
}

/** `version:<docId>:<versionId>` — one payload record per version. */
export function versionPayloadKey(docId: string, versionId: string): string {
  return `version:${docId}:${versionId}`;
}

/** `blob:<docId>:<sha256>` — shared, reference-counted part bytes. */
export function versionBlobKey(docId: string, hash: string): string {
  return `blob:${docId}:${hash}`;
}

/** The index of a document with no history yet. */
export function emptyVersionsIndex(docId: string): VersionsIndexRecord {
  return { id: versionsIndexKey(docId), store: "index", docId, versions: [], paused: false };
}

// --- measurement and digest -------------------------------------------------

const encoder = new TextEncoder();

/** Encoded (JSON, UTF-8) byte size of a value — what every size the budget
 * arithmetic compares is measured with. */
export function measureBytes(value: unknown): number {
  return encoder.encode(JSON.stringify(value)).length;
}

/** Lower-case hex SHA-256 of a string's UTF-8 bytes (crypto.subtle — no
 * dependency). Blobs are content-addressed by it over their base64 text. */
export async function sha256Hex(text: string): Promise<string> {
  const digest = await crypto.subtle.digest("SHA-256", encoder.encode(text));
  return Array.from(new Uint8Array(digest))
    .map((byte) => byte.toString(16).padStart(2, "0"))
    .join("");
}

// --- baseline/delta decision (SPEC-040 AC-23, checked in this order) --------

export interface ChainInfo {
  baselineId: string;
  baselineBytes: number;
  deltaCount: number;
  accumulatedDeltaBytes: number;
}

/** Walks the index (no payloads) from `tailId` back to its chain's baseline,
 * summing the deltas already in the chain. Returns null when the chain is
 * malformed/unreadable — treated as baseline trigger 1. */
export function gatherChainInfo(entries: VersionIndexEntry[], tailId: string): ChainInfo | null {
  const byId = new Map(entries.map((e) => [e.id, e]));
  let cur = byId.get(tailId);
  if (!cur) {
    return null;
  }
  let deltaCount = 0;
  let accumulatedDeltaBytes = 0;
  while (cur.kind === "delta") {
    deltaCount += 1;
    accumulatedDeltaBytes += cur.size;
    if (!cur.parentId) {
      return null;
    }
    const parent = byId.get(cur.parentId);
    if (!parent) {
      return null;
    }
    cur = parent;
  }
  return { baselineId: cur.id, baselineBytes: cur.size, deltaCount, accumulatedDeltaBytes };
}

/** The four baseline triggers, checked in order (SPEC-040 AC-23); returns
 * false only when none holds, i.e. the snapshot should be a delta (AC-24). */
export function decideBaseline(chainInfo: ChainInfo | null, deltaBytes: number): boolean {
  if (!chainInfo) {
    return true; // 1: no chain to extend
  }
  if (chainInfo.deltaCount >= CHAIN_DELTA_CAP) {
    return true; // 2: chain cap
  }
  if (chainInfo.accumulatedDeltaBytes + deltaBytes >= chainInfo.baselineBytes) {
    return true; // 3: chain outweighs a copy
  }
  if (deltaBytes >= chainInfo.baselineBytes * WHOLESALE_REWRITE_RATIO) {
    return true; // 4: wholesale rewrite
  }
  return false;
}

// --- prune-candidate selection (SPEC-040 AC-32..AC-36) ----------------------

/** A named version, or one of the newest RETAINED_RECENT_COUNT, is never
 * pruned. */
export function isRetained(entries: VersionIndexEntry[], entry: VersionIndexEntry): boolean {
  if (entry.name) {
    return true;
  }
  const newest = entries.slice(-RETAINED_RECENT_COUNT);
  return newest.includes(entry);
}

/** The oldest version pruning may remove, or null when every one is retained. */
export function oldestPrunable(entries: VersionIndexEntry[]): VersionIndexEntry | null {
  for (const entry of entries) {
    if (!isRetained(entries, entry)) {
      return entry;
    }
  }
  return null;
}

/** The oldest RETAINED version whose reconstruction chain passes through
 * `fromId`, found by walking forward one link at a time — cheap because a
 * chain never fans out (each id is at most one other version's parent). */
export function oldestRetainedDescendant(entries: VersionIndexEntry[], fromId: string): VersionIndexEntry | null {
  const childOf = new Map(entries.map((e) => [e.parentId, e]));
  let cur = childOf.get(fromId);
  while (cur) {
    if (isRetained(entries, cur)) {
      return cur;
    }
    cur = childOf.get(cur.id);
  }
  return null;
}
