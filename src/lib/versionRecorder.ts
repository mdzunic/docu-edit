// SPEC-040's recording loop for every product that keeps its history in the
// shared `versions` store beside the word processor's (SPEC-138 moved it here
// from presentationVersions.ts, where SPEC-103 first wrote it for decks): load
// the index → chain info → build the delta → decideBaseline → ensure blob refs
// → write the payload and the index → enforce the 50 MB budget and prune, plus
// the chain walk that reconstructs a version, the refcounted blobs and the
// panel's name / unname / delete. A product supplies only its payload encoding
// as a VersionCodec — the deck's (presentationVersions.ts) and the workbook's
// (spreadsheetVersions.ts). The text side's versionHistory.ts keeps its own.
//
// Records share the store's three key prefixes — `versions:<docId>` (the
// index), `version:<docId>:<versionId>` (a payload), `blob:<docId>:<sha256>`
// (part bytes, reference-counted) — so there is no schema change; payloads are
// told apart by their kinds (codec.owns).
//
// History is best-effort and never the document's problem: nothing here
// touches the `documents` store or the storage banner, every failure resolves
// to "not written", and the callers never await it on an edit's path.
import { yieldToMainThread } from "./core/idleScheduler";
import type { MediaPart } from "./core/media";
import type { OpaquePart } from "./core/packageParts";
import {
  decideBaseline,
  emptyVersionsIndex,
  gatherChainInfo,
  isRetained,
  measureBytes,
  oldestPrunable,
  oldestRetainedDescendant,
  sha256Hex,
  versionBlobKey,
  versionPayloadKey,
  versionsIndexKey,
  HISTORY_BUDGET_BYTES,
  type SnapshotTrigger,
  type VersionBlobRecord,
  type VersionIndexEntry,
  type VersionsIndexRecord,
} from "./core/versionChain";
import { deleteVersionsRecord, getVersionsRecord, putVersionsRecord as putStoreRecord } from "./idb";

// --- refs and blobs (shared by every codec) ------------------------------------

/** A retained package part by content: its bytes live in `blob:<docId>:<hash>`. */
export interface OpaqueRef {
  name: string;
  hash: string;
}

/** An image/clip part by content (bytes in its blob). */
export interface MediaRef {
  mediaId: string;
  contentType: string;
  hash: string;
}

/** One distinct blob a model references: its stored text and its measured
 * size (what the budget counts). */
export interface BlobSource {
  data: string;
  contentType?: string;
  size: number;
}

export type BlobSet = Map<string, BlobSource>;

/** What every payload looks like from the loop's side: a kind, and the
 * parent a delta applies to. */
export interface VersionPayload {
  kind: string;
  parentId?: string;
}

export interface VersionPayloadRecord<P extends VersionPayload = VersionPayload> {
  id: string;
  store: "payload";
  docId: string;
  versionId: string;
  payload: P;
}

/** Everything a recorder reads from and writes to the `versions` store. */
export type RecorderStoreRecord = VersionsIndexRecord | VersionPayloadRecord | VersionBlobRecord;

/** idb's put, over the recorder's generic records: the store's own union
 * names each product's payload shape, which a codec's `P` is one of. */
function putVersionsRecord(record: RecorderStoreRecord): Promise<void> {
  return putStoreRecord(record as Parameters<typeof putStoreRecord>[0]);
}

/** measureBytes of a base64 string, without encoding megabytes to count them:
 * base64 is ASCII and needs no JSON escaping, so it is its length plus the
 * two quotes. */
export function base64Bytes(data: string): number {
  return data.length + 2;
}

// Hashing a package is the expensive part of a snapshot, so each part's digest
// is kept beside the exact string it was taken over: an unchanged part is
// hashed once per session, and a part whose `data` was replaced is re-hashed.
const hashCache = new WeakMap<OpaquePart | MediaPart, { data: string; hash: string }>();

export async function hashPart(part: OpaquePart | MediaPart): Promise<string> {
  const cached = hashCache.get(part);
  if (cached && cached.data === part.data) {
    return cached.hash;
  }
  const hash = await sha256Hex(part.data);
  hashCache.set(part, { data: part.data, hash });
  return hash;
}

export async function opaqueRefsOf(parts: OpaquePart[] | undefined): Promise<OpaqueRef[]> {
  const refs: OpaqueRef[] = [];
  for (const part of parts ?? []) {
    refs.push({ name: part.name, hash: await hashPart(part) });
  }
  return refs;
}

export async function mediaRefsOf(parts: MediaPart[] | undefined): Promise<MediaRef[]> {
  const refs: MediaRef[] = [];
  for (const part of parts ?? []) {
    refs.push({ mediaId: part.mediaId, contentType: part.contentType, hash: await hashPart(part) });
  }
  return refs;
}

/** Each distinct blob a model's package and media reference, with its bytes —
 * the two deduplicated by hash (a deck's media duplicates image parts already
 * in its package). */
export async function partBlobsOf(opaque: OpaquePart[] | undefined, media: MediaPart[] | undefined): Promise<BlobSet> {
  const blobs: BlobSet = new Map();
  for (const part of opaque ?? []) {
    blobs.set(await hashPart(part), { data: part.data, size: base64Bytes(part.data) });
  }
  for (const part of media ?? []) {
    blobs.set(await hashPart(part), { data: part.data, contentType: part.contentType, size: base64Bytes(part.data) });
  }
  return blobs;
}

export function sameRefs(a: readonly (OpaqueRef | MediaRef)[], b: readonly (OpaqueRef | MediaRef)[]): boolean {
  return a.length === b.length && a.every((ref, i) => JSON.stringify(ref) === JSON.stringify(b[i]));
}

/** Reads each ref's bytes back, reusing a part of `live` whose content is the
 * same (and `live`'s whole array when every part is): a restore whose package
 * did not change keeps the model's own `opaque` reference, so the store's
 * latch skips rewriting it. */
export async function hydrateParts<P extends OpaquePart | MediaPart, R extends OpaqueRef | MediaRef>(
  refs: R[],
  live: P[] | undefined,
  build: (ref: R, data: string) => P,
  sameSlot: (part: P, ref: R) => boolean,
  read: (hash: string) => Promise<string>,
): Promise<P[]> {
  const parts: P[] = [];
  let allLive = live !== undefined && live.length === refs.length;
  for (const [i, ref] of refs.entries()) {
    const candidate = live?.[i];
    if (candidate && sameSlot(candidate, ref) && (await hashPart(candidate)) === ref.hash) {
      parts.push(candidate);
      continue;
    }
    allLive = false;
    parts.push(build(ref, await read(ref.hash)));
  }
  return allLive && live ? live : parts;
}

export function hydrateOpaque(refs: OpaqueRef[], live: OpaquePart[] | undefined, read: (hash: string) => Promise<string>): Promise<OpaquePart[]> {
  return hydrateParts(refs, live, (ref, data): OpaquePart => ({ name: ref.name, data }), (part, ref) => part.name === ref.name, read);
}

export function hydrateMedia(refs: MediaRef[], live: MediaPart[] | undefined, read: (hash: string) => Promise<string>): Promise<MediaPart[]> {
  return hydrateParts(
    refs,
    live,
    (ref, data): MediaPart => ({ mediaId: ref.mediaId, contentType: ref.contentType, data }),
    (part, ref) => part.mediaId === ref.mediaId && part.contentType === ref.contentType,
    read,
  );
}

// --- the codec -------------------------------------------------------------------

/** A product's payload encoding. `C` is the model's content captured at one
 * instant, `T` that capture as the chain compares it (the next delta's
 * predecessor, kept in memory as `cachedPrev`), `P` the stored payload and
 * `S` a version without its bytes (a chain applied, refs only). */
export interface VersionCodec<M, C, T extends { id: string }, P extends VersionPayload, S> {
  /** Whether a stored payload is this codec's (the store holds every
   * product's history). */
  owns(payload: VersionPayload): payload is P;
  isBaseline(payload: P): boolean;
  /** The model's content now — synchronous, taken before the loop awaits
   * anything, so an edit landing while a snapshot writes stays out of it. */
  capture(model: M): C;
  /** A refusal decided before any encode (SPEC-138's cell cap). */
  tooLargeToEncode?(capture: C): boolean;
  /** The capture as the chain compares it, its blobs hashed. */
  tail(id: string, capture: C): Promise<T>;
  /** Every distinct blob the capture references, with its bytes. */
  blobs(capture: C): Promise<BlobSet>;
  baseline(next: T): P;
  delta(prev: T, next: T): P;
  /** The index entry's own counts (a deck's slides; a workbook's sheets and
   * cells). */
  stats(next: T): Partial<VersionIndexEntry>;
  /** The blob hashes a payload itself embeds, once each — what its
   * refcounts count; an inherited ref is its ancestor's reference. */
  embeddedHashes(payload: P): string[];
  /** A baseline's skeleton, and a delta applied to its parent's. */
  skeleton(baseline: P): S;
  apply(parent: S, delta: P): S;
  /** A skeleton re-encoded to stand alone (pruning's re-baselining). */
  baselineOf(skeleton: S): P;
  /** A skeleton as the next delta's predecessor (a session that has not
   * recorded a version yet). */
  tailOf(id: string, skeleton: S): T;
  /** The whole model, its blobs read back; `live` lends any part whose
   * content is unchanged. */
  hydrate(docId: string, skeleton: S, read: (hash: string) => Promise<string>, live?: M): Promise<M>;
  /** The blobs the open model references: kept at refCount 0. */
  liveHashes(model: M): Promise<Set<string>>;
}

export interface RecordVersionOptions<M, T> {
  docId: string;
  model: M;
  author: string;
  trigger: SnapshotTrigger;
  /** The last version recorded this session; null on a fresh session. */
  cachedPrev: T | null;
}

export interface RecordVersionResult<T> {
  written: boolean;
  index: VersionsIndexRecord;
  newEntry: T | null;
}

// --- store IO --------------------------------------------------------------------

/** A document's index; an empty one when absent or unreadable. */
export async function loadVersionsIndex(docId: string): Promise<VersionsIndexRecord> {
  try {
    const record = await getVersionsRecord(versionsIndexKey(docId));
    return record && record.store === "index" ? record : emptyVersionsIndex(docId);
  } catch {
    return emptyVersionsIndex(docId);
  }
}

async function readBlob(docId: string, hash: string): Promise<string> {
  const record = await getVersionsRecord(versionBlobKey(docId, hash));
  if (!record || record.store !== "blob") {
    throw new Error(`Missing history blob ${hash} for ${docId}`);
  }
  return record.data;
}

async function ensureBlobRefs(docId: string, hashes: string[], blobs: BlobSet): Promise<void> {
  for (const hash of hashes) {
    const key = versionBlobKey(docId, hash);
    const existing = await getVersionsRecord(key);
    if (existing && existing.store === "blob") {
      await putVersionsRecord({ ...existing, refCount: existing.refCount + 1 });
      continue;
    }
    const source = blobs.get(hash);
    if (!source) {
      continue;
    }
    const blob: VersionBlobRecord = { id: key, store: "blob", docId, hash, data: source.data, size: source.size, refCount: 1 };
    if (source.contentType !== undefined) {
      blob.contentType = source.contentType;
    }
    await putVersionsRecord(blob);
  }
}

/** Drops one reference per hash. A blob the open document still references
 * stays at refCount 0 (SPEC-040 AC-29's rule) — the next snapshot
 * re-references it instead of rewriting megabytes — and goes once nothing
 * does. */
async function releaseBlobRefs(docId: string, hashes: string[], keep: Set<string>): Promise<void> {
  for (const hash of hashes) {
    const key = versionBlobKey(docId, hash);
    const existing = await getVersionsRecord(key);
    if (!existing || existing.store !== "blob") {
      continue;
    }
    const refCount = existing.refCount - 1;
    if (refCount > 0 || keep.has(hash)) {
      await putVersionsRecord({ ...existing, refCount: Math.max(0, refCount) });
    } else {
      await deleteVersionsRecord(key);
    }
  }
}

/** The unretained versions whose chain runs through `fromId` before it
 * reaches a retained one (or the chain's end) — oldest first. Once `fromId`
 * goes they can never be reconstructed again. */
function unretainedChainAfter(entries: VersionIndexEntry[], fromId: string): string[] {
  const childOf = new Map(entries.map((e) => [e.parentId, e]));
  const ids: string[] = [];
  for (let cur = childOf.get(fromId); cur && !isRetained(entries, cur); cur = childOf.get(cur.id)) {
    ids.push(cur.id);
  }
  return ids;
}

function withoutPausedReason(index: VersionsIndexRecord): VersionsIndexRecord {
  const { pausedReason: _cleared, ...rest } = index;
  return rest;
}

// --- the recorder ----------------------------------------------------------------

/** One product's history over its codec: the recording loop, reconstruction,
 * the budget and the panel's actions. */
export function createVersionRecorder<M, C, T extends { id: string }, P extends VersionPayload, S>(codec: VersionCodec<M, C, T, P, S>) {
  async function loadPayload(docId: string, versionId: string): Promise<P | null> {
    const stored = await getVersionsRecord(versionPayloadKey(docId, versionId));
    const payload = stored && stored.store === "payload" ? (stored.payload as VersionPayload) : null;
    return payload && codec.owns(payload) ? payload : null;
  }

  /** Walks a version's chain back to its baseline (at most CHAIN_DELTA_CAP +
   * 1 records) and applies it forward — refs only, no blob read. Null when
   * the chain is missing, malformed or unreadable. */
  async function reconstructSkeleton(docId: string, versionId: string): Promise<S | null> {
    const chain: P[] = [];
    const seen = new Set<string>();
    let currentId: string | null = versionId;
    try {
      while (currentId) {
        if (seen.has(currentId)) {
          return null;
        }
        seen.add(currentId);
        const payload = await loadPayload(docId, currentId);
        if (!payload) {
          return null;
        }
        chain.unshift(payload);
        currentId = codec.isBaseline(payload) ? null : (payload.parentId ?? null);
      }
    } catch {
      return null;
    }
    const [base, ...deltas] = chain;
    if (!base || !codec.isBaseline(base)) {
      return null;
    }
    let skeleton = codec.skeleton(base);
    for (const delta of deltas) {
      skeleton = codec.apply(skeleton, delta);
    }
    return skeleton;
  }

  /** Reconstructs a version's whole model — its blobs read back — or null
   * when its chain or a blob is unreadable. */
  async function reconstruct(docId: string, versionId: string, live?: M): Promise<M | null> {
    const skeleton = await reconstructSkeleton(docId, versionId);
    if (!skeleton) {
      return null;
    }
    const bytes = new Map<string, Promise<string>>();
    const read = (hash: string): Promise<string> => {
      let pending = bytes.get(hash);
      if (!pending) {
        pending = readBlob(docId, hash);
        bytes.set(hash, pending);
      }
      return pending;
    };
    try {
      return await codec.hydrate(docId, skeleton, read, live);
    } catch {
      return null;
    }
  }

  /** Bytes the history holds: every version's payload plus every blob a
   * version still references — the blobs it shares with the open document
   * too: a package is the bulk of what history stores (SPEC-103). */
  async function usedBytes(docId: string, index: VersionsIndexRecord): Promise<number> {
    let total = 0;
    const hashes = new Set<string>();
    for (const entry of index.versions) {
      total += entry.size;
      const payload = await loadPayload(docId, entry.id);
      for (const hash of payload ? codec.embeddedHashes(payload) : []) {
        hashes.add(hash);
      }
    }
    for (const hash of hashes) {
      const blob = await getVersionsRecord(versionBlobKey(docId, hash));
      if (blob && blob.store === "blob" && blob.refCount > 0) {
        total += blob.size;
      }
    }
    return total;
  }

  /** Re-encodes `entryId` as a standalone baseline so nothing after it
   * depends on anything older (pruning's chain-integrity step). */
  async function reencodeAsBaseline(docId: string, entryId: string, index: VersionsIndexRecord, keep: Set<string>): Promise<VersionsIndexRecord> {
    const oldPayload = await loadPayload(docId, entryId);
    const skeleton = await reconstructSkeleton(docId, entryId);
    if (!oldPayload || !skeleton) {
      return index;
    }
    const payload = codec.baselineOf(skeleton);
    await ensureBlobRefs(docId, codec.embeddedHashes(payload), new Map());
    await putVersionsRecord({ id: versionPayloadKey(docId, entryId), store: "payload", docId, versionId: entryId, payload });
    await releaseBlobRefs(docId, codec.embeddedHashes(oldPayload), keep);
    const size = measureBytes(payload);
    const versions = index.versions.map((e) => (e.id === entryId ? { ...e, kind: "baseline" as const, parentId: null, size } : e));
    const next: VersionsIndexRecord = { ...index, versions };
    await putVersionsRecord(next);
    return next;
  }

  async function deleteOne(docId: string, entryId: string, index: VersionsIndexRecord, keep: Set<string>): Promise<VersionsIndexRecord> {
    const payload = await loadPayload(docId, entryId);
    await deleteVersionsRecord(versionPayloadKey(docId, entryId));
    const next: VersionsIndexRecord = { ...index, versions: index.versions.filter((e) => e.id !== entryId) };
    await putVersionsRecord(next);
    if (payload) {
      await releaseBlobRefs(docId, codec.embeddedHashes(payload), keep);
    }
    return next;
  }

  /** SPEC-040's prune loop, over core's selection rules: delete the oldest
   * prunable version (re-encoding the oldest retained version that chains
   * through it first), re-measuring after every step; with nothing prunable
   * and still over, latch `paused`. The unretained versions between the
   * pruned one and that dependant go in the same step — they are the
   * next-oldest prunable anyway, and without their parent nothing could
   * restore them. Never rejects — a failed step leaves the index as the
   * store last held it. */
  async function enforceBudget(docId: string, keep: Set<string>, start: VersionsIndexRecord): Promise<VersionsIndexRecord> {
    let index = start;
    try {
      for (;;) {
        if ((await usedBytes(docId, index)) <= HISTORY_BUDGET_BYTES) {
          if (index.paused) {
            index = { ...withoutPausedReason(index), paused: false };
            await putVersionsRecord(index);
          }
          return index;
        }
        const prunable = oldestPrunable(index.versions);
        if (!prunable) {
          if (!index.paused || index.pausedReason !== undefined) {
            index = { ...withoutPausedReason(index), paused: true };
            await putVersionsRecord(index);
          }
          return index;
        }
        const dependent = oldestRetainedDescendant(index.versions, prunable.id);
        if (dependent) {
          index = await reencodeAsBaseline(docId, dependent.id, index, keep);
        }
        for (const id of [prunable.id, ...unretainedChainAfter(index.versions, prunable.id)]) {
          index = await deleteOne(docId, id, index, keep);
        }
      }
    } catch {
      return loadVersionsIndex(docId);
    }
  }

  /** The tail's content rebuilt from the chain (refs only), for a session
   * that has not recorded a version yet. */
  async function storedTail(docId: string, versionId: string): Promise<T | null> {
    const skeleton = await reconstructSkeleton(docId, versionId);
    return skeleton ? codec.tailOf(versionId, skeleton) : null;
  }

  /** The over-budget refusal: latch `paused` with pausedReason "too-large"
   * (written once — a latched index is not rewritten). */
  async function refuseTooLarge(index: VersionsIndexRecord): Promise<RecordVersionResult<T>> {
    let latched = index;
    if (index.pausedReason !== "too-large" || !index.paused) {
      latched = { ...index, paused: true, pausedReason: "too-large" };
      await putVersionsRecord(latched).catch(() => {});
    }
    return { written: false, index: latched, newEntry: null };
  }

  /** Writes one version, then enforces the budget. `pre-restore`/
   * `post-restore` write even while paused (they bracket a restore the user
   * asked for); `cadence`/`commit` respect the latch. A model the codec
   * refuses before encoding, or whose blob set plus one baseline would
   * exceed HISTORY_BUDGET_BYTES, is refused before any write and the index
   * latches `paused` with pausedReason "too-large" — re-checked on each
   * later attempt, since a model can shrink. Never rejects. */
  async function record(options: RecordVersionOptions<M, T>): Promise<RecordVersionResult<T>> {
    const { docId, author, trigger, cachedPrev } = options;
    const capture = codec.capture(options.model);
    let index = await loadVersionsIndex(docId);
    const unconditional = trigger === "pre-restore" || trigger === "post-restore";
    if (index.paused && index.pausedReason !== "too-large" && !unconditional) {
      return { written: false, index, newEntry: null };
    }
    if (codec.tooLargeToEncode?.(capture)) {
      return refuseTooLarge(index);
    }
    let next: T;
    let blobs: BlobSet;
    try {
      // The encode below is the snapshot's one long piece of work: never in
      // the task that asked for it.
      await yieldToMainThread();
      next = await codec.tail(`version-${crypto.randomUUID()}`, capture);
      blobs = await codec.blobs(capture);
    } catch {
      return { written: false, index, newEntry: null };
    }
    const baseline = codec.baseline(next);
    let blobBytes = 0;
    for (const blob of blobs.values()) {
      blobBytes += blob.size;
    }
    if (blobBytes + measureBytes(baseline) > HISTORY_BUDGET_BYTES) {
      return refuseTooLarge(index);
    }

    const tail = index.versions.at(-1) ?? null;
    let prev: T | null = null;
    if (tail) {
      prev = cachedPrev && cachedPrev.id === tail.id ? cachedPrev : await storedTail(docId, tail.id);
    }
    let payload: P = baseline;
    if (prev) {
      const delta = codec.delta(prev, next);
      if (!decideBaseline(gatherChainInfo(index.versions, prev.id), measureBytes(delta))) {
        payload = delta;
      }
    }
    const baselinePayload = codec.isBaseline(payload);
    const entry: VersionIndexEntry = {
      id: next.id,
      timestamp: Date.now(),
      author,
      // Products other than text count their own units (VersionIndexEntry
      // keeps `words` so every stored index stays readable).
      words: 0,
      ...codec.stats(next),
      kind: baselinePayload ? "baseline" : "delta",
      parentId: baselinePayload ? null : (payload.parentId ?? null),
      size: measureBytes(payload),
      trigger,
    };
    try {
      await ensureBlobRefs(docId, codec.embeddedHashes(payload), blobs);
      await putVersionsRecord({ id: versionPayloadKey(docId, next.id), store: "payload", docId, versionId: next.id, payload });
      index = { ...withoutPausedReason(index), versions: [...index.versions, entry], paused: false };
      await putVersionsRecord(index);
    } catch {
      // Abandon this write, leaving no entry in the index; the document's own
      // records are another store and are never touched (SPEC-040 AC-40).
      return { written: false, index: await loadVersionsIndex(docId), newEntry: null };
    }
    index = await enforceBudget(docId, new Set(blobs.keys()), index);
    return { written: true, index, newEntry: next };
  }

  /** Re-runs the budget without adding a version — after a manual delete or
   * "Remove name", so the paused notice clears as soon as it should. */
  async function reconcile(docId: string, live: M): Promise<VersionsIndexRecord> {
    return enforceBudget(docId, await codec.liveHashes(live), await loadVersionsIndex(docId));
  }

  /** Names (or, with a blank name, unnames) a version — index only. */
  async function name(docId: string, versionId: string, label: string): Promise<VersionsIndexRecord> {
    const index = await loadVersionsIndex(docId);
    const trimmed = label.trim();
    const versions = index.versions.map((e) => {
      if (e.id !== versionId) {
        return e;
      }
      const { name: _old, ...rest } = e;
      return trimmed.length > 0 ? { ...rest, name: trimmed } : rest;
    });
    const next = { ...index, versions };
    await putVersionsRecord(next);
    return next;
  }

  /** "Remove name": the version becomes prunable again. */
  async function removeName(docId: string, versionId: string, live: M): Promise<VersionsIndexRecord> {
    await name(docId, versionId, "");
    return reconcile(docId, live);
  }

  /** "Delete": any version, named or not. Its immediate child (if any) is
   * re-encoded as a baseline first so every other version still
   * reconstructs. */
  async function deleteVersion(docId: string, versionId: string, live: M): Promise<VersionsIndexRecord> {
    const keep = await codec.liveHashes(live);
    let index = await loadVersionsIndex(docId);
    const child = index.versions.find((e) => e.parentId === versionId);
    if (child) {
      index = await reencodeAsBaseline(docId, child.id, index, keep);
    }
    await deleteOne(docId, versionId, index, keep);
    return reconcile(docId, live);
  }

  return { record, reconstruct, reconcile, name, removeName, deleteVersion, loadIndex: loadVersionsIndex };
}
