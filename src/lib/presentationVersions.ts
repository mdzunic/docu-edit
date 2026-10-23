// Deck version history (SPEC-103): SPEC-040's chain — baselines, deltas
// against the immediately preceding version, content-addressed blobs, the
// 50 MB budget and its prune loop — recorded for presentations. The chain
// arithmetic is core's (core/versionChain.ts); this module owns the deck's
// payload encoding and its store IO. It lives at the top of src/lib, beside
// documentStore.ts, because lib/presentation may not import ./idb.
//
// Records share the EXISTING `versions` object store and its key layout with
// the word processor's history — `versions:<docId>` (the index),
// `version:<docId>:<versionId>` (a payload), `blob:<docId>:<sha256>` (part
// bytes, reference-counted) — so there is no schema change: a deck's payloads
// are told apart by their `deck-` kinds.
//
// A baseline carries `slides` verbatim (`dirty`, `partName`, `sourceXml` and
// `preserved` included — they decide what the writers rebuild) and every
// other deck-level field; a delta carries the changed slide window and only
// the fields that changed (`null` = cleared, the text side's convention).
// `opaque` and `media` are never inline: each part is a ref to a blob hashed
// with SHA-256 over its base64 text, so a deck's `media` — which duplicates
// image parts already inside `opaque` — collapses onto the same blobs, and a
// delta carries refs only when a ref changed.
//
// The slide window is trimmed with windowOf (core/arrayWindow.ts) over each
// slide's CONTENT, not its object identity. useSlideRender.commit()
// re-identifies only the current slide, but several deck-wide commands
// (resize the deck, apply a theme or a transition to every slide, header &
// footer, sections) mutate other slides in place — a reference-equality window
// would trim those away as untouched and a restore would silently lose them.
// Keyed by content, an untouched slide still trims, a structuredClone'd one
// (after an undo) trims too, and nothing changed is ever left out.
//
// History is best-effort and never the deck's problem: nothing here touches
// the `documents` store or the storage banner, every failure resolves to "not
// written", and the caller never awaits it on an edit's path.
import { windowOf } from "./core/arrayWindow";
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
import type { PresentationModel, Slide } from "./presentation/model";
import { deleteVersionsRecord, getVersionsRecord, putVersionsRecord } from "./idb";

/** Shown in the panel's paused notice when a deck's package alone (with one
 * baseline) is past the history budget — refused before any write. Normal
 * exhaustion (nothing prunable and still over) keeps core's
 * HISTORY_PAUSED_MESSAGE. */
export const DECK_HISTORY_TOO_LARGE_MESSAGE =
  "Version history off — this deck's package is larger than the 50 MB history budget";

// --- record shapes -----------------------------------------------------------

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

/** Every deck-level field but the slides and the two part arrays — title,
 * format, slide size, properties, structureDirty, odpContentEnvelope and the
 * rest of the deck's settings (theme, layouts, header/footer, sections…), so a
 * restore puts back everything the writers read. All optional; in a delta,
 * present only when changed, `null` meaning "changed to absent". */
export type DeckFieldKey = Exclude<keyof PresentationModel, "kind" | "id" | "slides" | "opaque" | "media">;
export type DeckFields = { [K in DeckFieldKey]?: PresentationModel[K] | null };

export interface DeckBaselinePayload {
  kind: "deck-baseline";
  slides: Slide[];
  fields: DeckFields;
  opaque: OpaqueRef[];
  media: MediaRef[];
}

export interface DeckDeltaPayload {
  kind: "deck-delta";
  parentId: string;
  fields: DeckFields;
  /** The parent's slides with `prefixLen` kept from the front, `suffixLen`
   * from the back, and `replaced` between them. */
  slides: { prefixLen: number; suffixLen: number; replaced: Slide[] };
  /** Present only when a ref changed. */
  opaque?: OpaqueRef[];
  media?: MediaRef[];
}

export type DeckPayload = DeckBaselinePayload | DeckDeltaPayload;

export interface DeckVersionPayloadRecord {
  id: string;
  store: "payload";
  docId: string;
  versionId: string;
  payload: DeckPayload;
}

export type DeckVersionStoreRecord = VersionsIndexRecord | DeckVersionPayloadRecord | VersionBlobRecord;

/** What the next snapshot's delta is built against: the content of the last
 * version this session recorded, kept in memory so a delta never has to
 * reconstruct its parent (after a reload it does, from the chain's refs —
 * never its blobs). */
export interface DeckVersionTail {
  id: string;
  slideKeys: string[];
  fieldJson: Record<string, string>;
  opaque: OpaqueRef[];
  media: MediaRef[];
}

export interface RecordDeckVersionOptions {
  docId: string;
  model: PresentationModel;
  author: string;
  trigger: SnapshotTrigger;
  /** The last version recorded this session; null on a fresh session. */
  cachedPrev: DeckVersionTail | null;
}

export interface RecordDeckVersionResult {
  written: boolean;
  index: VersionsIndexRecord;
  newEntry: DeckVersionTail | null;
}

// --- content keys and refs ---------------------------------------------------

const EXCLUDED_FIELDS = new Set(["kind", "id", "slides", "opaque", "media"]);

function deckFieldKeys(model: PresentationModel): DeckFieldKey[] {
  return Object.keys(model).filter((key) => !EXCLUDED_FIELDS.has(key)) as DeckFieldKey[];
}

/** The deck-level fields present on a model, by value. */
function fieldsOf(model: PresentationModel): DeckFields {
  const fields: Record<string, unknown> = {};
  for (const key of deckFieldKeys(model)) {
    if (model[key] !== undefined) {
      fields[key] = model[key];
    }
  }
  return fields as DeckFields;
}

function fieldJsonOf(fields: DeckFields): Record<string, string> {
  const json: Record<string, string> = {};
  for (const [key, value] of Object.entries(fields)) {
    if (value !== undefined && value !== null) {
      json[key] = JSON.stringify(value);
    }
  }
  return json;
}

// Hashing a package is the expensive part of a snapshot, so each part's digest
// is kept beside the exact string it was taken over: an unchanged part is
// hashed once per session, and a part whose `data` was replaced is re-hashed.
const hashCache = new WeakMap<OpaquePart | MediaPart, { data: string; hash: string }>();

async function hashPart(part: OpaquePart | MediaPart): Promise<string> {
  const cached = hashCache.get(part);
  if (cached && cached.data === part.data) {
    return cached.hash;
  }
  const hash = await sha256Hex(part.data);
  hashCache.set(part, { data: part.data, hash });
  return hash;
}

async function opaqueRefsOf(parts: OpaquePart[] | undefined): Promise<OpaqueRef[]> {
  const refs: OpaqueRef[] = [];
  for (const part of parts ?? []) {
    refs.push({ name: part.name, hash: await hashPart(part) });
  }
  return refs;
}

async function mediaRefsOf(parts: MediaPart[] | undefined): Promise<MediaRef[]> {
  const refs: MediaRef[] = [];
  for (const part of parts ?? []) {
    refs.push({ mediaId: part.mediaId, contentType: part.contentType, hash: await hashPart(part) });
  }
  return refs;
}

/** A deck's content at one instant, taken synchronously before a snapshot
 * awaits anything: the live model keeps changing under an edit (commit()
 * re-identifies the current slide inside the array a payload would hold), so
 * a payload is built from these strings, never from the live objects. The
 * part arrays are held by reference — they are replaced, never mutated. */
interface DeckCapture {
  slideKeys: string[];
  fieldJson: Record<string, string>;
  opaque?: OpaquePart[];
  media?: MediaPart[];
}

function captureDeck(model: PresentationModel): DeckCapture {
  return {
    slideKeys: model.slides.map((slide) => JSON.stringify(slide)),
    fieldJson: fieldJsonOf(fieldsOf(model)),
    opaque: model.opaque,
    media: model.media,
  };
}

/** A capture's content as the chain compares it. */
async function tailOf(id: string, capture: DeckCapture): Promise<DeckVersionTail> {
  return {
    id,
    slideKeys: capture.slideKeys,
    fieldJson: capture.fieldJson,
    opaque: await opaqueRefsOf(capture.opaque),
    media: await mediaRefsOf(capture.media),
  };
}

function parseFields(json: Record<string, string>): DeckFields {
  const fields: Record<string, unknown> = {};
  for (const [key, value] of Object.entries(json)) {
    fields[key] = JSON.parse(value);
  }
  return fields as DeckFields;
}

const parseSlide = (key: string): Slide => JSON.parse(key) as Slide;

function sameRefs(a: readonly (OpaqueRef | MediaRef)[], b: readonly (OpaqueRef | MediaRef)[]): boolean {
  return a.length === b.length && a.every((ref, i) => JSON.stringify(ref) === JSON.stringify(b[i]));
}

/** measureBytes of a base64 string, without encoding megabytes to count them:
 * base64 is ASCII and needs no JSON escaping, so it is its length plus the
 * two quotes. */
function base64Bytes(data: string): number {
  return data.length + 2;
}

/** Each distinct blob a deck references, with its bytes — the package and
 * the media deduplicated by hash. */
async function blobsOf(parts: Pick<DeckCapture, "opaque" | "media">): Promise<Map<string, { data: string; contentType?: string }>> {
  const blobs = new Map<string, { data: string; contentType?: string }>();
  for (const part of parts.opaque ?? []) {
    blobs.set(await hashPart(part), { data: part.data });
  }
  for (const part of parts.media ?? []) {
    blobs.set(await hashPart(part), { data: part.data, contentType: part.contentType });
  }
  return blobs;
}

// --- payload build ------------------------------------------------------------

function baselineOf(next: DeckVersionTail): DeckBaselinePayload {
  return {
    kind: "deck-baseline",
    slides: next.slideKeys.map(parseSlide),
    fields: parseFields(next.fieldJson),
    opaque: next.opaque,
    media: next.media,
  };
}

function deltaOf(prev: DeckVersionTail, next: DeckVersionTail): DeckDeltaPayload {
  const window = windowOf(prev.slideKeys, next.slideKeys);
  const fields: Record<string, unknown> = {};
  for (const key of new Set([...Object.keys(prev.fieldJson), ...Object.keys(next.fieldJson)])) {
    if (prev.fieldJson[key] !== next.fieldJson[key]) {
      fields[key] = key in next.fieldJson ? JSON.parse(next.fieldJson[key]) : null;
    }
  }
  const delta: DeckDeltaPayload = {
    kind: "deck-delta",
    parentId: prev.id,
    fields: fields as DeckFields,
    slides: {
      prefixLen: window.prefixLen,
      suffixLen: window.suffixLen,
      replaced: window.afterMid.map(parseSlide),
    },
  };
  if (!sameRefs(prev.opaque, next.opaque)) {
    delta.opaque = next.opaque;
  }
  if (!sameRefs(prev.media, next.media)) {
    delta.media = next.media;
  }
  return delta;
}

/** The blob hashes a payload itself embeds (once each) — what its refcounts
 * count; an inherited ref is its ancestor's reference, not this one's. */
function embeddedHashes(payload: DeckPayload): string[] {
  return [...new Set([...(payload.opaque ?? []), ...(payload.media ?? [])].map((ref) => ref.hash))];
}

// --- reconstruction ---------------------------------------------------------

/** A version without its bytes: slides, fields and part refs. */
interface DeckSkeleton {
  slides: Slide[];
  fields: DeckFields;
  opaque: OpaqueRef[];
  media: MediaRef[];
}

function applyDelta(parent: DeckSkeleton, delta: DeckDeltaPayload): DeckSkeleton {
  const { prefixLen, suffixLen, replaced } = delta.slides;
  const slides = [
    ...parent.slides.slice(0, prefixLen),
    ...replaced,
    ...(suffixLen > 0 ? parent.slides.slice(parent.slides.length - suffixLen) : []),
  ];
  const fields: Record<string, unknown> = { ...parent.fields };
  for (const [key, value] of Object.entries(delta.fields)) {
    if (value === null) {
      delete fields[key];
    } else {
      fields[key] = value;
    }
  }
  return { slides, fields: fields as DeckFields, opaque: delta.opaque ?? parent.opaque, media: delta.media ?? parent.media };
}

async function loadPayload(docId: string, versionId: string): Promise<DeckPayload | null> {
  const record = await getVersionsRecord(versionPayloadKey(docId, versionId));
  const payload = record && record.store === "payload" ? record.payload : null;
  return payload && (payload.kind === "deck-baseline" || payload.kind === "deck-delta") ? payload : null;
}

/** Walks a version's chain back to its baseline (at most CHAIN_DELTA_CAP + 1
 * records) and applies it forward — refs only, no blob read. Null when the
 * chain is missing, malformed or unreadable. */
async function reconstructSkeleton(docId: string, versionId: string): Promise<DeckSkeleton | null> {
  const chain: DeckPayload[] = [];
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
      currentId = payload.kind === "deck-delta" ? payload.parentId : null;
    }
  } catch {
    return null;
  }
  const [base, ...deltas] = chain;
  if (!base || base.kind !== "deck-baseline") {
    return null;
  }
  let skeleton: DeckSkeleton = { slides: base.slides, fields: base.fields, opaque: base.opaque, media: base.media };
  for (const delta of deltas) {
    skeleton = applyDelta(skeleton, delta as DeckDeltaPayload);
  }
  return skeleton;
}

/** Reads each ref's bytes back, reusing a part of `live` whose content is the
 * same (and `live`'s whole array when every part is): a restore whose package
 * did not change keeps the deck's own `opaque` reference, so the store's latch
 * skips rewriting it. */
async function hydrateParts<P extends OpaquePart | MediaPart, R extends OpaqueRef | MediaRef>(
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

/** Reconstructs a deck version's whole model — its blobs read back — or null
 * when its chain or a blob is unreadable. `live` (the open deck) lends any
 * part whose content is unchanged. */
export async function reconstructDeckVersion(
  docId: string,
  versionId: string,
  live?: PresentationModel,
): Promise<PresentationModel | null> {
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
    const opaque = await hydrateParts(
      skeleton.opaque,
      live?.opaque,
      (ref, data): OpaquePart => ({ name: ref.name, data }),
      (part, ref) => part.name === ref.name,
      read,
    );
    const media = await hydrateParts(
      skeleton.media,
      live?.media,
      (ref, data): MediaPart => ({ mediaId: ref.mediaId, contentType: ref.contentType, data }),
      (part, ref) => part.mediaId === ref.mediaId && part.contentType === ref.contentType,
      read,
    );
    const model = { kind: "presentation", id: docId, ...skeleton.fields, slides: skeleton.slides } as PresentationModel;
    if (opaque.length > 0) {
      model.opaque = opaque;
    }
    if (media.length > 0) {
      model.media = media;
    }
    return model;
  } catch {
    return null;
  }
}

async function readBlob(docId: string, hash: string): Promise<string> {
  const record = await getVersionsRecord(versionBlobKey(docId, hash));
  if (!record || record.store !== "blob") {
    throw new Error(`Missing history blob ${hash} for deck ${docId}`);
  }
  return record.data;
}

// --- store IO -------------------------------------------------------------------

/** The deck's index; an empty one when absent or unreadable. */
export async function loadDeckIndex(docId: string): Promise<VersionsIndexRecord> {
  try {
    const record = await getVersionsRecord(versionsIndexKey(docId));
    return record && record.store === "index" ? record : emptyVersionsIndex(docId);
  } catch {
    return emptyVersionsIndex(docId);
  }
}

async function ensureBlobRefs(
  docId: string,
  hashes: string[],
  blobs: Map<string, { data: string; contentType?: string }>,
): Promise<void> {
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
    const blob: VersionBlobRecord = { id: key, store: "blob", docId, hash, data: source.data, size: base64Bytes(source.data), refCount: 1 };
    if (source.contentType !== undefined) {
      blob.contentType = source.contentType;
    }
    await putVersionsRecord(blob);
  }
}

/** Drops one reference per hash. A blob the open deck still references stays
 * at refCount 0 (SPEC-040 AC-29's rule) — the next snapshot re-references it
 * instead of rewriting megabytes — and goes once nothing does. */
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

async function liveHashes(model: PresentationModel): Promise<Set<string>> {
  return new Set((await blobsOf(model)).keys());
}

// --- recording a version -----------------------------------------------------

/** Writes one deck version, then enforces the budget. `pre-restore`/
 * `post-restore` write even while paused (they bracket a restore the user
 * asked for); `cadence`/`commit` respect the latch. Before any write, a deck
 * whose blob set plus one baseline would exceed HISTORY_BUDGET_BYTES is
 * refused and the index latches `paused` with pausedReason "too-large" —
 * re-checked on each later attempt, since media can shrink. Never rejects. */
export async function recordDeckVersion(options: RecordDeckVersionOptions): Promise<RecordDeckVersionResult> {
  const { docId, author, trigger, cachedPrev } = options;
  const capture = captureDeck(options.model);
  let index = await loadDeckIndex(docId);
  const unconditional = trigger === "pre-restore" || trigger === "post-restore";
  if (index.paused && index.pausedReason !== "too-large" && !unconditional) {
    return { written: false, index, newEntry: null };
  }
  let next: DeckVersionTail;
  let blobs: Map<string, { data: string; contentType?: string }>;
  try {
    next = await tailOf(`version-${crypto.randomUUID()}`, capture);
    blobs = await blobsOf(capture);
  } catch {
    return { written: false, index, newEntry: null };
  }
  const baseline = baselineOf(next);
  let blobBytes = 0;
  for (const blob of blobs.values()) {
    blobBytes += base64Bytes(blob.data);
  }
  if (blobBytes + measureBytes(baseline) > HISTORY_BUDGET_BYTES) {
    if (index.pausedReason !== "too-large" || !index.paused) {
      index = { ...index, paused: true, pausedReason: "too-large" };
      await putVersionsRecord(index).catch(() => {});
    }
    return { written: false, index, newEntry: null };
  }

  const tail = index.versions.at(-1) ?? null;
  let prev: DeckVersionTail | null = null;
  if (tail) {
    prev = cachedPrev && cachedPrev.id === tail.id ? cachedPrev : await storedTail(docId, tail.id);
  }
  let payload: DeckPayload = baseline;
  if (prev) {
    const delta = deltaOf(prev, next);
    if (!decideBaseline(gatherChainInfo(index.versions, prev.id), measureBytes(delta))) {
      payload = delta;
    }
  }
  const entry: VersionIndexEntry = {
    id: next.id,
    timestamp: Date.now(),
    author,
    // A deck counts slides, not words (VersionIndexEntry keeps `words` so
    // every stored index stays readable).
    words: 0,
    slides: next.slideKeys.length,
    kind: payload.kind === "deck-baseline" ? "baseline" : "delta",
    parentId: payload.kind === "deck-delta" ? payload.parentId : null,
    size: measureBytes(payload),
    trigger,
  };
  try {
    await ensureBlobRefs(docId, embeddedHashes(payload), blobs);
    await putVersionsRecord({ id: versionPayloadKey(docId, next.id), store: "payload", docId, versionId: next.id, payload });
    const { pausedReason: _cleared, ...rest } = index;
    index = { ...rest, versions: [...index.versions, entry], paused: false };
    await putVersionsRecord(index);
  } catch {
    // Abandon this write, leaving no entry in the index; the deck's own
    // records are another store and are never touched (SPEC-040 AC-40).
    return { written: false, index: await loadDeckIndex(docId), newEntry: null };
  }
  index = await enforceDeckBudget(docId, new Set(blobs.keys()), index);
  return { written: true, index, newEntry: next };
}

/** The tail's content rebuilt from the chain (refs only), for a session that
 * has not recorded a version yet. */
async function storedTail(docId: string, versionId: string): Promise<DeckVersionTail | null> {
  const skeleton = await reconstructSkeleton(docId, versionId);
  if (!skeleton) {
    return null;
  }
  return {
    id: versionId,
    slideKeys: skeleton.slides.map((slide) => JSON.stringify(slide)),
    fieldJson: fieldJsonOf(skeleton.fields),
    opaque: skeleton.opaque,
    media: skeleton.media,
  };
}

// --- budget and pruning ------------------------------------------------------

/** Bytes the deck's history holds: every version's payload plus every blob a
 * version still references. Unlike the text side, a deck counts the blobs it
 * shares with the open deck too — its package is the bulk of what history
 * stores, and the budget is set by it (SPEC-103's worked case). */
async function usedBytes(docId: string, index: VersionsIndexRecord): Promise<number> {
  let total = 0;
  const hashes = new Set<string>();
  for (const entry of index.versions) {
    total += entry.size;
    const payload = await loadPayload(docId, entry.id);
    for (const hash of payload ? embeddedHashes(payload) : []) {
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

/** Re-encodes `entryId` as a standalone baseline so nothing after it depends
 * on anything older (pruning's chain-integrity step). */
async function reencodeAsBaseline(docId: string, entryId: string, index: VersionsIndexRecord, keep: Set<string>): Promise<VersionsIndexRecord> {
  const oldPayload = await loadPayload(docId, entryId);
  const skeleton = await reconstructSkeleton(docId, entryId);
  if (!oldPayload || !skeleton) {
    return index;
  }
  const payload: DeckBaselinePayload = { kind: "deck-baseline", ...skeleton };
  await ensureBlobRefs(docId, embeddedHashes(payload), new Map());
  await putVersionsRecord({ id: versionPayloadKey(docId, entryId), store: "payload", docId, versionId: entryId, payload });
  await releaseBlobRefs(docId, embeddedHashes(oldPayload), keep);
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
    await releaseBlobRefs(docId, embeddedHashes(payload), keep);
  }
  return next;
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

/** SPEC-040's prune loop, over core's selection rules: delete the oldest
 * prunable version (re-encoding the oldest retained version that chains
 * through it first), re-measuring after every step; with nothing prunable
 * and still over, latch `paused`. The unretained versions between the pruned
 * one and that dependant go in the same step — they are the next-oldest
 * prunable anyway, and without their parent nothing could restore them, so
 * the list never offers a version that cannot be restored. Never rejects —
 * a failed step leaves the index as the store last held it. */
async function enforceDeckBudget(docId: string, keep: Set<string>, start: VersionsIndexRecord): Promise<VersionsIndexRecord> {
  let index = start;
  try {
    for (;;) {
      if ((await usedBytes(docId, index)) <= HISTORY_BUDGET_BYTES) {
        if (index.paused) {
          const { pausedReason: _cleared, ...rest } = index;
          index = { ...rest, paused: false };
          await putVersionsRecord(index);
        }
        return index;
      }
      const prunable = oldestPrunable(index.versions);
      if (!prunable) {
        if (!index.paused || index.pausedReason !== undefined) {
          const { pausedReason: _cleared, ...rest } = index;
          index = { ...rest, paused: true };
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
    return loadDeckIndex(docId);
  }
}

/** Re-runs the budget without adding a version — after a manual delete or
 * "Remove name", so the paused notice clears as soon as it should. */
export async function reconcileDeckBudget(docId: string, live: PresentationModel): Promise<VersionsIndexRecord> {
  return enforceDeckBudget(docId, await liveHashes(live), await loadDeckIndex(docId));
}

// --- the panel's actions -------------------------------------------------------

/** Names (or, with a blank name, unnames) a version — index only. */
export async function nameDeckVersion(docId: string, versionId: string, name: string): Promise<VersionsIndexRecord> {
  const index = await loadDeckIndex(docId);
  const trimmed = name.trim();
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
export async function removeDeckVersionName(docId: string, versionId: string, live: PresentationModel): Promise<VersionsIndexRecord> {
  await nameDeckVersion(docId, versionId, "");
  return reconcileDeckBudget(docId, live);
}

/** "Delete": any version, named or not. Its immediate child (if any) is
 * re-encoded as a baseline first so every other version still reconstructs. */
export async function deleteDeckVersion(docId: string, versionId: string, live: PresentationModel): Promise<VersionsIndexRecord> {
  const keep = await liveHashes(live);
  let index = await loadDeckIndex(docId);
  const child = index.versions.find((e) => e.parentId === versionId);
  if (child) {
    index = await reencodeAsBaseline(docId, child.id, index, keep);
  }
  await deleteOne(docId, versionId, index, keep);
  return reconcileDeckBudget(docId, live);
}

// --- the list -------------------------------------------------------------------

export interface DeckVersionListEntry extends VersionIndexEntry {
  /** Slide-count change against the chronologically preceding version; the
   * oldest version's is its own slide count. */
  slideDelta: number;
}

/** Newest first, each with its slide-count delta — the index alone, no
 * payload read. */
export function deckVersionEntries(index: VersionsIndexRecord): DeckVersionListEntry[] {
  return index.versions
    .map((entry, i) => ({
      ...entry,
      slideDelta: (entry.slides ?? 0) - (i === 0 ? 0 : (index.versions[i - 1].slides ?? 0)),
    }))
    .toReversed();
}
