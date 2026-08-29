// Local version history (SPEC-040): periodic and commit-point snapshots of
// the document, stored as a chain of deltas over periodic full baselines in
// the `versions` IndexedDB object store (see lib/idb.ts), budgeted at 50 MB
// per document.
//
// Storage shape follows the spec's two record kinds plus a third for shared
// blobs, all in the one `versions` store (mirroring how lib/documentStore.ts
// already mixes committed/draft records in `documents` by key prefix):
//   `versions:<docId>`              — VersionsIndexRecord, one per document.
//   `version:<docId>:<versionId>`   — VersionPayloadRecord, one per version.
//   `blob:<docId>:<sha256>`         — VersionBlobRecord, shared media/opaque/
//                                      sourceXml bytes, reference-counted.
//
// A version's payload is a baseline (the whole model, sections carried
// verbatim) or a delta (a patch against the immediately preceding version).
// Deltas are block-granular: commands.ts's container-cloning contract means
// an untouched Section/Block keeps its exact array reference across an edit
// (see trackChangesDiff.ts's header comment for the same guarantee used the
// same way), so a delta is built by trimming reference-equal runs from the
// sections/blocks arrays and carrying only the changed middle in full — no
// intra-block patching. This only works session-to-session because the
// "previous version" model is kept as a live object in memory (see
// RecordVersionOptions.cachedPrev); after a reload there is no such
// reference, every block looks changed, and the delta naturally grows large
// enough to trip the wholesale-rewrite rule into a fresh baseline instead —
// a safe degradation, not a special case.
//
// Image bytes (SPEC-024's MediaPart), SPEC-008's OpaquePart parts and
// SourceXml are content-addressed by SHA-256 (crypto.subtle.digest, no new
// dependency) and stored once per document, shared by every version that
// references them — sections/blocks hold no binary data themselves (an
// Image block only carries a mediaId), so only these three top-level model
// fields need blob-reduction; everything else in a delta's `fields` is
// diffed by simple reference/value inequality, one field at a time, the same
// coarse "carry the whole changed thing" granularity as blocks.
import {
  type Block,
  type DocumentModel,
  type OpaquePart,
  type Section,
  type SourceXml,
} from "./documentModel";
import { ensureMediaEncoded, type MediaPart } from "./core/media";
import { windowOf } from "./core/arrayWindow";
import { documentStats } from "./wordCount";
import { deleteVersionsRecord, getVersionsRecord, putVersionsRecord } from "./idb";
import { MAX_FILE_SIZE_BYTES } from "./openDocument";

// --- constants ---------------------------------------------------------------

/** The per-document history budget — the same ceiling SPEC-009 puts on a
 * single file the app will open (see the spec's "The 50 MB budget"). */
export const HISTORY_BUDGET_BYTES = MAX_FILE_SIZE_BYTES;

/** Baseline trigger 2: the reconstruction bound (at most 51 records read and
 * applied for any retained version). */
export const CHAIN_DELTA_CAP = 50;

/** Baseline trigger 4: a delta at or above this fraction of its chain's
 * baseline size stores as a fresh baseline instead. */
export const WHOLESALE_REWRITE_RATIO = 0.5;

/** Pruning never removes the newest this-many versions. */
export const RETAINED_RECENT_COUNT = 10;

/** The exact notice text AC-38 requires. */
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
   * measured, not estimated (AC-44), and what the budget/pruning arithmetic
   * and the four baseline triggers are computed from. */
  size: number;
}

export interface VersionsIndexRecord {
  id: string;
  store: "index";
  docId: string;
  /** Chronological, oldest first — AC-10's "newest first" is a display
   * concern, applied where the list renders. */
  versions: VersionIndexEntry[];
  /** True once a snapshot attempt found nothing prunable and the budget was
   * still exceeded (AC-37/38); cleared the next time usage drops back under
   * budget, whether from a new prune or a manual delete/unname. */
  paused: boolean;
}

interface MediaRef {
  mediaId: string;
  contentType: string;
  hash: string;
}

interface OpaqueRef {
  name: string;
  hash: string;
}

/** The top-level `DocumentModel` fields diffed/stored as one coarse unit
 * each — no per-field sub-patching, same principle as block-granular
 * deltas. `media`/`opaque`/`sourceXml` are represented by their
 * content-hash refs instead of inline bytes. */
type FieldKey =
  | "title"
  | "defaultStyles"
  | "styles"
  | "listDefinitions"
  | "notes"
  | "footnoteSettings"
  | "endnoteSettings"
  | "format"
  | "suppressConversionReport"
  | "measurementUnit"
  | "showRuler"
  | "showVerticalRuler"
  | "showFormattingMarks"
  | "comments"
  | "trackChangesRecording"
  | "mediaRefs"
  | "opaqueRefs"
  | "sourceXmlHash";

const FIELD_KEYS: readonly FieldKey[] = [
  "title",
  "defaultStyles",
  "styles",
  "listDefinitions",
  "notes",
  "footnoteSettings",
  "endnoteSettings",
  "format",
  "suppressConversionReport",
  "measurementUnit",
  "showRuler",
  "showVerticalRuler",
  "showFormattingMarks",
  "comments",
  "trackChangesRecording",
];

type Fields = Partial<Record<FieldKey, unknown>>;

interface BlockRangePatch {
  prefixLen: number;
  suffixLen: number;
  middle: Block[];
}

interface SectionPatch {
  index: number;
  /** Full replacement of every non-`blocks` field (present only when one of
   * them changed) — the same coarse, no-sub-patching rule as top-level
   * fields. */
  meta?: Omit<Section, "blocks">;
  blocks?: BlockRangePatch;
  /** Full replacement section, used only where the sections array's own
   * length changed within the touched window and before/after can't be
   * index-paired. */
  full?: Section;
}

interface SectionsPatch {
  prefixLen: number;
  suffixLen: number;
  afterCount: number;
  patches: SectionPatch[];
}

export interface BaselinePayload {
  kind: "baseline";
  sections: Section[];
  fields: Fields;
}

export interface DeltaPayload {
  kind: "delta";
  parentId: string;
  sections: SectionsPatch;
  fields: Fields;
}

export interface VersionPayloadRecord {
  id: string;
  store: "payload";
  docId: string;
  versionId: string;
  payload: BaselinePayload | DeltaPayload;
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

export type VersionStoreRecord = VersionsIndexRecord | VersionPayloadRecord | VersionBlobRecord;

function indexKey(docId: string): string {
  return `versions:${docId}`;
}

function payloadKey(docId: string, versionId: string): string {
  return `version:${docId}:${versionId}`;
}

function blobKey(docId: string, hash: string): string {
  return `blob:${docId}:${hash}`;
}

// --- measurement ---------------------------------------------------------

const encoder = new TextEncoder();

function measureBytes(value: unknown): number {
  return encoder.encode(JSON.stringify(value)).length;
}

// --- content hashing (media/opaque/sourceXml dedup) -----------------------

async function sha256Hex(text: string): Promise<string> {
  const digest = await crypto.subtle.digest("SHA-256", encoder.encode(text));
  return Array.from(new Uint8Array(digest))
    .map((byte) => byte.toString(16).padStart(2, "0"))
    .join("");
}

// Hashing base64 image/opaque bytes is the expensive part of taking a
// snapshot; these caches (keyed by object identity, which stays stable for
// an unmodified part across snapshots — see the header comment) mean an
// unchanged image is hashed once per session, not once per 5-minute
// snapshot.
const mediaHashCache = new WeakMap<MediaPart, string>();
const opaqueHashCache = new WeakMap<OpaquePart, string>();
const sourceXmlHashCache = new WeakMap<SourceXml, string>();

async function hashMediaPart(part: MediaPart): Promise<string> {
  // SPEC-047: force-encode a still-deferred part before reading its bytes, so
  // a snapshot taken during the deferral window hashes the real image rather
  // than empty `data` (which would also poison this per-part hash cache).
  ensureMediaEncoded(part);
  const cached = mediaHashCache.get(part);
  if (cached) {
    return cached;
  }
  const hash = await sha256Hex(part.data);
  mediaHashCache.set(part, hash);
  return hash;
}

async function hashOpaquePart(part: OpaquePart): Promise<string> {
  const cached = opaqueHashCache.get(part);
  if (cached) {
    return cached;
  }
  const hash = await sha256Hex(part.data);
  opaqueHashCache.set(part, hash);
  return hash;
}

async function hashSourceXml(sourceXml: SourceXml): Promise<string> {
  const cached = sourceXmlHashCache.get(sourceXml);
  if (cached) {
    return cached;
  }
  const hash = await sha256Hex(JSON.stringify(sourceXml));
  sourceXmlHashCache.set(sourceXml, hash);
  return hash;
}

interface BlobBytes {
  hash: string;
  contentType?: string;
  data: string;
}

/** Every blob a set of media/opaque/sourceXml refs points at, with the
 * actual bytes attached — collected up front so writing them is a plain
 * loop, no re-hashing. */
async function collectBlobBytes(model: DocumentModel): Promise<Map<string, BlobBytes>> {
  const out = new Map<string, BlobBytes>();
  for (const part of model.media ?? []) {
    const hash = await hashMediaPart(part);
    out.set(hash, { hash, contentType: part.contentType, data: part.data });
  }
  for (const part of model.opaque ?? []) {
    const hash = await hashOpaquePart(part);
    out.set(hash, { hash, data: part.data });
  }
  if (model.sourceXml) {
    const hash = await hashSourceXml(model.sourceXml);
    out.set(hash, { hash, data: JSON.stringify(model.sourceXml) });
  }
  return out;
}

async function mediaRefsOf(model: DocumentModel): Promise<MediaRef[] | undefined> {
  if (!model.media || model.media.length === 0) {
    return undefined;
  }
  const refs: MediaRef[] = [];
  for (const part of model.media) {
    refs.push({ mediaId: part.mediaId, contentType: part.contentType, hash: await hashMediaPart(part) });
  }
  return refs;
}

async function opaqueRefsOf(model: DocumentModel): Promise<OpaqueRef[] | undefined> {
  if (!model.opaque || model.opaque.length === 0) {
    return undefined;
  }
  const refs: OpaqueRef[] = [];
  for (const part of model.opaque) {
    refs.push({ name: part.name, hash: await hashOpaquePart(part) });
  }
  return refs;
}

async function sourceXmlHashOf(model: DocumentModel): Promise<string | undefined> {
  return model.sourceXml ? await hashSourceXml(model.sourceXml) : undefined;
}

// --- section/block delta -----------------------------------------------------

function sectionMetaOf(section: Section): Omit<Section, "blocks"> {
  const { blocks: _blocks, ...meta } = section;
  return meta;
}

function sectionMetaEqual(a: Section, b: Section): boolean {
  return (
    a.start === b.start &&
    a.pageSize === b.pageSize &&
    a.margins === b.margins &&
    a.columns === b.columns &&
    a.docxSectPr === b.docxSectPr &&
    a.headers === b.headers &&
    a.footers === b.footers &&
    a.titlePage === b.titlePage &&
    a.evenAndOddHeaders === b.evenAndOddHeaders &&
    a.pageNumberStart === b.pageNumberStart &&
    a.opaqueBlocks === b.opaqueBlocks
  );
}

function buildBlockRangePatch(before: Block[], after: Block[]): BlockRangePatch {
  const w = windowOf(before, after);
  return { prefixLen: w.prefixLen, suffixLen: w.suffixLen, middle: w.afterMid };
}

function applyBlockRangePatch(before: Block[], patch: BlockRangePatch): Block[] {
  const prefix = before.slice(0, patch.prefixLen);
  const suffix = patch.suffixLen > 0 ? before.slice(before.length - patch.suffixLen) : [];
  return [...prefix, ...patch.middle, ...suffix];
}

function buildSectionsPatch(before: Section[], after: Section[]): SectionsPatch {
  const w = windowOf(before, after);
  const patches: SectionPatch[] = [];
  if (w.beforeMid.length === w.afterMid.length) {
    for (let i = 0; i < w.beforeMid.length; i++) {
      const prevSection = w.beforeMid[i];
      const nextSection = w.afterMid[i];
      if (prevSection === nextSection) {
        continue;
      }
      const index = w.prefixLen + i;
      const metaChanged = !sectionMetaEqual(prevSection, nextSection);
      const blocksChanged = prevSection.blocks !== nextSection.blocks;
      patches.push({
        index,
        meta: metaChanged ? sectionMetaOf(nextSection) : undefined,
        blocks: blocksChanged ? buildBlockRangePatch(prevSection.blocks, nextSection.blocks) : undefined,
      });
    }
  } else {
    for (let i = 0; i < w.afterMid.length; i++) {
      patches.push({ index: w.prefixLen + i, full: w.afterMid[i] });
    }
  }
  return { prefixLen: w.prefixLen, suffixLen: w.suffixLen, afterCount: after.length, patches };
}

function applySectionsPatch(before: Section[], patch: SectionsPatch): Section[] {
  const after: Section[] = Array.from({ length: patch.afterCount });
  for (let i = 0; i < patch.prefixLen; i++) {
    after[i] = before[i];
  }
  for (let i = 0; i < patch.suffixLen; i++) {
    after[patch.afterCount - 1 - i] = before[before.length - 1 - i];
  }
  const byIndex = new Map(patch.patches.map((p) => [p.index, p]));
  for (let i = patch.prefixLen; i < patch.afterCount - patch.suffixLen; i++) {
    const p = byIndex.get(i);
    if (!p) {
      // Unchanged-by-reference section inside an otherwise-touched window
      // (only possible when the sections array's own length didn't change).
      after[i] = before[i];
      continue;
    }
    if (p.full) {
      after[i] = p.full;
      continue;
    }
    const prevSection = before[i];
    const blocks = p.blocks ? applyBlockRangePatch(prevSection.blocks, p.blocks) : prevSection.blocks;
    // `p.meta` is a full replacement of every non-`blocks` field (see
    // SectionPatch's doc comment), not a partial patch to merge onto
    // `prevSection` — a field the newer section cleared (e.g. `docxSectPr`
    // deleted by a page-setup edit) is simply absent from `meta`, and
    // merging onto `prevSection` would let that stale value survive.
    after[i] = p.meta ? { ...p.meta, blocks } : { ...prevSection, blocks };
  }
  return after;
}

// --- top-level field delta ---------------------------------------------------

function readAllFields(model: DocumentModel, mediaRefs: MediaRef[] | undefined, opaqueRefs: OpaqueRef[] | undefined, sourceXmlHash: string | undefined): Fields {
  const fields: Fields = {};
  for (const key of FIELD_KEYS) {
    const value = model[key as keyof DocumentModel];
    if (value !== undefined) {
      fields[key] = value;
    }
  }
  if (mediaRefs) {
    fields.mediaRefs = mediaRefs;
  }
  if (opaqueRefs) {
    fields.opaqueRefs = opaqueRefs;
  }
  if (sourceXmlHash) {
    fields.sourceXmlHash = sourceXmlHash;
  }
  return fields;
}

/** `null` in a delta's `fields` means "changed to absent"; a present,
 * non-null value is the new value; an absent key means "unchanged, inherit
 * from the parent" — JSON drops `undefined`-valued keys, so `null` is the
 * only way to represent "cleared" across the round trip. */
async function diffFields(prev: DocumentModel, next: DocumentModel): Promise<Fields> {
  const fields: Fields = {};
  for (const key of FIELD_KEYS) {
    const a = prev[key as keyof DocumentModel];
    const b = next[key as keyof DocumentModel];
    if (a !== b) {
      fields[key] = b === undefined ? null : b;
    }
  }
  if (prev.media !== next.media) {
    fields.mediaRefs = (await mediaRefsOf(next)) ?? null;
  }
  if (prev.opaque !== next.opaque) {
    fields.opaqueRefs = (await opaqueRefsOf(next)) ?? null;
  }
  if (prev.sourceXml !== next.sourceXml) {
    fields.sourceXmlHash = (await sourceXmlHashOf(next)) ?? null;
  }
  return fields;
}

async function readBlobData(docId: string, hash: string): Promise<string> {
  const record = await getVersionsRecord(blobKey(docId, hash));
  if (!record || record.store !== "blob") {
    throw new Error(`Missing history blob ${hash} for document ${docId}`);
  }
  return record.data;
}

async function hydrateBlobFields(docId: string, base: DocumentModel, fields: Fields): Promise<DocumentModel> {
  const model = { ...base };
  if ("mediaRefs" in fields) {
    const refs = fields.mediaRefs as MediaRef[] | null;
    model.media = refs
      ? await Promise.all(refs.map(async (ref): Promise<MediaPart> => ({ mediaId: ref.mediaId, contentType: ref.contentType, data: await readBlobData(docId, ref.hash) })))
      : undefined;
  }
  if ("opaqueRefs" in fields) {
    const refs = fields.opaqueRefs as OpaqueRef[] | null;
    model.opaque = refs
      ? await Promise.all(refs.map(async (ref): Promise<OpaquePart> => ({ name: ref.name, data: await readBlobData(docId, ref.hash) })))
      : undefined;
  }
  if ("sourceXmlHash" in fields) {
    const hash = fields.sourceXmlHash as string | null;
    model.sourceXml = hash ? (JSON.parse(await readBlobData(docId, hash)) as SourceXml) : undefined;
  }
  return model;
}

function applyScalarFields(base: DocumentModel, fields: Fields): DocumentModel {
  const model: Record<string, unknown> = { ...base };
  for (const key of FIELD_KEYS) {
    if (key in fields) {
      const value = fields[key];
      model[key] = value === null ? undefined : value;
    }
  }
  return model as unknown as DocumentModel;
}

// --- payload build/hydrate ---------------------------------------------------

async function buildBaselinePayload(model: DocumentModel): Promise<BaselinePayload> {
  const [mediaRefs, opaqueRefs, sourceXmlHash] = await Promise.all([mediaRefsOf(model), opaqueRefsOf(model), sourceXmlHashOf(model)]);
  return {
    kind: "baseline",
    sections: model.sections,
    fields: readAllFields(model, mediaRefs, opaqueRefs, sourceXmlHash),
  };
}

async function buildDeltaPayload(prev: DocumentModel, next: DocumentModel, parentId: string): Promise<DeltaPayload> {
  const [sections, fields] = await Promise.all([
    Promise.resolve(buildSectionsPatch(prev.sections, next.sections)),
    diffFields(prev, next),
  ]);
  return { kind: "delta", parentId, sections, fields };
}

async function hydrateBaseline(docId: string, payload: BaselinePayload): Promise<DocumentModel> {
  const skeleton = applyScalarFields({} as DocumentModel, payload.fields);
  const hydrated = await hydrateBlobFields(docId, skeleton, payload.fields);
  hydrated.sections = payload.sections;
  return hydrated;
}

async function hydrateDelta(docId: string, parent: DocumentModel, payload: DeltaPayload): Promise<DocumentModel> {
  const withFields = applyScalarFields(parent, payload.fields);
  const withBlobs = await hydrateBlobFields(docId, withFields, payload.fields);
  withBlobs.sections = applySectionsPatch(parent.sections, payload.sections);
  return withBlobs;
}

/** Every hash a single payload embeds directly — used for blob refcounting
 * (a payload "references" only what it itself carries; an inherited value
 * is the ancestor's reference, not this one's — see the header comment on
 * AC-29's delete trigger). */
function embeddedHashes(payload: BaselinePayload | DeltaPayload): string[] {
  const hashes: string[] = [];
  const mediaRefs = payload.fields.mediaRefs as MediaRef[] | null | undefined;
  const opaqueRefs = payload.fields.opaqueRefs as OpaqueRef[] | null | undefined;
  const sourceXmlHash = payload.fields.sourceXmlHash as string | null | undefined;
  if (mediaRefs) {
    hashes.push(...mediaRefs.map((r) => r.hash));
  }
  if (opaqueRefs) {
    hashes.push(...opaqueRefs.map((r) => r.hash));
  }
  if (sourceXmlHash) {
    hashes.push(sourceXmlHash);
  }
  return hashes;
}

// --- store IO -----------------------------------------------------------------

function emptyIndex(docId: string): VersionsIndexRecord {
  return { id: indexKey(docId), store: "index", docId, versions: [], paused: false };
}

export async function loadIndex(docId: string): Promise<VersionsIndexRecord> {
  try {
    const record = await getVersionsRecord(indexKey(docId));
    return record && record.store === "index" ? record : emptyIndex(docId);
  } catch {
    return emptyIndex(docId);
  }
}

async function loadPayload(docId: string, versionId: string): Promise<BaselinePayload | DeltaPayload | null> {
  const record = await getVersionsRecord(payloadKey(docId, versionId));
  return record && record.store === "payload" ? record.payload : null;
}

/** Reconstructs a version's full model by loading its chain's baseline and
 * applying every delta up to it (AC-41: capped at 50 deltas, so at most 51
 * records). Returns null when the chain is missing or unreadable. */
export async function reconstructVersion(docId: string, versionId: string): Promise<DocumentModel | null> {
  const chain: (BaselinePayload | DeltaPayload)[] = [];
  let currentId: string | null = versionId;
  const seen = new Set<string>();
  while (currentId) {
    if (seen.has(currentId)) {
      return null;
    }
    seen.add(currentId);
    let payload: BaselinePayload | DeltaPayload | null;
    try {
      payload = await loadPayload(docId, currentId);
    } catch {
      return null;
    }
    if (!payload) {
      return null;
    }
    chain.unshift(payload);
    currentId = payload.kind === "delta" ? payload.parentId : null;
  }
  if (chain.length === 0 || chain[0].kind !== "baseline") {
    return null;
  }
  try {
    let model = await hydrateBaseline(docId, chain[0]);
    for (let i = 1; i < chain.length; i++) {
      model = await hydrateDelta(docId, model, chain[i] as DeltaPayload);
    }
    return model;
  } catch {
    return null;
  }
}

async function ensureBlobRefs(docId: string, model: DocumentModel, hashes: string[]): Promise<void> {
  if (hashes.length === 0) {
    return;
  }
  const bytes = await collectBlobBytes(model);
  for (const hash of hashes) {
    const key = blobKey(docId, hash);
    const existing = await getVersionsRecord(key);
    if (existing && existing.store === "blob") {
      await putVersionsRecord({ ...existing, refCount: existing.refCount + 1 });
      continue;
    }
    const source = bytes.get(hash);
    if (!source) {
      continue;
    }
    await putVersionsRecord({
      id: key,
      store: "blob",
      docId,
      hash,
      contentType: source.contentType,
      data: source.data,
      size: measureBytes(source.data),
      refCount: 1,
    });
  }
}

async function releaseBlobRefs(docId: string, hashes: string[], keepHashes: Set<string>): Promise<void> {
  for (const hash of hashes) {
    const key = blobKey(docId, hash);
    const existing = await getVersionsRecord(key);
    if (!existing || existing.store !== "blob") {
      continue;
    }
    const refCount = existing.refCount - 1;
    if (refCount > 0) {
      await putVersionsRecord({ ...existing, refCount });
    } else if (!keepHashes.has(hash)) {
      // AC-29: only delete once the current document no longer references
      // the same content by hash either.
      await deleteVersionsRecord(key);
    } else {
      await putVersionsRecord({ ...existing, refCount: 0 });
    }
  }
}

async function currentModelHashes(model: DocumentModel): Promise<Set<string>> {
  const [mediaRefs, opaqueRefs, sourceXmlHash] = await Promise.all([mediaRefsOf(model), opaqueRefsOf(model), sourceXmlHashOf(model)]);
  const hashes = new Set<string>();
  for (const ref of mediaRefs ?? []) {
    hashes.add(ref.hash);
  }
  for (const ref of opaqueRefs ?? []) {
    hashes.add(ref.hash);
  }
  if (sourceXmlHash) {
    hashes.add(sourceXmlHash);
  }
  return hashes;
}

// --- baseline/delta decision (AC-23, checked in this order) ------------------

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

/** The four baseline triggers, checked in order (AC-23); returns false only
 * when none holds, i.e. the snapshot should be a delta (AC-24). */
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

// --- recording a version -------------------------------------------------

export interface RecordVersionOptions {
  docId: string;
  model: DocumentModel;
  author: string;
  trigger: SnapshotTrigger;
  /** The in-memory model object from the last version taken *this session*,
   * enabling the reference-equality fast path (see the header comment).
   * Pass null when there is none yet (fresh session). */
  cachedPrev: { id: string; model: DocumentModel } | null;
}

export interface RecordVersionResult {
  written: boolean;
  index: VersionsIndexRecord;
  newEntry: { id: string; model: DocumentModel } | null;
}

/** Writes one version (AC-1..AC-5, AC-22..AC-28), then enforces the budget
 * (AC-30..AC-38). `pre-restore`/`post-restore` triggers always write, even
 * while paused, since they bracket a restore the user just asked for;
 * `cadence`/`commit` respect the paused flag (AC-37). */
export async function recordVersion(options: RecordVersionOptions): Promise<RecordVersionResult> {
  const { docId, model, author, trigger, cachedPrev } = options;
  let index = await loadIndex(docId);
  const unconditional = trigger === "pre-restore" || trigger === "post-restore";
  if (index.paused && !unconditional) {
    return { written: false, index, newEntry: null };
  }

  const tail = index.versions.length > 0 ? index.versions[index.versions.length - 1] : null;
  let prevModel: DocumentModel | null = null;
  if (tail) {
    prevModel = cachedPrev && cachedPrev.id === tail.id ? cachedPrev.model : await reconstructVersion(docId, tail.id);
  }
  const chainInfo = tail ? gatherChainInfo(index.versions, tail.id) : null;

  let deltaCandidate: DeltaPayload | null = null;
  let deltaBytes = 0;
  if (prevModel && tail) {
    deltaCandidate = await buildDeltaPayload(prevModel, model, tail.id);
    deltaBytes = measureBytes(deltaCandidate);
  }
  const payload: BaselinePayload | DeltaPayload =
    deltaCandidate && !decideBaseline(chainInfo, deltaBytes) ? deltaCandidate : await buildBaselinePayload(model);

  const size = measureBytes(payload);
  const versionId = `version-${crypto.randomUUID()}`;
  const words = documentStats(model, { includeNotes: false }).words;
  const entry: VersionIndexEntry = {
    id: versionId,
    timestamp: Date.now(),
    author,
    words,
    kind: payload.kind,
    parentId: payload.kind === "delta" ? payload.parentId : null,
    size,
  };

  try {
    await ensureBlobRefs(docId, model, embeddedHashes(payload));
    await putVersionsRecord({ id: payloadKey(docId, versionId), store: "payload", docId, versionId, payload });
    index = { ...index, versions: [...index.versions, entry], paused: false };
    await putVersionsRecord(index);
  } catch {
    // AC-40: abandon this write; never touch the documents store, and never
    // leave a half-written entry in the index.
    return { written: false, index: await loadIndex(docId), newEntry: null };
  }

  index = await enforceBudget(docId, model, index);
  return { written: true, index, newEntry: { id: versionId, model } };
}

// --- pruning and chain integrity (AC-32..AC-37) -------------------------

function isRetained(entries: VersionIndexEntry[], entry: VersionIndexEntry): boolean {
  if (entry.name) {
    return true;
  }
  const newest = entries.slice(-RETAINED_RECENT_COUNT);
  return newest.includes(entry);
}

function oldestPrunable(entries: VersionIndexEntry[]): VersionIndexEntry | null {
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
function oldestRetainedDescendant(entries: VersionIndexEntry[], fromId: string): VersionIndexEntry | null {
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

/** Re-encodes `entry` in place as a standalone baseline, so nothing after it
 * in the chain depends on anything older (pruning step 2). Reuses the
 * payload it's replacing to know which blob hashes to release. */
async function reencodeAsBaseline(docId: string, entryId: string, index: VersionsIndexRecord): Promise<VersionsIndexRecord> {
  const oldPayload = await loadPayload(docId, entryId);
  const model = await reconstructVersion(docId, entryId);
  if (!oldPayload || !model) {
    return index; // unreadable — leave it; the outer loop will stop on it.
  }
  const newPayload = await buildBaselinePayload(model);
  const oldHashes = embeddedHashes(oldPayload);
  const newHashes = embeddedHashes(newPayload);
  await ensureBlobRefs(docId, model, newHashes);
  await putVersionsRecord({ id: payloadKey(docId, entryId), store: "payload", docId, versionId: entryId, payload: newPayload });
  const currentHashes = await currentModelHashes(model);
  await releaseBlobRefs(docId, oldHashes, currentHashes);
  const size = measureBytes(newPayload);
  const versions = index.versions.map((e) => (e.id === entryId ? { ...e, kind: "baseline" as const, parentId: null, size } : e));
  const next: VersionsIndexRecord = { ...index, versions };
  await putVersionsRecord(next);
  return next;
}

async function deletePrunedVersion(docId: string, entryId: string, index: VersionsIndexRecord, currentModel: DocumentModel): Promise<VersionsIndexRecord> {
  const payload = await loadPayload(docId, entryId);
  await deleteVersionsRecord(payloadKey(docId, entryId));
  const versions = index.versions.filter((e) => e.id !== entryId);
  const next: VersionsIndexRecord = { ...index, versions };
  await putVersionsRecord(next);
  if (payload) {
    const currentHashes = await currentModelHashes(currentModel);
    await releaseBlobRefs(docId, embeddedHashes(payload), currentHashes);
  }
  return next;
}

async function computeUsedBytes(docId: string, index: VersionsIndexRecord, currentModel: DocumentModel): Promise<number> {
  let total = 0;
  for (const entry of index.versions) {
    total += entry.size;
  }
  const currentHashes = await currentModelHashes(currentModel);
  const hashes = new Set<string>();
  for (const entry of index.versions) {
    const payload = await loadPayload(docId, entry.id);
    if (payload) {
      for (const hash of embeddedHashes(payload)) {
        hashes.add(hash);
      }
    }
  }
  for (const hash of hashes) {
    if (currentHashes.has(hash)) {
      continue; // AC-30: not counted — also part of the live document.
    }
    const blob = await getVersionsRecord(blobKey(docId, hash));
    if (blob && blob.store === "blob" && blob.refCount > 0) {
      total += blob.size;
    }
  }
  return total;
}

/** The pruning loop (AC-32..AC-37): repeatedly deletes the oldest prunable
 * version, re-encoding a dependent retained version first when needed, and
 * re-measures after every step rather than pre-computing a drop count. */
async function enforceBudget(docId: string, currentModel: DocumentModel, startIndex: VersionsIndexRecord): Promise<VersionsIndexRecord> {
  let index = startIndex;
  for (;;) {
    const used = await computeUsedBytes(docId, index, currentModel);
    if (used <= HISTORY_BUDGET_BYTES) {
      if (index.paused) {
        index = { ...index, paused: false };
        await putVersionsRecord(index);
      }
      return index;
    }
    const prunable = oldestPrunable(index.versions);
    if (!prunable) {
      if (!index.paused) {
        index = { ...index, paused: true };
        await putVersionsRecord(index);
      }
      return index;
    }
    const dependent = oldestRetainedDescendant(index.versions, prunable.id);
    if (dependent) {
      index = await reencodeAsBaseline(docId, dependent.id, index);
    }
    index = await deletePrunedVersion(docId, prunable.id, index, currentModel);
  }
}

/** Re-runs the budget check without adding a version — after a manual
 * delete or "remove name" (spec: "Removing a name is what makes a version
 * prunable again"), so the paused notice clears as soon as it should. */
export async function reconcileBudget(docId: string, currentModel: DocumentModel): Promise<VersionsIndexRecord> {
  const index = await loadIndex(docId);
  return enforceBudget(docId, currentModel, index);
}

// --- manual pane actions ---------------------------------------------------

export async function nameVersion(docId: string, versionId: string, name: string): Promise<VersionsIndexRecord> {
  const index = await loadIndex(docId);
  const trimmed = name.trim();
  const versions = index.versions.map((e) => (e.id === versionId ? { ...e, name: trimmed.length > 0 ? trimmed : undefined } : e));
  const next = { ...index, versions };
  await putVersionsRecord(next);
  return next;
}

export async function removeVersionName(docId: string, versionId: string, currentModel: DocumentModel): Promise<VersionsIndexRecord> {
  await nameVersion(docId, versionId, "");
  return reconcileBudget(docId, currentModel);
}

/** Manual "Delete version" — unlike automatic pruning, any version can be
 * targeted (named or not), so chain integrity is preserved by re-encoding
 * the version's own immediate child (if any) rather than searching for the
 * nearest retained descendant: with only one deletion happening, every
 * other surviving version must stay reconstructable, not just the
 * protected ones. */
export async function deleteVersion(docId: string, versionId: string, currentModel: DocumentModel): Promise<VersionsIndexRecord> {
  let index = await loadIndex(docId);
  const child = index.versions.find((e) => e.parentId === versionId);
  if (child) {
    index = await reencodeAsBaseline(docId, child.id, index);
  }
  index = await deletePrunedVersion(docId, versionId, index, currentModel);
  return reconcileBudget(docId, currentModel);
}

// --- word-count delta (AC-9) -------------------------------------------------

export interface VersionListEntry extends VersionIndexEntry {
  /** Word-count change vs. the immediately preceding version chronologically
   * (AC-9); the oldest version's delta is its own word count. */
  wordDelta: number;
}

/** Newest-first (AC-10) with each entry's word-count delta attached —
 * arithmetic over the small index array only (AC-43: no payload reads). */
export function versionListEntries(index: VersionsIndexRecord): VersionListEntry[] {
  const withDeltas: VersionListEntry[] = index.versions.map((entry, i) => ({
    ...entry,
    wordDelta: i === 0 ? entry.words : entry.words - index.versions[i - 1].words,
  }));
  return withDeltas.reverse();
}
