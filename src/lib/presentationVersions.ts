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
//
// SPEC-138 moved the recording loop itself — chain, blobs, budget, pruning,
// the panel's actions — to versionRecorder.ts; this module is the deck's
// VersionCodec over it and keeps every export it had.
import { windowOf } from "./core/arrayWindow";
import type { MediaPart } from "./core/media";
import type { OpaquePart } from "./core/packageParts";
import type { SnapshotTrigger, VersionBlobRecord, VersionIndexEntry, VersionsIndexRecord } from "./core/versionChain";
import type { PresentationModel, Slide } from "./presentation/model";
import {
  createVersionRecorder,
  hydrateMedia,
  hydrateOpaque,
  mediaRefsOf,
  opaqueRefsOf,
  partBlobsOf,
  sameRefs,
  type MediaRef,
  type OpaqueRef,
  type VersionCodec,
  type VersionPayload,
} from "./versionRecorder";

/** Shown in the panel's paused notice when a deck's package alone (with one
 * baseline) is past the history budget — refused before any write. Normal
 * exhaustion (nothing prunable and still over) keeps core's
 * HISTORY_PAUSED_MESSAGE. */
export const DECK_HISTORY_TOO_LARGE_MESSAGE =
  "Version history off — this deck's package is larger than the 50 MB history budget";

// --- record shapes -----------------------------------------------------------

// A retained package part and an image/clip part by content — versionRecorder's
// refs, shared with the workbook's history.
export type { MediaRef, OpaqueRef };

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

function parseFields(json: Record<string, string>): DeckFields {
  const fields: Record<string, unknown> = {};
  for (const [key, value] of Object.entries(json)) {
    fields[key] = JSON.parse(value);
  }
  return fields as DeckFields;
}

const parseSlide = (key: string): Slide => JSON.parse(key) as Slide;

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

/** The deck's codec for versionRecorder's loop. */
const deckCodec: VersionCodec<PresentationModel, DeckCapture, DeckVersionTail, DeckPayload, DeckSkeleton> = {
  owns: (payload: VersionPayload): payload is DeckPayload => payload.kind === "deck-baseline" || payload.kind === "deck-delta",
  isBaseline: (payload) => payload.kind === "deck-baseline",
  capture: captureDeck,
  /** A capture's content as the chain compares it. */
  tail: async (id, capture) => ({
    id,
    slideKeys: capture.slideKeys,
    fieldJson: capture.fieldJson,
    opaque: await opaqueRefsOf(capture.opaque),
    media: await mediaRefsOf(capture.media),
  }),
  /** Each distinct blob a deck references, with its bytes — the package and
   * the media deduplicated by hash. */
  blobs: (capture) => partBlobsOf(capture.opaque, capture.media),
  baseline: baselineOf,
  delta: deltaOf,
  // A deck counts slides, not words.
  stats: (next) => ({ slides: next.slideKeys.length }),
  embeddedHashes,
  skeleton: (base) => {
    const baseline = base as DeckBaselinePayload;
    return { slides: baseline.slides, fields: baseline.fields, opaque: baseline.opaque, media: baseline.media };
  },
  apply: (parent, delta) => applyDelta(parent, delta as DeckDeltaPayload),
  baselineOf: (skeleton) => ({ kind: "deck-baseline", ...skeleton }),
  tailOf: (id, skeleton) => ({
    id,
    slideKeys: skeleton.slides.map((slide) => JSON.stringify(slide)),
    fieldJson: fieldJsonOf(skeleton.fields),
    opaque: skeleton.opaque,
    media: skeleton.media,
  }),
  hydrate: async (docId, skeleton, read, live) => {
    const opaque = await hydrateOpaque(skeleton.opaque, live?.opaque, read);
    const media = await hydrateMedia(skeleton.media, live?.media, read);
    const model = { kind: "presentation", id: docId, ...skeleton.fields, slides: skeleton.slides } as PresentationModel;
    if (opaque.length > 0) {
      model.opaque = opaque;
    }
    if (media.length > 0) {
      model.media = media;
    }
    return model;
  },
  liveHashes: async (model) => new Set((await partBlobsOf(model.opaque, model.media)).keys()),
};

const recorder = createVersionRecorder(deckCodec);

/** Reconstructs a deck version's whole model — its blobs read back — or null
 * when its chain or a blob is unreadable. `live` (the open deck) lends any
 * part whose content is unchanged. */
export function reconstructDeckVersion(docId: string, versionId: string, live?: PresentationModel): Promise<PresentationModel | null> {
  return recorder.reconstruct(docId, versionId, live);
}

// --- store IO -------------------------------------------------------------------

/** The deck's index; an empty one when absent or unreadable. */
export function loadDeckIndex(docId: string): Promise<VersionsIndexRecord> {
  return recorder.loadIndex(docId);
}

// --- recording a version -----------------------------------------------------

/** Writes one deck version, then enforces the budget. `pre-restore`/
 * `post-restore` write even while paused (they bracket a restore the user
 * asked for); `cadence`/`commit` respect the latch. Before any write, a deck
 * whose blob set plus one baseline would exceed HISTORY_BUDGET_BYTES is
 * refused and the index latches `paused` with pausedReason "too-large" —
 * re-checked on each later attempt, since media can shrink. Never rejects. */
export function recordDeckVersion(options: RecordDeckVersionOptions): Promise<RecordDeckVersionResult> {
  return recorder.record(options);
}

// --- budget and pruning ------------------------------------------------------

/** Re-runs the budget without adding a version — after a manual delete or
 * "Remove name", so the paused notice clears as soon as it should. */
export function reconcileDeckBudget(docId: string, live: PresentationModel): Promise<VersionsIndexRecord> {
  return recorder.reconcile(docId, live);
}

// --- the panel's actions -------------------------------------------------------

/** Names (or, with a blank name, unnames) a version — index only. */
export function nameDeckVersion(docId: string, versionId: string, name: string): Promise<VersionsIndexRecord> {
  return recorder.name(docId, versionId, name);
}

/** "Remove name": the version becomes prunable again. */
export function removeDeckVersionName(docId: string, versionId: string, live: PresentationModel): Promise<VersionsIndexRecord> {
  return recorder.removeName(docId, versionId, live);
}

/** "Delete": any version, named or not. Its immediate child (if any) is
 * re-encoded as a baseline first so every other version still reconstructs. */
export function deleteDeckVersion(docId: string, versionId: string, live: PresentationModel): Promise<VersionsIndexRecord> {
  return recorder.deleteVersion(docId, versionId, live);
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
