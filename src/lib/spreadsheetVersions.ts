// Workbook version history (SPEC-138 "Workbook version payloads"): SPEC-040's
// chain — baselines, deltas against the immediately preceding version,
// content-addressed blobs, the 50 MB budget and its prune loop — recorded for
// workbooks through versionRecorder.ts's loop, this module being the second
// VersionCodec beside the deck's. It lives at the top of src/lib, beside
// documentStore.ts and presentationVersions.ts, because lib/spreadsheet may
// not import ./idb. Records share the EXISTING `versions` object store and
// its key prefixes (no DB_VERSION bump); a workbook's payloads are told apart
// by their `workbook-` kinds.
//
// A baseline carries every sheet whole — its meta (every Sheet field but
// `rows` and `sourceXml`: `dirty` and the rest included, they decide what the
// writers rebuild), its rows verbatim and its `sourceXml` as a blob ref — and
// every workbook-level field but the sheets, the two part arrays, the preview
// and the ods envelope. A delta is built over row IDENTITY, which SPEC-116's
// copy-on-write invariant makes sound: a sheet whose `rows` array, every meta
// value and source reference are the predecessor's contributes nothing; a
// changed one carries its meta whole when a meta value was replaced, its
// source ref when the source was, and its rows as windows — the rows cut into
// bands of VERSION_ROW_BAND row INDICES, windowOf trimming each changed
// band's reference-equal head and tail, so an edit at row 5 and one at row
// 900,000 cost two one-row windows. A structural edit that re-indexes every
// row windows every band and lands, correctly, in SPEC-040's wholesale-rewrite
// rule as a baseline — as does the first delta after a reload, when every row
// read back from storage is new.
//
// `opaque` and `media` parts, each sheet's `sourceXml` and the ods envelope
// are blobs (`blob:<docId>:<sha256>`), each hashed at most once per reference
// (SPEC-124/125's retained-reference rewrite replaces a source, never
// mutates it), and a delta carries a blob ref only when it changed — a 30 MB
// .ods table's source costs one hash and one blob per distinct content.
//
// Over budget a workbook degrades to no history: more than
// WORKBOOK_HISTORY_MAX_CELLS stored cells is refused before any encode (a
// count, no JSON), a smaller one whose blobs plus one baseline exceed the
// budget before any write; either latches the index `paused` with
// pausedReason "too-large" (WORKBOOK_HISTORY_TOO_LARGE_MESSAGE).
import { windowOf } from "./core/arrayWindow";
import type { MediaPart } from "./core/media";
import type { OpaquePart } from "./core/packageParts";
import { measureBytes, sha256Hex, type SnapshotTrigger, type VersionBlobRecord, type VersionIndexEntry, type VersionsIndexRecord } from "./core/versionChain";
import type { Row, Sheet, SpreadsheetModel } from "./spreadsheet/model";
import {
  createVersionRecorder,
  hydrateMedia,
  hydrateOpaque,
  mediaRefsOf,
  opaqueRefsOf,
  partBlobsOf,
  sameRefs,
  type BlobSet,
  type MediaRef,
  type OpaqueRef,
  type VersionCodec,
  type VersionPayload,
} from "./versionRecorder";

/** Rows per delta band: a band's window is trimmed on its own, so an edit's
 * cost is the rows it touched, never the distance between two edits. */
export const VERSION_ROW_BAND = 4096;

/** Stored cells above which a workbook keeps no history: its baseline alone
 * would be a fifth of the budget, and encoding it a visible pause. */
export const WORKBOOK_HISTORY_MAX_CELLS = 250_000;

/** The history panel's paused notice for either over-budget refusal. Normal
 * exhaustion (nothing prunable and still over) keeps core's
 * HISTORY_PAUSED_MESSAGE. */
export const WORKBOOK_HISTORY_TOO_LARGE_MESSAGE =
  "Version history off — this workbook is larger than the 50 MB history budget";

// --- payload shapes ----------------------------------------------------------

/** A blob by content: its text lives in `blob:<docId>:<hash>`. */
export interface BlobRef {
  hash: string;
}

/** A sheet without its rows and its source text: name, visibility, columns,
 * merges, the rest of its settings, and `dirty`. */
export type SheetMeta = Omit<Sheet, "rows" | "sourceXml">;

export interface SheetSnapshot {
  id: string;
  meta: SheetMeta;
  rows: Row[];
  sourceXml?: BlobRef;
}

/** Every workbook-level field but the ones a payload carries elsewhere or
 * not at all — title, format, packageVariant, readerVersion,
 * activeSheetIndex, dateSystem, styles, theme, properties, names,
 * structureDirty and whatever later specs add. In a delta, present only when
 * replaced; `null` = changed to absent. */
export type WorkbookFieldKey = Exclude<keyof SpreadsheetModel, "kind" | "id" | "sheets" | "opaque" | "media" | "preview" | "odsContentEnvelope">;
export type WorkbookFields = { [K in WorkbookFieldKey]?: SpreadsheetModel[K] | null };

export interface WorkbookBaselinePayload {
  kind: "workbook-baseline";
  sheets: SheetSnapshot[];
  fields: WorkbookFields;
  opaque: OpaqueRef[];
  media: MediaRef[];
  envelope?: BlobRef;
}

/** One changed band of a sheet's rows: the parent's band with `prefixLen`
 * rows kept from the front, `suffixLen` from the back and `replaced`
 * between them. */
export interface RowWindow {
  band: number;
  prefixLen: number;
  suffixLen: number;
  replaced: Row[];
}

export type SheetDelta =
  | { id: string; kind: "added"; snapshot: SheetSnapshot }
  /** `sourceXml: null` — the sheet no longer has a source. */
  | { id: string; kind: "changed"; meta?: SheetMeta; rows?: RowWindow[]; sourceXml?: BlobRef | null };

export interface WorkbookDeltaPayload {
  kind: "workbook-delta";
  parentId: string;
  /** Every sheet id, in order; a parent's sheet absent here was deleted. */
  order: string[];
  sheets: SheetDelta[];
  fields: WorkbookFields;
  /** Present only when changed (`envelope: null` — removed). */
  opaque?: OpaqueRef[];
  media?: MediaRef[];
  envelope?: BlobRef | null;
}

export type WorkbookPayload = WorkbookBaselinePayload | WorkbookDeltaPayload;

export interface WorkbookVersionPayloadRecord {
  id: string;
  store: "payload";
  docId: string;
  versionId: string;
  payload: WorkbookPayload;
}

export type WorkbookVersionStoreRecord = VersionsIndexRecord | WorkbookVersionPayloadRecord | VersionBlobRecord;

// --- the capture ---------------------------------------------------------------

const EXCLUDED_FIELDS = new Set(["kind", "id", "sheets", "opaque", "media", "preview", "odsContentEnvelope"]);

interface CapturedSheet {
  id: string;
  meta: SheetMeta;
  rows: Row[];
  sourceXml?: string;
}

/** A workbook at one instant, taken synchronously before a snapshot awaits
 * anything: every reference it holds is copy-on-write (rows, the sheets'
 * values, the workbook's fields, the part arrays), so holding them is the
 * snapshot — nothing is serialised until the payload is written. */
interface WorkbookCapture {
  docId: string;
  sheets: CapturedSheet[];
  fields: Record<string, unknown>;
  opaque: OpaquePart[];
  media?: MediaPart[];
  envelope?: { prefix: string; suffix: string };
  cells: number;
}

/** What the next snapshot's delta is built against: the capture with its
 * blobs hashed, kept in memory as `cachedPrev` (after a reload, rebuilt from
 * the chain — whose rows are all new objects, so the delta trips rule 4). */
export interface WorkbookVersionTail {
  id: string;
  sheets: (CapturedSheet & { source?: BlobRef })[];
  fields: Record<string, unknown>;
  opaque: OpaqueRef[];
  media: MediaRef[];
  envelope?: BlobRef;
  cells: number;
}

// Rows arrays are copy-on-write, so a count taken once holds for the array.
const cellCounts = new WeakMap<readonly Row[], number>();

/** A sheet's stored cells — the sum of `row.cells.length`, no encode. */
export function storedCells(rows: readonly Row[]): number {
  let count = cellCounts.get(rows);
  if (count === undefined) {
    count = 0;
    for (const row of rows) {
      count += row.cells.length;
    }
    cellCounts.set(rows, count);
  }
  return count;
}

/** A workbook's stored cells, every sheet's summed. */
export function workbookCells(model: Pick<SpreadsheetModel, "sheets">): number {
  return model.sheets.reduce((total, sheet) => total + storedCells(sheet.rows), 0);
}

function metaOf(sheet: Sheet): SheetMeta {
  const { rows: _rows, sourceXml: _source, ...meta } = sheet;
  return meta;
}

function captureWorkbook(model: SpreadsheetModel): WorkbookCapture {
  const fields: Record<string, unknown> = {};
  for (const [key, value] of Object.entries(model)) {
    if (!EXCLUDED_FIELDS.has(key) && value !== undefined) {
      fields[key] = value;
    }
  }
  const sheets = model.sheets.map((sheet): CapturedSheet => {
    const captured: CapturedSheet = { id: sheet.id, meta: metaOf(sheet), rows: sheet.rows };
    if (sheet.sourceXml !== undefined) {
      captured.sourceXml = sheet.sourceXml;
    }
    return captured;
  });
  const capture: WorkbookCapture = { docId: model.id, sheets, fields, opaque: model.opaque, cells: workbookCells(model) };
  if (model.media !== undefined) {
    capture.media = model.media;
  }
  if (model.odsContentEnvelope !== undefined) {
    capture.envelope = model.odsContentEnvelope;
  }
  return capture;
}

// --- text blobs: sources and the envelope ---------------------------------------

/** A text blob's digest and measured size, kept beside the exact string it
 * was taken over, per `docId:sheetId` (and `docId:` for the envelope): the
 * string is compared by identity first, so an unchanged source is hashed
 * once per reference and never re-encoded to be measured. */
const textCache = new Map<string, { text: string; hash: string; size: number }>();

async function textBlob(key: string, text: string): Promise<{ hash: string; size: number }> {
  const cached = textCache.get(key);
  if (cached && cached.text === text) {
    return cached;
  }
  const entry = { text, hash: await sha256Hex(text), size: measureBytes(text) };
  textCache.set(key, entry);
  return entry;
}

const envelopeTexts = new WeakMap<object, string>();

function envelopeText(envelope: { prefix: string; suffix: string }): string {
  let text = envelopeTexts.get(envelope);
  if (text === undefined) {
    text = JSON.stringify({ prefix: envelope.prefix, suffix: envelope.suffix });
    envelopeTexts.set(envelope, text);
  }
  return text;
}

function sourceKey(docId: string, sheetId: string): string {
  return `${docId}:${sheetId}`;
}

function envelopeKey(docId: string): string {
  return `${docId}:`;
}

async function tailOf(id: string, capture: WorkbookCapture): Promise<WorkbookVersionTail> {
  const sheets: WorkbookVersionTail["sheets"] = [];
  for (const sheet of capture.sheets) {
    if (sheet.sourceXml === undefined) {
      sheets.push(sheet);
    } else {
      sheets.push({ ...sheet, source: { hash: (await textBlob(sourceKey(capture.docId, sheet.id), sheet.sourceXml)).hash } });
    }
  }
  const tail: WorkbookVersionTail = {
    id,
    sheets,
    fields: capture.fields,
    opaque: await opaqueRefsOf(capture.opaque),
    media: await mediaRefsOf(capture.media),
    cells: capture.cells,
  };
  if (capture.envelope) {
    tail.envelope = { hash: (await textBlob(envelopeKey(capture.docId), envelopeText(capture.envelope))).hash };
  }
  return tail;
}

async function blobsOf(capture: WorkbookCapture): Promise<BlobSet> {
  const blobs = await partBlobsOf(capture.opaque, capture.media);
  for (const sheet of capture.sheets) {
    if (sheet.sourceXml !== undefined) {
      const { hash, size } = await textBlob(sourceKey(capture.docId, sheet.id), sheet.sourceXml);
      blobs.set(hash, { data: sheet.sourceXml, size });
    }
  }
  if (capture.envelope) {
    const text = envelopeText(capture.envelope);
    const { hash, size } = await textBlob(envelopeKey(capture.docId), text);
    blobs.set(hash, { data: text, size });
  }
  return blobs;
}

// --- row bands ---------------------------------------------------------------------

/** Each band's rows in `rows` (sorted ascending by index), keyed by band. */
function bandsOf(rows: readonly Row[]): Map<number, Row[]> {
  const bands = new Map<number, Row[]>();
  let start = 0;
  while (start < rows.length) {
    const band = Math.floor(rows[start]!.index / VERSION_ROW_BAND);
    let end = start + 1;
    while (end < rows.length && Math.floor(rows[end]!.index / VERSION_ROW_BAND) === band) {
      end += 1;
    }
    bands.set(band, rows.slice(start, end));
    start = end;
  }
  return bands;
}

function sameSlice(a: readonly Row[], b: readonly Row[]): boolean {
  return a.length === b.length && a.every((row, i) => row === b[i]);
}

/** The windows that turn `before` into `after`: one per band whose rows
 * differ by identity, trimmed by windowOf. */
export function rowWindows(before: Row[], after: Row[]): RowWindow[] {
  if (before === after) {
    return [];
  }
  const was = bandsOf(before);
  const now = bandsOf(after);
  const bands = [...new Set([...was.keys(), ...now.keys()])].toSorted((a, b) => a - b);
  const windows: RowWindow[] = [];
  for (const band of bands) {
    const beforeBand = was.get(band) ?? [];
    const afterBand = now.get(band) ?? [];
    if (sameSlice(beforeBand, afterBand)) {
      continue;
    }
    const window = windowOf(beforeBand, afterBand);
    windows.push({ band, prefixLen: window.prefixLen, suffixLen: window.suffixLen, replaced: window.afterMid });
  }
  return windows;
}

/** `rows` with each window applied to its band, bands re-concatenated in
 * order. */
export function applyRowWindows(rows: readonly Row[], windows: readonly RowWindow[]): Row[] {
  if (windows.length === 0) {
    return [...rows];
  }
  const bands = bandsOf(rows);
  for (const { band, prefixLen, suffixLen, replaced } of windows) {
    const before = bands.get(band) ?? [];
    bands.set(band, [...before.slice(0, prefixLen), ...replaced, ...(suffixLen > 0 ? before.slice(before.length - suffixLen) : [])]);
  }
  return [...bands.keys()].toSorted((a, b) => a - b).flatMap((band) => bands.get(band)!);
}

// --- payload build ---------------------------------------------------------------

function snapshotOf(sheet: WorkbookVersionTail["sheets"][number]): SheetSnapshot {
  const snapshot: SheetSnapshot = { id: sheet.id, meta: sheet.meta, rows: sheet.rows };
  if (sheet.source) {
    snapshot.sourceXml = sheet.source;
  }
  return snapshot;
}

function baselineOf(next: WorkbookVersionTail): WorkbookBaselinePayload {
  const payload: WorkbookBaselinePayload = {
    kind: "workbook-baseline",
    sheets: next.sheets.map(snapshotOf),
    fields: next.fields as WorkbookFields,
    opaque: next.opaque,
    media: next.media,
  };
  if (next.envelope) {
    payload.envelope = next.envelope;
  }
  return payload;
}

/** Whether every meta value is the same reference (and the key sets match). */
function sameMeta(a: SheetMeta, b: SheetMeta): boolean {
  const ra = a as unknown as Record<string, unknown>;
  const rb = b as unknown as Record<string, unknown>;
  const keys = Object.keys(ra).filter((key) => ra[key] !== undefined);
  return keys.length === Object.keys(rb).filter((key) => rb[key] !== undefined).length && keys.every((key) => ra[key] === rb[key]);
}

function deltaOf(prev: WorkbookVersionTail, next: WorkbookVersionTail): WorkbookDeltaPayload {
  const before = new Map(prev.sheets.map((sheet) => [sheet.id, sheet]));
  const sheets: SheetDelta[] = [];
  for (const sheet of next.sheets) {
    const was = before.get(sheet.id);
    if (!was) {
      sheets.push({ id: sheet.id, kind: "added", snapshot: snapshotOf(sheet) });
      continue;
    }
    const change: Extract<SheetDelta, { kind: "changed" }> = { id: sheet.id, kind: "changed" };
    if (!sameMeta(was.meta, sheet.meta)) {
      change.meta = sheet.meta;
    }
    if (was.source?.hash !== sheet.source?.hash) {
      change.sourceXml = sheet.source ?? null;
    }
    const windows = rowWindows(was.rows, sheet.rows);
    if (windows.length > 0) {
      change.rows = windows;
    }
    if (change.meta !== undefined || change.sourceXml !== undefined || change.rows !== undefined) {
      sheets.push(change);
    }
  }
  const fields: Record<string, unknown> = {};
  for (const key of new Set([...Object.keys(prev.fields), ...Object.keys(next.fields)])) {
    if (prev.fields[key] !== next.fields[key]) {
      fields[key] = key in next.fields ? next.fields[key] : null;
    }
  }
  const delta: WorkbookDeltaPayload = {
    kind: "workbook-delta",
    parentId: prev.id,
    order: next.sheets.map((sheet) => sheet.id),
    sheets,
    fields: fields as WorkbookFields,
  };
  if (!sameRefs(prev.opaque, next.opaque)) {
    delta.opaque = next.opaque;
  }
  if (!sameRefs(prev.media, next.media)) {
    delta.media = next.media;
  }
  if (prev.envelope?.hash !== next.envelope?.hash) {
    delta.envelope = next.envelope ?? null;
  }
  return delta;
}

/** The blob hashes a payload itself embeds, once each. */
function embeddedHashes(payload: WorkbookPayload): string[] {
  const hashes: string[] = [...(payload.opaque ?? []), ...(payload.media ?? [])].map((ref) => ref.hash);
  if (payload.envelope) {
    hashes.push(payload.envelope.hash);
  }
  for (const sheet of payload.sheets) {
    const source = "snapshot" in sheet ? sheet.snapshot.sourceXml : sheet.sourceXml;
    if (source) {
      hashes.push(source.hash);
    }
  }
  return [...new Set(hashes)];
}

// --- reconstruction ----------------------------------------------------------

/** A version without its bytes: sheets with their rows, fields and blob refs. */
interface WorkbookSkeleton {
  sheets: SheetSnapshot[];
  fields: WorkbookFields;
  opaque: OpaqueRef[];
  media: MediaRef[];
  envelope?: BlobRef;
}

function applyDelta(parent: WorkbookSkeleton, delta: WorkbookDeltaPayload): WorkbookSkeleton {
  const byId = new Map(parent.sheets.map((sheet) => [sheet.id, sheet]));
  for (const change of delta.sheets) {
    if (change.kind === "added") {
      byId.set(change.id, change.snapshot);
      continue;
    }
    const was = byId.get(change.id);
    if (!was) {
      continue;
    }
    const sheet: SheetSnapshot = { id: was.id, meta: change.meta ?? was.meta, rows: change.rows ? applyRowWindows(was.rows, change.rows) : was.rows };
    const source = change.sourceXml === undefined ? was.sourceXml : (change.sourceXml ?? undefined);
    if (source) {
      sheet.sourceXml = source;
    }
    byId.set(change.id, sheet);
  }
  const fields: Record<string, unknown> = { ...parent.fields };
  for (const [key, value] of Object.entries(delta.fields)) {
    if (value === null) {
      delete fields[key];
    } else {
      fields[key] = value;
    }
  }
  const skeleton: WorkbookSkeleton = {
    sheets: delta.order.flatMap((id) => {
      const sheet = byId.get(id);
      return sheet ? [sheet] : [];
    }),
    fields: fields as WorkbookFields,
    opaque: delta.opaque ?? parent.opaque,
    media: delta.media ?? parent.media,
  };
  const envelope = delta.envelope === undefined ? parent.envelope : (delta.envelope ?? undefined);
  if (envelope) {
    skeleton.envelope = envelope;
  }
  return skeleton;
}

/** A text blob back: the live string when its cached digest is the ref's
 * (a restore whose source did not change keeps the live reference, so the
 * store's source latch skips rewriting it), else the blob's. */
async function hydrateText(key: string, ref: BlobRef, live: string | undefined, read: (hash: string) => Promise<string>): Promise<string> {
  if (live !== undefined && (await textBlob(key, live)).hash === ref.hash) {
    return live;
  }
  return read(ref.hash);
}

async function hydrate(docId: string, skeleton: WorkbookSkeleton, read: (hash: string) => Promise<string>, live?: SpreadsheetModel): Promise<SpreadsheetModel> {
  const liveSheets = new Map((live?.sheets ?? []).map((sheet) => [sheet.id, sheet]));
  const sheets: Sheet[] = [];
  for (const snapshot of skeleton.sheets) {
    const sheet = { ...snapshot.meta, id: snapshot.id, rows: snapshot.rows } as Sheet;
    if (snapshot.sourceXml) {
      sheet.sourceXml = await hydrateText(sourceKey(docId, snapshot.id), snapshot.sourceXml, liveSheets.get(snapshot.id)?.sourceXml, read);
    }
    sheets.push(sheet);
  }
  const model = { kind: "spreadsheet", id: docId, ...skeleton.fields, sheets } as SpreadsheetModel;
  model.opaque = await hydrateOpaque(skeleton.opaque, live?.opaque, read);
  const media = await hydrateMedia(skeleton.media, live?.media, read);
  if (media.length > 0) {
    model.media = media;
  }
  if (skeleton.envelope) {
    const envelope = live?.odsContentEnvelope;
    const text = await hydrateText(envelopeKey(docId), skeleton.envelope, envelope ? envelopeText(envelope) : undefined, read);
    model.odsContentEnvelope = envelope && text === envelopeText(envelope) ? envelope : (JSON.parse(text) as { prefix: string; suffix: string });
  }
  return model;
}

/** The workbook's codec for versionRecorder's loop. */
const workbookCodec: VersionCodec<SpreadsheetModel, WorkbookCapture, WorkbookVersionTail, WorkbookPayload, WorkbookSkeleton> = {
  owns: (payload: VersionPayload): payload is WorkbookPayload => payload.kind === "workbook-baseline" || payload.kind === "workbook-delta",
  isBaseline: (payload) => payload.kind === "workbook-baseline",
  capture: captureWorkbook,
  tooLargeToEncode: (capture) => capture.cells > WORKBOOK_HISTORY_MAX_CELLS,
  tail: tailOf,
  blobs: blobsOf,
  baseline: baselineOf,
  delta: deltaOf,
  stats: (next) => ({ sheets: next.sheets.length, cells: next.cells }),
  embeddedHashes,
  skeleton: (base) => {
    const baseline = base as WorkbookBaselinePayload;
    const skeleton: WorkbookSkeleton = { sheets: baseline.sheets, fields: baseline.fields, opaque: baseline.opaque, media: baseline.media };
    if (baseline.envelope) {
      skeleton.envelope = baseline.envelope;
    }
    return skeleton;
  },
  apply: (parent, delta) => applyDelta(parent, delta as WorkbookDeltaPayload),
  baselineOf: (skeleton) => ({ kind: "workbook-baseline", ...skeleton }),
  tailOf: (id, skeleton) => {
    const tail: WorkbookVersionTail = {
      id,
      sheets: skeleton.sheets.map((sheet) => (sheet.sourceXml ? { id: sheet.id, meta: sheet.meta, rows: sheet.rows, source: sheet.sourceXml } : { id: sheet.id, meta: sheet.meta, rows: sheet.rows })),
      fields: skeleton.fields as Record<string, unknown>,
      opaque: skeleton.opaque,
      media: skeleton.media,
      cells: skeleton.sheets.reduce((total, sheet) => total + storedCells(sheet.rows), 0),
    };
    if (skeleton.envelope) {
      tail.envelope = skeleton.envelope;
    }
    return tail;
  },
  hydrate,
  liveHashes: async (model) => new Set((await blobsOf(captureWorkbook(model))).keys()),
};

const recorder = createVersionRecorder(workbookCodec);

// --- the API -------------------------------------------------------------------

export interface RecordWorkbookVersionOptions {
  docId: string;
  model: SpreadsheetModel;
  author: string;
  trigger: SnapshotTrigger;
  /** The last version recorded this session; null on a fresh session. */
  cachedPrev: WorkbookVersionTail | null;
}

export interface RecordWorkbookVersionResult {
  written: boolean;
  index: VersionsIndexRecord;
  newEntry: WorkbookVersionTail | null;
}

/** Writes one workbook version, then enforces the budget (versionRecorder's
 * loop: `pre-restore`/`post-restore` write even while paused; the two
 * over-budget refusals latch "too-large" before any encode or write). The
 * model is captured synchronously — call it with the live workbook. Never
 * rejects. */
export function recordWorkbookVersion(options: RecordWorkbookVersionOptions): Promise<RecordWorkbookVersionResult> {
  return recorder.record(options);
}

/** Reconstructs a workbook version's whole model — its blobs read back — or
 * null when its chain or a blob is unreadable. `live` (the open workbook)
 * lends any part or source whose content is unchanged. The model has the
 * version's reader version: callers run upgradeSpreadsheetModel. */
export function reconstructWorkbookVersion(docId: string, versionId: string, live?: SpreadsheetModel): Promise<SpreadsheetModel | null> {
  return recorder.reconstruct(docId, versionId, live);
}

/** The workbook's index; an empty one when absent or unreadable. */
export function loadWorkbookIndex(docId: string): Promise<VersionsIndexRecord> {
  return recorder.loadIndex(docId);
}

/** Re-runs the budget without adding a version. */
export function reconcileWorkbookBudget(docId: string, live: SpreadsheetModel): Promise<VersionsIndexRecord> {
  return recorder.reconcile(docId, live);
}

/** Names (or, with a blank name, unnames) a version — index only. */
export function nameWorkbookVersion(docId: string, versionId: string, name: string): Promise<VersionsIndexRecord> {
  return recorder.name(docId, versionId, name);
}

/** "Remove name": the version becomes prunable again. */
export function removeWorkbookVersionName(docId: string, versionId: string, live: SpreadsheetModel): Promise<VersionsIndexRecord> {
  return recorder.removeName(docId, versionId, live);
}

/** "Delete": any version; its child is re-baselined first. */
export function deleteWorkbookVersion(docId: string, versionId: string, live: SpreadsheetModel): Promise<VersionsIndexRecord> {
  return recorder.deleteVersion(docId, versionId, live);
}

export interface WorkbookVersionListEntry extends VersionIndexEntry {
  /** Stored-cell change against the chronologically preceding version; the
   * oldest version's is its own cell count. */
  cellDelta: number;
}

/** Newest first, each with its cell-count delta — the index alone, no
 * payload read. */
export function workbookVersionEntries(index: VersionsIndexRecord): WorkbookVersionListEntry[] {
  return index.versions
    .map((entry, i) => ({
      ...entry,
      cellDelta: (entry.cells ?? 0) - (i === 0 ? 0 : (index.versions[i - 1]!.cells ?? 0)),
    }))
    .toReversed();
}
