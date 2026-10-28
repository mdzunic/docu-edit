// The library's search index as a running service (SPEC-114): what the
// index worker holds — every persisted per-document record, the searchable
// merge of them, and the operations the library asks for. Product-neutral:
// a document arrives as an entry (id, stamp, title, facets) plus either its
// extracted text or, for a PDF, its bytes, which are read here through
// core/pdfText so the parse never runs on the main thread.
//
// The index is a cache with a stamp, never a source of truth. A record is
// trusted only while its `indexVersion` is INDEX_VERSION and its
// `lastModified` is the entry's; a record whose document is gone is deleted;
// a trashed document's record is kept but excluded from search. Records are
// written one per document as each is built, so closing the tab mid-build
// loses only what was not built yet. The library's records together stay
// under MAX_INDEX_BYTES: past it — or whenever storage refuses a write —
// the oldest-used records are evicted first, each leaving a small marker so
// the "Not searched" list can say why that document is missing.
//
// Storage is injected (IndexStore): the worker hands in IndexedDB, tests a
// map. Every storage failure is absorbed — losing the index costs time and
// nothing else.
import {
  findTextFor,
  indexDocument,
  INDEX_VERSION,
  isSearchQueryError,
  MAX_SNIPPETS_PER_DOC,
  parseQuery,
  rank,
  snippets,
  titleSegments,
  withPositions,
  type DocPostings,
  type QueryError,
  type SearchIndex,
  type SnippetSegment,
} from "./searchIndex";
import { extractPdfText } from "./pdfText";
import { capExtractedText, COVERAGE_REASONS, MAX_INDEXED_CHARS, type ExtractedText, type TextCoverage } from "./extractedText";
import { base64ToBytes } from "./packageParts";

/** The most the library's index records may occupy together. */
export const MAX_INDEX_BYTES = 64 * 1024 * 1024;

/** One document's persisted index record: its stamp, its coverage and its
 * own posting lists (`tokens`), plus the text snippets are cut from. An
 * evicted marker keeps the stamp and the reason, and no tokens. */
export interface IndexRecord {
  docId: string;
  indexVersion: number;
  lastModified: number;
  /** Eviction order: the document's last modification, or later the last
   * time it was opened from a search result. */
  usedAt: number;
  coverage: TextCoverage;
  reason?: string;
  title: string;
  text: string;
  tokens: DocPostings | null;
  /** The record's approximate stored size. */
  bytes: number;
  evicted?: true;
}

/** Where records persist. Every method may reject (no storage, quota). */
export interface IndexStore {
  loadAll(): Promise<IndexRecord[]>;
  put(record: IndexRecord): Promise<void>;
  remove(docId: string): Promise<void>;
  clear(): Promise<void>;
}

/** A library entry as the index sees it. */
export interface IndexEntry {
  docId: string;
  lastModified: number;
  title: string;
  facets: Record<string, string[]>;
  trashed: boolean;
}

/** Whether a document is searched, and why not. */
export interface IndexStatus {
  docId: string;
  coverage: TextCoverage;
  reason?: string;
}

export interface SyncResult {
  /** Entries with no usable record, newest first: build these. */
  pending: string[];
  /** Entries whose record is current (searched, partly or not at all). */
  statuses: IndexStatus[];
}

/** One document to (re)build: its text as extracted, or a PDF's bytes. */
export interface BuildRequest {
  entry: IndexEntry;
  extracted?: ExtractedText;
  pdfBase64?: string;
}

export interface BuildResult extends IndexStatus {
  /** Documents evicted to make room, now not searched. */
  evicted: IndexStatus[];
}

export interface SearchRequest {
  input: string;
  limit: number;
  /** When given, only these documents may be hits (the library's filter
   * chips); every document still counts towards the ranking statistics. */
  ids?: string[];
}

export interface SearchResultHit {
  id: string;
  score: number;
  /** Occurrences matched, body and title together. */
  matches: number;
  /** The first body offset matched — where "open at match" lands. */
  at?: number;
  /** What the document's own find surface should look for. */
  findText: string;
  title: SnippetSegment[];
  snippets: SnippetSegment[][];
}

export interface SearchResult {
  error?: QueryError;
  /** Every hit, however many are returned. */
  total: number;
  hits: SearchResultHit[];
  /** How long the search itself took, in ms. */
  elapsedMs: number;
}

/** The reason an evicted document is not searched. */
export const INDEX_FULL_REASON = `${COVERAGE_REASONS.indexFull} — the least recently used documents are dropped to stay under ${MAX_INDEX_BYTES / (1024 * 1024)} MB`;

/** A record's approximate stored size: its JSON length. */
export function recordBytes(record: Omit<IndexRecord, "bytes">): number {
  return JSON.stringify(record).length;
}

function now(): number {
  return typeof performance !== "undefined" ? performance.now() : Date.now();
}

export class SearchEngine {
  private records = new Map<string, IndexRecord>();
  private entries = new Map<string, IndexEntry>();
  private loaded: Promise<void> | null = null;
  private view: SearchIndex | null = null;

  constructor(
    private readonly store: IndexStore,
    private readonly options: { maxBytes?: number; clock?: () => number } = {},
  ) {}

  private get maxBytes(): number {
    return this.options.maxBytes ?? MAX_INDEX_BYTES;
  }

  private load(): Promise<void> {
    this.loaded ??= this.store.loadAll().then(
      (records) => {
        for (const record of records) {
          this.records.set(record.docId, record);
        }
      },
      () => undefined,
    );
    return this.loaded;
  }

  private fresh(record: IndexRecord, entry: IndexEntry): boolean {
    return record.indexVersion === INDEX_VERSION && record.lastModified === entry.lastModified;
  }

  /** Reconciles the stored records with the library: deletes the records of
   * documents that are gone, and reports which entries need building. */
  async sync(entries: IndexEntry[]): Promise<SyncResult> {
    await this.load();
    this.entries = new Map(entries.map((entry) => [entry.docId, entry]));
    // Deleting the entry being visited is safe in a Map's iteration.
    for (const docId of this.records.keys()) {
      if (!this.entries.has(docId)) {
        this.records.delete(docId);
        this.store.remove(docId).catch(() => undefined);
      }
    }
    this.view = null;
    const pending: IndexEntry[] = [];
    const statuses: IndexStatus[] = [];
    for (const entry of entries) {
      // A trashed document is never built; a record it already has is kept
      // (a restore costs nothing) but stays out of every search.
      if (entry.trashed) {
        continue;
      }
      const record = this.records.get(entry.docId);
      if (record && this.fresh(record, entry)) {
        statuses.push(statusOf(record));
      } else {
        pending.push(entry);
      }
    }
    pending.sort((a, b) => b.lastModified - a.lastModified || (a.docId < b.docId ? -1 : 1));
    return { pending: pending.map((entry) => entry.docId), statuses };
  }

  /** Builds and stores one document's record. */
  async build(request: BuildRequest): Promise<BuildResult> {
    await this.load();
    const { entry } = request;
    let extracted: ExtractedText = request.extracted ?? { text: "", coverage: "none", reason: COVERAGE_REASONS.unreadable };
    if (request.pdfBase64 !== undefined) {
      extracted = await extractPdfText(base64ToBytes(request.pdfBase64));
    }
    extracted = capExtractedText(extracted, MAX_INDEXED_CHARS);
    const text = extracted.coverage === "none" ? "" : extracted.text;
    const tokens = indexDocument({ id: entry.docId, lastModified: entry.lastModified, title: entry.title, body: text, facets: entry.facets });
    const draft: Omit<IndexRecord, "bytes"> = {
      docId: entry.docId,
      indexVersion: INDEX_VERSION,
      lastModified: entry.lastModified,
      usedAt: Math.max(entry.lastModified, this.records.get(entry.docId)?.usedAt ?? 0),
      coverage: extracted.coverage,
      ...(extracted.reason ? { reason: extracted.reason } : {}),
      title: entry.title,
      text,
      tokens,
    };
    const record: IndexRecord = { ...draft, bytes: recordBytes(draft) };
    this.records.set(entry.docId, record);
    this.entries.set(entry.docId, entry);
    this.view = null;
    const evicted = this.enforceBound(entry.docId);
    await this.persist(record, evicted);
    const current = this.records.get(entry.docId) ?? record;
    return { ...statusOf(current), evicted };
  }

  /** Evicts the oldest-used records while the total is over the bound. */
  private enforceBound(justBuilt: string): IndexStatus[] {
    const evicted: IndexStatus[] = [];
    let total = 0;
    for (const record of this.records.values()) {
      total += record.bytes;
    }
    while (total > this.maxBytes) {
      const victim = this.oldestUsed(justBuilt) ?? (this.records.get(justBuilt)?.tokens ? this.records.get(justBuilt) : undefined);
      if (!victim) {
        break;
      }
      total -= victim.bytes;
      const marker = this.evict(victim);
      total += marker.bytes;
      evicted.push(statusOf(marker));
      if (victim.docId === justBuilt) {
        break;
      }
    }
    return evicted;
  }

  /** The searchable record used longest ago, other than `except`. */
  private oldestUsed(except: string): IndexRecord | undefined {
    let oldest: IndexRecord | undefined;
    for (const record of this.records.values()) {
      if (record.docId === except || !record.tokens) {
        continue;
      }
      if (!oldest || record.usedAt < oldest.usedAt || (record.usedAt === oldest.usedAt && record.docId < oldest.docId)) {
        oldest = record;
      }
    }
    return oldest;
  }

  /** Replaces a record with its evicted marker (kept, so the document is
   * not rebuilt in a loop and says why it is not searched). */
  private evict(record: IndexRecord): IndexRecord {
    const draft: Omit<IndexRecord, "bytes"> = {
      docId: record.docId,
      indexVersion: record.indexVersion,
      lastModified: record.lastModified,
      usedAt: record.usedAt,
      coverage: "none",
      reason: INDEX_FULL_REASON,
      title: record.title,
      text: "",
      tokens: null,
      evicted: true,
    };
    const marker: IndexRecord = { ...draft, bytes: recordBytes(draft) };
    this.records.set(record.docId, marker);
    this.view = null;
    this.store.put(marker).catch(() => undefined);
    return marker;
  }

  /** Writes a record; under quota pressure (a refused write) evicts the
   * oldest-used records one at a time and retries. A record that still
   * cannot be written stays searchable for this session only. */
  private async persist(record: IndexRecord, evicted: IndexStatus[]): Promise<void> {
    for (;;) {
      try {
        await this.store.put(record);
        return;
      } catch {
        const victim = this.oldestUsed(record.docId);
        if (!victim) {
          return;
        }
        evicted.push(statusOf(this.evict(victim)));
      }
    }
  }

  /** The document was opened from a search result: it was used now. */
  async touch(docId: string): Promise<void> {
    const record = this.records.get(docId);
    if (!record || !record.tokens) {
      return;
    }
    record.usedAt = (this.options.clock ?? Date.now)();
    await this.store.put(record).catch(() => undefined);
  }

  /** Deletes every stored record and forgets them all. */
  async clear(): Promise<void> {
    await this.load();
    this.records.clear();
    this.view = null;
    await this.store.clear().catch(() => undefined);
  }

  /** Storage is under pressure: drops every stored record (the index stays
   * in memory, and searchable, for this session). */
  async dropStored(): Promise<void> {
    await this.store.clear().catch(() => undefined);
  }

  /** The searchable merge: current records of documents in the library and
   * not in the trash. */
  private searchable(): SearchIndex {
    if (!this.view) {
      const docs: DocPostings[] = [];
      for (const record of this.records.values()) {
        const entry = this.entries.get(record.docId);
        if (record.tokens && entry && !entry.trashed && this.fresh(record, entry)) {
          docs.push(record.tokens);
        }
      }
      this.view = { version: INDEX_VERSION, docs };
    }
    return this.view;
  }

  search(request: SearchRequest): SearchResult {
    const start = now();
    const parsed = parseQuery(request.input);
    if (isSearchQueryError(parsed)) {
      return { error: parsed, total: 0, hits: [], elapsedMs: now() - start };
    }
    const allowed = request.ids ? new Set(request.ids) : null;
    const index = this.searchable();
    const all = rank(index, parsed, allowed ? (id) => allowed.has(id) : undefined);
    const shown = withPositions(index, parsed, all.slice(0, Math.max(0, request.limit)));
    const hits = shown.map((hit) => {
      const record = this.records.get(hit.id) as IndexRecord;
      return {
        id: hit.id,
        score: hit.score,
        matches: hit.bodyMatches + hit.titleMatches,
        ...(hit.positions.length > 0 ? { at: hit.positions[0] } : {}),
        findText: findTextFor(parsed, hit, record.text),
        title: titleSegments(record.title, parsed),
        snippets: snippets(record.text, hit, { max: MAX_SNIPPETS_PER_DOC }),
      };
    });
    return { total: all.length, hits, elapsedMs: now() - start };
  }
}

function statusOf(record: IndexRecord): IndexStatus {
  return { docId: record.docId, coverage: record.coverage, ...(record.reason ? { reason: record.reason } : {}) };
}
