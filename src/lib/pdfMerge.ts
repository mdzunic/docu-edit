// The page organiser's rules (SPEC-111) — the framework-free half of the
// assembly workspace's PDF merge: how a page card describes its page (number,
// source, size, rotation, and the real aspect its paper-shaped placeholder
// takes — no page is ever rasterised), the "What changed" grouping of
// PdfDropRecords, the source-contiguous runs "Save as separate PDFs" cuts the
// tray into and the file names each run takes, the per-source keys a merge
// reports under, and the output bar's estimate line. The PDF work itself is
// core/pdfDocument's assemblePdf; nothing here imports Vue.
import { normaliseRotation, PDF_DROP_KINDS, type PdfDropKind, type PdfDropRecord } from "./core/pdfDocument";
import { pageCountLabel } from "./assembly";

/** The three destinations of the page-level output (one flow). */
export type MergeDestination = "merge" | "extract" | "split";

export const MERGE_DESTINATIONS: readonly MergeDestination[] = ["merge", "extract", "split"];

export const MERGE_DESTINATION_LABELS: Record<MergeDestination, string> = {
  merge: "Merge to one PDF",
  extract: "Extract selected pages",
  split: "Save as separate PDFs",
};

/** The title a merged PDF starts from, in place of SPEC-109's "Combined
 * image", when the tray starts with PDF pages in it. */
export const DEFAULT_MERGE_TITLE = "Combined PDF";

/** The source key a blank page reports under. */
export const BLANK_PAGE_TITLE = "Blank page";

/** "What changed" headings, one per kind. */
export const DROP_KIND_LABELS: Record<PdfDropKind, string> = {
  acroForm: "Form fields",
  structureTree: "Structure tree (accessibility tags)",
  javascript: "JavaScript",
  openAction: "Open actions",
  embeddedFile: "Attached files",
  namedDestination: "Named destinations",
  outlineEntry: "Bookmarks",
  annotation: "Links and annotations",
  pageLabel: "Page labels",
  encryptedSource: "Password-protected files",
  unreadableSource: "Unreadable files",
};

export interface DropGroup {
  kind: PdfDropKind;
  label: string;
  records: PdfDropRecord[];
}

/** The records grouped by kind, in PDF_DROP_KINDS order, each group's
 * records in report order; kinds with no record are left out. */
export function groupDrops(records: readonly PdfDropRecord[]): DropGroup[] {
  return PDF_DROP_KINDS.map((kind) => ({
    kind,
    label: DROP_KIND_LABELS[kind],
    records: records.filter((record) => record.kind === kind),
  })).filter((group) => group.records.length > 0);
}

/** One row of the disclosure: where the record came from, then why. */
export function dropRowText(record: PdfDropRecord): string {
  const where = record.page === undefined ? `“${record.source}”` : `“${record.source}”, page ${record.page}`;
  return `${where}: ${record.reason}`;
}

/** A page as its card shows it: its own size and the rotation it will
 * have (its /Rotate plus the user's turns). */
export interface CardPage {
  widthPt: number;
  heightPt: number;
  rotate: 0 | 90 | 180 | 270;
  rotateBy: 0 | 90 | 180 | 270;
}

/** The rotation the output page will have. */
export function effectiveRotation(page: Pick<CardPage, "rotate" | "rotateBy">): 0 | 90 | 180 | 270 {
  return normaliseRotation(page.rotate + page.rotateBy);
}

/** The card's rotation line: "rotate 90". */
export function rotationLabel(page: Pick<CardPage, "rotate" | "rotateBy">): string {
  return `rotate ${effectiveRotation(page)}`;
}

/** The aspect (width / height) of the page as it will display — a quarter
 * turn swaps the sides — for the paper-shaped placeholder; 1 for a page with
 * no area. */
export function cardAspect(page: CardPage): number {
  const turned = effectiveRotation(page) % 180 === 90;
  const width = turned ? page.heightPt : page.widthPt;
  const height = turned ? page.widthPt : page.heightPt;
  return width > 0 && height > 0 ? width / height : 1;
}

/** `delta` (±90) turns added to `rotateBy`, modulo 360. */
export function turned(rotateBy: number, delta: number): 0 | 90 | 180 | 270 {
  return normaliseRotation(rotateBy + delta);
}

/** Consecutive entries sharing a source key, in order: the files "Save as
 * separate PDFs" writes. */
export function sourceRuns<T>(entries: readonly T[], keyOf: (entry: T) => string): T[][] {
  const runs: T[][] = [];
  let last: string | null = null;
  for (const entry of entries) {
    const key = keyOf(entry);
    if (runs.length === 0 || key !== last) {
      runs.push([]);
    }
    runs[runs.length - 1].push(entry);
    last = key;
  }
  return runs;
}

/** A title made safe as a file name inside a zip (no path separators or
 * control characters). */
function safeFileTitle(title: string): string {
  const cleaned = Array.from(title)
    .map((char) => (char === "/" || char === "\\" || char.charCodeAt(0) < 32 ? "-" : char))
    .join("")
    .trim();
  return cleaned || "Untitled";
}

/** One file name per run, numbered in tray order so the zip lists them in
 * the order they were planned: "01 contract.pdf", "02 signature.pdf". */
export function runFileNames(titles: readonly string[]): string[] {
  const digits = Math.max(2, String(titles.length).length);
  return titles.map((title, index) => `${String(index + 1).padStart(digits, "0")} ${safeFileTitle(title)}.pdf`);
}

/** Report keys for the distinct sources, in first-appearance order: each
 * one's title, numbered "(2)", "(3)"… where two share one, so a record never
 * names two files at once. */
export function uniqueSourceKeys(titles: readonly string[]): string[] {
  const taken = new Set<string>();
  return titles.map((title) => {
    const base = title.trim() || "Untitled";
    let key = base;
    for (let n = 2; taken.has(key); n += 1) {
      key = `${base} (${n})`;
    }
    taken.add(key);
    return key;
  });
}

/** The output bar's line for the page-level output. */
export function mergeEstimateLine(
  destination: MergeDestination,
  counts: { pages: number; selected: number; sources: number; runs: number },
): string {
  if (destination === "extract") {
    return `${pageCountLabel(counts.selected)} of ${counts.pages} selected`;
  }
  if (destination === "split") {
    return `${counts.runs} ${counts.runs === 1 ? "PDF" : "PDFs"}, one per run of pages from the same file · downloads as a .zip`;
  }
  return `${pageCountLabel(counts.pages)} from ${counts.sources} ${counts.sources === 1 ? "file" : "files"}`;
}

/** "Save as separate PDFs" downloads under the assembly's title. */
export function zipFileName(title: string): string {
  const trimmed = title.trim();
  return `${safeFileTitle(trimmed || DEFAULT_MERGE_TITLE)}.zip`;
}
