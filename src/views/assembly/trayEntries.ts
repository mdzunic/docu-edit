// What one row of the assembly tray is (SPEC-109, SPEC-111): a library image
// to lay out, or — SPEC-111 — one page of a library PDF (with the quarter
// turns the user adds to it) or a blank page, each with a key of its own so a
// duplicate is its own entry. Built from the library's records and stored
// summaries only: no payload is read, and no source is ever touched. The
// controllers (useAssemblyTray, useTrayPages) and the persisted plan build
// their entries here.
import { getMedia, getStoredRecord } from "../../lib/documentStore";
import { isImageFormat } from "../../lib/libraryMenu";
import { DEFAULT_OVERLAY_TRANSFORM, MAX_ASSEMBLY_ITEMS, normaliseTransform, type OverlayTransform } from "../../lib/core/assemblyLayout";
import { MAX_MERGE_PAGES } from "../../lib/core/pdfDocument";
import { BLANK_PAGE_TITLE, rotationLabel } from "../../lib/pdfMerge";
import type { PdfSummary } from "../../lib/media/model";
import { sourceItemPx } from "../library/useConvertToPdf";

export const TRAY_FULL_MESSAGE = `A tray holds at most ${MAX_ASSEMBLY_ITEMS} images.`;
export const PAGES_FULL_MESSAGE = `A PDF plan holds at most ${MAX_MERGE_PAGES} pages.`;

/** Whether a library entry of `format` can go in the tray: an image, or
 * (SPEC-111) a PDF, as its pages. */
export function isCombinableFormat(format: string | undefined): boolean {
  return isImageFormat(format) || format === "pdf";
}

/** A library image as the tray needs it. */
export interface TrayImage {
  id: string;
  title: string;
  format: string;
  lastModified: number;
  /** 0 when the image's size could not be read. */
  widthPx: number;
  heightPx: number;
}

/** SPEC-111: what a page entry knows of its page, from the PDF's summary. */
export interface TrayPage {
  /** 0-based index into the PDF's pages. */
  index: number;
  /** How many pages the PDF has. */
  count: number;
  widthPt: number;
  heightPt: number;
  /** The page's own /Rotate. */
  rotate: 0 | 90 | 180 | 270;
}

export interface TrayEntry extends TrayImage {
  /** Unique within the tray: a duplicate shares its source's `id`. */
  key: string;
  /** overlay only. */
  transform: OverlayTransform;
  /** SPEC-111: a page of the library PDF `id`. */
  page?: TrayPage;
  /** SPEC-111: an empty page at the sheet's size (`id` is ""). */
  blank?: boolean;
  /** SPEC-111: quarter turns added to a page, mod 360 (0 when absent). */
  rotateBy?: 0 | 90 | 180 | 270;
}

/** Whether an entry is a page of a PDF plan (a PDF page or a blank page)
 * rather than an image to lay out. */
export function isPageEntry(entry: Pick<TrayEntry, "page" | "blank">): boolean {
  return entry.page !== undefined || entry.blank === true;
}

/** How an announcement names an entry: “shot”, page 2 of “contract”, the
 * blank page. */
export function entryLabel(entry: Pick<TrayEntry, "title" | "page" | "blank">): string {
  if (entry.blank) {
    return "the blank page";
  }
  return entry.page ? `page ${entry.page.index + 1} of “${entry.title}”` : `“${entry.title}”`;
}

/** entryLabel at the start of a sentence. */
export function entrySentenceLabel(entry: Pick<TrayEntry, "title" | "page" | "blank">): string {
  const label = entryLabel(entry);
  return label.charAt(0).toUpperCase() + label.slice(1);
}

/** A page card's accessible name: "Page 3: page 2 of “contract”, rotate 90,
 * selected". */
export function pageCardLabel(entry: TrayEntry, position: number, selected: boolean): string {
  const rotation = rotationLabel({ rotate: entry.page?.rotate ?? 0, rotateBy: entry.rotateBy ?? 0 });
  return `Page ${position}: ${entryLabel(entry)}, ${rotation}${selected ? ", selected" : ""}`;
}

/** The library image stored under `id` — null when it is missing, trashed
 * or not an image. */
export function libraryImage(id: string): TrayImage | null {
  const record = getStoredRecord(id);
  const model = getMedia(id);
  if (!record || !model || record.trashedAt !== undefined || !isImageFormat(model.format)) {
    return null;
  }
  let size = null;
  try {
    size = sourceItemPx(model);
  } catch {
    size = null;
  }
  return {
    id,
    title: model.title,
    format: model.format,
    lastModified: record.lastModified,
    widthPx: size?.widthPx ?? 0,
    heightPx: size?.heightPx ?? 0,
  };
}

/** A library PDF as the tray expands it: the source, and its pages from the
 * stored summary (no payload is read). */
export interface TrayPdf {
  source: TrayImage;
  pages: PdfSummary["pages"];
}

/** The library PDF stored under `id` — null when it is missing, trashed or
 * has no summary. */
export function libraryPdf(id: string): TrayPdf | null {
  const record = getStoredRecord(id);
  const model = getMedia(id);
  if (!record || !model || record.trashedAt !== undefined || model.format !== "pdf" || !model.pdf) {
    return null;
  }
  return {
    source: { id, title: model.title, format: "pdf", lastModified: record.lastModified, widthPx: 0, heightPx: 0 },
    pages: model.pdf.pages,
  };
}

let keySeed = 0;

/** A key no other entry has. */
export function freshKey(): string {
  keySeed += 1;
  return `tray-${keySeed}`;
}

/** A tray entry over `image`, with its own key. */
export function trayEntry(image: TrayImage, transform?: Partial<OverlayTransform>): TrayEntry {
  return { ...image, key: freshKey(), transform: normaliseTransform(transform ?? DEFAULT_OVERLAY_TRANSFORM) };
}

/** Page entries for `pdf`: every page in order by default, or the pages
 * `items` name (each with its turns); an index the PDF lacks is skipped. */
export function pageEntries(
  pdf: TrayPdf,
  items: ReadonlyArray<{ index: number; rotateBy?: 0 | 90 | 180 | 270 }> = pdf.pages.map((_, index) => ({ index })),
): TrayEntry[] {
  return items
    .filter((item) => Number.isInteger(item.index) && item.index >= 0 && item.index < pdf.pages.length)
    .map((item) => {
      const page = pdf.pages[item.index];
      return {
        ...trayEntry(pdf.source),
        page: { index: item.index, count: pdf.pages.length, widthPt: page.widthPt, heightPt: page.heightPt, rotate: page.rotate },
        ...(item.rotateBy ? { rotateBy: item.rotateBy } : {}),
      };
    });
}

/** A blank page entry. */
export function blankEntry(rotateBy: 0 | 90 | 180 | 270 = 0): TrayEntry {
  return {
    ...trayEntry({ id: "", title: BLANK_PAGE_TITLE, format: "pdf", lastModified: 0, widthPx: 0, heightPx: 0 }),
    blank: true,
    ...(rotateBy ? { rotateBy } : {}),
  };
}

/** Typed pass-through for entries read off a `.vue` prop (runtime-typed
 * `Object`; espree bars the annotation in `.vue` scripts). */
export function asTrayEntries(value: unknown): TrayEntry[] {
  return Array.isArray(value) ? (value as TrayEntry[]) : [];
}
