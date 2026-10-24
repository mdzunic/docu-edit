// The assembly workspace's persisted plan (SPEC-109): what /assemble keeps
// under the one `assemblyPlan` key of the SPEC-066 preferences store so the
// workspace survives a reload — the tray's item ids (with each item's title,
// so a notice can name an item that has since been deleted), the layout, the
// per-item overlay transforms, the sheet setup and the output settings
// (SPEC-110's PDF page mode, n-up, page numbers and metadata among them).
// Never pixels: an item's size is read back from the library on return.
//
// SPEC-111 adds PDF pages to the same list: an item may be one page of a
// library PDF (its page index and the quarter turns added to it) or a blank
// page; a tray holds up to MAX_ASSEMBLY_ITEMS images and MAX_MERGE_PAGES
// entries in all.
//
// Kept apart from lib/assembly.ts and imported by lib/preferences.ts, so it
// imports core and the framework-free media layer only — no cycle back
// through the preferences store. The reader is tolerant in the store's own
// way: a bad field falls back to its default and one bad item is dropped,
// never the whole plan.
import {
  ASSEMBLY_PLAN_DEFAULTS,
  LAYOUT_KINDS,
  MAX_ASSEMBLY_ITEMS,
  MAX_GAP_CM,
  normaliseTransform,
  type AssemblyColumns,
  type LayoutKind,
  type OverlayTransform,
} from "./core/assemblyLayout";
import {
  isSheetBackground,
  MAX_SHEET_CM,
  SHEET_ALIGNS,
  SHEET_DPI_CHOICES,
  SHEET_FITS,
  SHEET_PAPER_NAMES,
  type SheetSetup,
} from "./core/sheetLayout";
import { DEFAULT_N_UP, DEFAULT_PAGE_MODE, N_UP_CHOICES, PAGE_MODES, type PageMode } from "./media/imagesToPdf";
import { MAX_MERGE_PAGES, normaliseRotation } from "./core/pdfDocument";

/** The raster outputs, and SPEC-110's PDF beside them. */
export type AssemblyFormat = "png" | "jpeg" | "pdf";

export const ASSEMBLY_FORMATS: readonly AssemblyFormat[] = ["png", "jpeg", "pdf"];

export const DEFAULT_JPEG_QUALITY = 90;

/** The title a saved assembly starts from (numbered past any taken one). */
export const DEFAULT_ASSEMBLY_TITLE = "Combined image";

/** The colour a JPEG paints where the sheet asked for transparency. */
export const DEFAULT_MATTE = "#ffffff";

export interface AssemblyOutputSettings {
  format: AssemblyFormat;
  /** JPEG only, 1–100. */
  quality: number;
  title: string;
  /** "#rrggbb": what JPEG output paints behind a transparent sheet. */
  matte: string;
  /** PDF only (SPEC-110): how the tray is cut into pages. */
  pageMode: PageMode;
  /** PDF n-up only: items per page, one of N_UP_CHOICES. */
  nUp: number;
  /** PDF only: a page number on every page. */
  pageNumbers: boolean;
  /** PDF only: /Author and /Subject — empty unless the user types them. */
  author: string;
  subject: string;
}

export const ASSEMBLY_OUTPUT_DEFAULTS: AssemblyOutputSettings = {
  format: "png",
  quality: DEFAULT_JPEG_QUALITY,
  title: DEFAULT_ASSEMBLY_TITLE,
  matte: DEFAULT_MATTE,
  pageMode: DEFAULT_PAGE_MODE,
  nUp: DEFAULT_N_UP,
  pageNumbers: false,
  author: "",
  subject: "",
};

/** The layout half of an AssemblyPlan; fit and align are the sheet's. */
export interface AssemblyLayoutSettings {
  layout: LayoutKind;
  columns: AssemblyColumns;
  gapCm: number;
  cellBackground: string;
}

export const ASSEMBLY_LAYOUT_DEFAULTS: AssemblyLayoutSettings = {
  layout: ASSEMBLY_PLAN_DEFAULTS.layout,
  columns: ASSEMBLY_PLAN_DEFAULTS.columns,
  gapCm: ASSEMBLY_PLAN_DEFAULTS.gapCm,
  cellBackground: ASSEMBLY_PLAN_DEFAULTS.cellBackground,
};

export interface PersistedAssemblyItem {
  /** The library id of the source media document ("" for a blank page). */
  id: string;
  /** Its title when it was added — named by the notice if it is gone. */
  title: string;
  transform?: OverlayTransform;
  /** SPEC-111: the 0-based page of the library PDF `id` this item is. */
  page?: number;
  /** SPEC-111: an empty page at the sheet's size. */
  blank?: boolean;
  /** SPEC-111: quarter turns added to a page, mod 360. */
  rotateBy?: 0 | 90 | 180 | 270;
}

/** Whether an item is a page of a PDF plan (a PDF page or a blank page)
 * rather than an image. */
export function isPageItem(item: Pick<PersistedAssemblyItem, "page" | "blank">): boolean {
  return item.page !== undefined || item.blank === true;
}

/** `items` within a tray's limits, in order: images past MAX_ASSEMBLY_ITEMS
 * and anything past MAX_MERGE_PAGES left out (a list of images alone is
 * simply its first MAX_ASSEMBLY_ITEMS). `isPage` tells pages from images. */
export function capTrayItems<T>(items: readonly T[], isPage: (item: T) => boolean): T[] {
  const kept: T[] = [];
  let images = 0;
  for (const item of items) {
    if (kept.length >= MAX_MERGE_PAGES) {
      break;
    }
    if (!isPage(item)) {
      if (images >= MAX_ASSEMBLY_ITEMS) {
        continue;
      }
      images += 1;
    }
    kept.push(item);
  }
  return kept;
}

export interface PersistedAssembly {
  items: PersistedAssemblyItem[];
  layout: AssemblyLayoutSettings;
  /** Null: the sheet the user's assembly preferences start from. */
  sheet: SheetSetup | null;
  marginsLinked: boolean;
  output: AssemblyOutputSettings;
}

/** A plan holding `items` and every other setting at its default (or as
 * `output` says) — what "Combine…" and "New › Combined image" start the
 * workspace with. */
export function freshAssembly(
  items: PersistedAssemblyItem[] = [],
  output: Partial<AssemblyOutputSettings> = {},
): PersistedAssembly {
  return {
    items: capTrayItems(items, isPageItem).map((item) => ({
      id: item.id,
      title: item.title,
      ...(item.page === undefined ? {} : { page: item.page }),
      ...(item.blank ? { blank: true } : {}),
      ...(item.rotateBy ? { rotateBy: item.rotateBy } : {}),
    })),
    layout: { ...ASSEMBLY_LAYOUT_DEFAULTS },
    sheet: null,
    marginsLinked: true,
    output: { ...ASSEMBLY_OUTPUT_DEFAULTS, ...output },
  };
}

function record(raw: unknown): Record<string, unknown> | null {
  return raw !== null && typeof raw === "object" && !Array.isArray(raw) ? (raw as Record<string, unknown>) : null;
}

function finiteIn(raw: unknown, min: number, max: number): number | undefined {
  return typeof raw === "number" && Number.isFinite(raw) && raw >= min && raw <= max ? raw : undefined;
}

function itemOf(raw: unknown): PersistedAssemblyItem | null {
  const value = record(raw);
  const rotateBy = typeof value?.rotateBy === "number" ? normaliseRotation(value.rotateBy) : 0;
  if (value?.blank === true) {
    return { id: "", title: typeof value.title === "string" ? value.title : "", blank: true, ...(rotateBy ? { rotateBy } : {}) };
  }
  if (!value || typeof value.id !== "string" || value.id.length === 0) {
    return null;
  }
  const item: PersistedAssemblyItem = { id: value.id, title: typeof value.title === "string" ? value.title : "" };
  if (value.page !== undefined) {
    if (typeof value.page !== "number" || !Number.isInteger(value.page) || value.page < 0) {
      return null;
    }
    item.page = value.page;
    if (rotateBy) {
      item.rotateBy = rotateBy;
    }
    return item;
  }
  const transform = record(value.transform);
  if (transform) {
    item.transform = normaliseTransform(transform as Partial<OverlayTransform>);
  }
  return item;
}

function layoutOf(raw: unknown): AssemblyLayoutSettings {
  const value = record(raw) ?? {};
  const columns = value.columns === "auto" ? "auto" : finiteIn(value.columns, 1, MAX_ASSEMBLY_ITEMS);
  return {
    layout: LAYOUT_KINDS.includes(value.layout as LayoutKind) ? (value.layout as LayoutKind) : ASSEMBLY_LAYOUT_DEFAULTS.layout,
    columns: columns === undefined ? ASSEMBLY_LAYOUT_DEFAULTS.columns : columns === "auto" ? "auto" : Math.round(columns),
    gapCm: finiteIn(value.gapCm, 0, MAX_GAP_CM) ?? ASSEMBLY_LAYOUT_DEFAULTS.gapCm,
    cellBackground:
      typeof value.cellBackground === "string" && isSheetBackground(value.cellBackground)
        ? value.cellBackground.toLowerCase()
        : ASSEMBLY_LAYOUT_DEFAULTS.cellBackground,
  };
}

/** A stored sheet setup, whole and valid, or null (the plan then starts
 * from the user's preferences rather than a half-repaired sheet). */
function sheetOf(raw: unknown): SheetSetup | null {
  const value = record(raw);
  const margins = record(value?.margins);
  if (!value || !margins) {
    return null;
  }
  const edges = [margins.topCm, margins.rightCm, margins.bottomCm, margins.leftCm].map((edge) => finiteIn(edge, 0, MAX_SHEET_CM));
  const valid =
    SHEET_PAPER_NAMES.includes(value.paper as SheetSetup["paper"]) &&
    (value.orientation === "portrait" || value.orientation === "landscape") &&
    edges.every((edge) => edge !== undefined) &&
    SHEET_FITS.includes(value.fit as SheetSetup["fit"]) &&
    SHEET_ALIGNS.includes(value.align as SheetSetup["align"]) &&
    typeof value.background === "string" &&
    isSheetBackground(value.background) &&
    SHEET_DPI_CHOICES.includes(value.dpi as number);
  if (!valid) {
    return null;
  }
  const [topCm, rightCm, bottomCm, leftCm] = edges as number[];
  const setup: SheetSetup = {
    paper: value.paper as SheetSetup["paper"],
    orientation: value.orientation as SheetSetup["orientation"],
    margins: { topCm, rightCm, bottomCm, leftCm },
    fit: value.fit as SheetSetup["fit"],
    align: value.align as SheetSetup["align"],
    background: (value.background as string).toLowerCase(),
    dpi: value.dpi as number,
  };
  if (setup.paper === "Custom") {
    const widthCm = finiteIn(value.widthCm, Number.MIN_VALUE, MAX_SHEET_CM);
    const heightCm = finiteIn(value.heightCm, Number.MIN_VALUE, MAX_SHEET_CM);
    if (widthCm === undefined || heightCm === undefined) {
      return null;
    }
    setup.widthCm = widthCm;
    setup.heightCm = heightCm;
  }
  return setup;
}

function outputOf(raw: unknown): AssemblyOutputSettings {
  const value = record(raw) ?? {};
  const quality = finiteIn(value.quality, 1, 100);
  return {
    format: ASSEMBLY_FORMATS.includes(value.format as AssemblyFormat) ? (value.format as AssemblyFormat) : ASSEMBLY_OUTPUT_DEFAULTS.format,
    quality: quality === undefined ? ASSEMBLY_OUTPUT_DEFAULTS.quality : Math.round(quality),
    title: typeof value.title === "string" ? value.title : ASSEMBLY_OUTPUT_DEFAULTS.title,
    matte:
      typeof value.matte === "string" && /^#[0-9a-f]{6}$/i.test(value.matte)
        ? value.matte.toLowerCase()
        : ASSEMBLY_OUTPUT_DEFAULTS.matte,
    pageMode: PAGE_MODES.includes(value.pageMode as PageMode) ? (value.pageMode as PageMode) : ASSEMBLY_OUTPUT_DEFAULTS.pageMode,
    nUp: N_UP_CHOICES.includes(value.nUp as number) ? (value.nUp as number) : ASSEMBLY_OUTPUT_DEFAULTS.nUp,
    pageNumbers: typeof value.pageNumbers === "boolean" ? value.pageNumbers : ASSEMBLY_OUTPUT_DEFAULTS.pageNumbers,
    author: typeof value.author === "string" ? value.author : ASSEMBLY_OUTPUT_DEFAULTS.author,
    subject: typeof value.subject === "string" ? value.subject : ASSEMBLY_OUTPUT_DEFAULTS.subject,
  };
}

/** Reads a stored plan: undefined for anything that is not a plan object;
 * otherwise every field repaired to its default where it is wrong, and
 * items without an id dropped (within the tray's limits — capTrayItems). */
export function parsePersistedAssembly(raw: unknown): PersistedAssembly | undefined {
  const value = record(raw);
  if (!value) {
    return undefined;
  }
  const items = Array.isArray(value.items) ? value.items.map(itemOf).filter((item) => item !== null) : [];
  return {
    items: capTrayItems(items, isPageItem),
    layout: layoutOf(value.layout),
    sheet: sheetOf(value.sheet),
    marginsLinked: typeof value.marginsLinked === "boolean" ? value.marginsLinked : true,
    output: outputOf(value.output),
  };
}
