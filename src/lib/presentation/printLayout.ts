// The deck print/export sheet planner (SPEC-104, the normative source for the
// handout/notes sheet geometry): which slides a print or export covers, and
// where each lands on which page — pure maths over the model, no DOM. The
// export sheet (DeckPrintSheet.vue) renders exactly these rectangles; the PDF,
// PNG and browser-print paths all read the same plan, so they cannot disagree
// about what a page holds. All lengths are cm.
//
// Layouts, columns × rows: `slides` 1 × 1 full bleed on a page that IS the
// deck's slide size; `notes` the slide fitted into the top half with its
// speaker notes below (overflowing onto notes-only continuation pages);
// `handout-1` 1 × 1, `-2` 1 × 2, `-3` 1 × 3, `-4` 2 × 2, `-6` 2 × 3, `-9`
// 3 × 3 on the chosen paper. Each slot fits the slide at its native aspect,
// centred. With note lines on, the slot splits 55/45 — slide left, ruled
// lines right — which is why 6/9-up cannot carry them and 3-up always does.
import { pageSizeFor, type Orientation, type PaperSizeName, type PageSize } from "../core/pageSetup";
import { parseNumberRange } from "../core/pageRange";
import type { PresentationModel } from "./model";

export type DeckPrintLayout =
  | "slides"
  | "notes"
  | "handout-1"
  | "handout-2"
  | "handout-3"
  | "handout-4"
  | "handout-6"
  | "handout-9";

/** Whether a layout's slots can carry ruled note lines. */
export type NoteLinesRule = "never" | "optional" | "always";

export interface DeckPrintLayoutInfo {
  value: DeckPrintLayout;
  label: string;
  columns: number;
  rows: number;
  noteLines: NoteLinesRule;
}

/** Every layout, in the dialog's order. */
export const DECK_PRINT_LAYOUTS: readonly DeckPrintLayoutInfo[] = [
  { value: "slides", label: "Slides", columns: 1, rows: 1, noteLines: "never" },
  { value: "notes", label: "Notes pages", columns: 1, rows: 1, noteLines: "never" },
  { value: "handout-1", label: "Handout 1 per page", columns: 1, rows: 1, noteLines: "optional" },
  { value: "handout-2", label: "Handout 2 per page", columns: 1, rows: 2, noteLines: "optional" },
  { value: "handout-3", label: "Handout 3 per page", columns: 1, rows: 3, noteLines: "always" },
  { value: "handout-4", label: "Handout 4 per page", columns: 2, rows: 2, noteLines: "optional" },
  { value: "handout-6", label: "Handout 6 per page", columns: 2, rows: 3, noteLines: "never" },
  { value: "handout-9", label: "Handout 9 per page", columns: 3, rows: 3, noteLines: "never" },
];

/** The paper a handout or notes page prints on (the Slides layout ignores it). */
export type DeckPrintPaper = Exclude<PaperSizeName, "Custom">;

export type DeckPrintRange = "all" | "current" | "custom";

export const HANDOUT_MARGIN_CM = 1.27;
export const HANDOUT_GUTTER_CM = 0.6;
/** Ruled note lines: 0.5 pt #808080 rules this far apart. */
export const NOTE_LINE_SPACING_CM = 0.85;
export const NOTE_LINE_WIDTH_PT = 0.5;
export const NOTE_LINE_COLOR = "#808080";
/** "Frame slides": a 0.75 pt black border round each slot's slide. */
export const SLIDE_FRAME_WIDTH_PT = 0.75;
export const SLIDE_FRAME_COLOR = "#000000";
/** The slide's share of a slot's width when note lines are on. */
export const NOTE_LINES_SLIDE_SHARE = 0.55;
/** Handout and notes pages carry a centred page number at this size. */
export const PAGE_NUMBER_SIZE_PT = 9;
/** Speaker notes on a notes page: 11 pt at 1.3 line height. */
export const NOTES_FONT_SIZE_PT = 11;
export const NOTES_LINE_HEIGHT = 1.3;
/** A deliberately generous average glyph advance (in em) for estimating how
 * many characters a notes line holds: overestimating the wrap only costs a
 * little white space, underestimating it would clip text off the page. */
export const NOTES_AVERAGE_ADVANCE_EM = 0.55;

const CM_PER_PT = 2.54 / 72;

export interface DeckPrintOptions {
  layout: DeckPrintLayout;
  paper: DeckPrintPaper;
  orientation: Orientation;
  range: DeckPrintRange;
  /** The custom range text ("1-3, 5"), parsed as slide numbers. */
  rangeText: string;
  /** The editor's current slide (0-based), for `range: "current"`. */
  currentIndex: number;
  includeHidden: boolean;
  /** The user's note-lines choice; `noteLinesApply` decides what prints. */
  noteLines: boolean;
  frameSlides: boolean;
}

export const DEFAULT_DECK_PRINT_OPTIONS: DeckPrintOptions = {
  layout: "slides",
  paper: "A4",
  orientation: "portrait",
  range: "all",
  rangeText: "",
  currentIndex: 0,
  includeHidden: false,
  noteLines: false,
  frameSlides: false,
};

export function layoutInfo(layout: DeckPrintLayout): DeckPrintLayoutInfo {
  return DECK_PRINT_LAYOUTS.find((entry) => entry.value === layout) ?? DECK_PRINT_LAYOUTS[0];
}

/** Whether a layout prints note lines given the user's choice: always for
 * 3-up, never for slides/notes/6-up/9-up, otherwise as chosen. */
export function noteLinesApply(layout: DeckPrintLayout, requested: boolean): boolean {
  const rule = layoutInfo(layout).noteLines;
  return rule === "always" || (rule === "optional" && requested);
}

/** "Frame slides" defaults on for every layout but Slides. */
export function defaultFrameSlides(layout: DeckPrintLayout): boolean {
  return layout !== "slides";
}

/** A sheet page's size: the deck's own slide size for the Slides layout, the
 * chosen paper in the chosen orientation otherwise. */
export function sheetPageSize(
  options: Pick<DeckPrintOptions, "layout" | "paper" | "orientation"> & { slideWidthCm: number; slideHeightCm: number },
): PageSize {
  if (options.layout === "slides") {
    return { widthCm: options.slideWidthCm, heightCm: options.slideHeightCm };
  }
  return pageSizeFor(options.paper, options.orientation);
}

/** Which slides (0-based, deck order) a print covers: the range — all, the
 * current slide, or a custom list of slide numbers — then the hidden filter
 * unless `includeHidden`. An empty result carries the error the dialog shows
 * (and every action is disabled). */
export function resolveSlideRange(
  model: Pick<PresentationModel, "slides">,
  options: Pick<DeckPrintOptions, "range" | "rangeText" | "currentIndex" | "includeHidden">,
): { slides: number[]; error?: string } {
  const total = model.slides.length;
  if (total === 0) {
    return { slides: [], error: "This deck has no slides to print" };
  }
  let picked: number[];
  if (options.range === "current") {
    picked = [Math.min(Math.max(Math.trunc(options.currentIndex) || 0, 0), total - 1)];
  } else if (options.range === "custom") {
    const parsed = parseNumberRange(options.rangeText, total, "slide");
    if (parsed.error !== undefined || !parsed.pages) {
      return { slides: [], error: parsed.error ?? "Enter a slide range" };
    }
    picked = parsed.pages.map((number) => number - 1);
  } else {
    picked = model.slides.map((_, index) => index);
  }
  const slides = options.includeHidden ? picked : picked.filter((index) => model.slides[index]?.hidden !== true);
  if (slides.length === 0) {
    return {
      slides,
      error:
        picked.length === 1
          ? "This slide is hidden — turn on Include hidden slides to print it"
          : "Every slide in this range is hidden — turn on Include hidden slides to print them",
    };
  }
  return { slides };
}

export interface CmRect {
  xCm: number;
  yCm: number;
  widthCm: number;
  heightCm: number;
}

/** One slide placed on a sheet page. */
export interface SheetSlot {
  /** The slide's 0-based deck index. */
  slideIndex: number;
  /** The grid cell the slide sits in. */
  cell: CmRect;
  /** The slide picture: fitted into the cell's slide area at its native
   * aspect, centred. */
  slide: CmRect;
  /** Sheet cm per slide cm (`slide.widthCm / model.slideWidthCm`). */
  scale: number;
  /** "Frame slides": a border round the slide picture. */
  framed: boolean;
  /** The ruled note lines beside the slide: their x extent and each rule's
   * y. Absent when the layout prints none. */
  noteLines?: { xCm: number; widthCm: number; yCm: number[] };
}

/** A notes page's notes block: its box and the paragraphs it holds. */
export interface SheetNotes {
  slideIndex: number;
  box: CmRect;
  paragraphs: string[];
  /** A continuation page: the notes alone, the slide not repeated. */
  continued: boolean;
}

export interface SheetPage {
  /** 0-based position in the sheet. */
  index: number;
  widthCm: number;
  heightCm: number;
  slots: SheetSlot[];
  notes?: SheetNotes;
  /** The 1-based page number label — handout and notes pages only; the
   * Slides layout carries none. */
  pageNumber?: string;
  /** Where the page number is centred (the bottom margin band). */
  pageNumberBox?: CmRect;
}

/** `slideWidthCm × slideHeightCm` fitted into `area`, centred. */
function fitInto(area: CmRect, slideWidthCm: number, slideHeightCm: number): { rect: CmRect; scale: number } {
  const scale = Math.max(0, Math.min(area.widthCm / slideWidthCm, area.heightCm / slideHeightCm));
  const widthCm = slideWidthCm * scale;
  const heightCm = slideHeightCm * scale;
  return {
    rect: {
      xCm: area.xCm + (area.widthCm - widthCm) / 2,
      yCm: area.yCm + (area.heightCm - heightCm) / 2,
      widthCm,
      heightCm,
    },
    scale,
  };
}

function noteLineRows(top: number, height: number): number[] {
  const rows: number[] = [];
  for (let y = NOTE_LINE_SPACING_CM; y <= height + 1e-9; y += NOTE_LINE_SPACING_CM) {
    rows.push(round(top + y));
  }
  return rows;
}

function round(value: number): number {
  return Math.round(value * 1e6) / 1e6;
}

function roundRect(rect: CmRect): CmRect {
  return { xCm: round(rect.xCm), yCm: round(rect.yCm), widthCm: round(rect.widthCm), heightCm: round(rect.heightCm) };
}

/** One slot in a cell: the slide fitted at its native aspect (into the left
 * 55% when note lines are on) and the ruled lines beside it. */
function slotIn(
  cell: CmRect,
  slideIndex: number,
  model: Pick<PresentationModel, "slideWidthCm" | "slideHeightCm">,
  lines: boolean,
  framed: boolean,
): SheetSlot {
  const slideArea: CmRect = lines ? { ...cell, widthCm: cell.widthCm * NOTE_LINES_SLIDE_SHARE } : cell;
  const { rect, scale } = fitInto(slideArea, model.slideWidthCm, model.slideHeightCm);
  const slot: SheetSlot = { slideIndex, cell: roundRect(cell), slide: roundRect(rect), scale, framed };
  if (lines) {
    const xCm = cell.xCm + cell.widthCm * NOTE_LINES_SLIDE_SHARE + HANDOUT_GUTTER_CM / 2;
    slot.noteLines = {
      xCm: round(xCm),
      widthCm: round(cell.xCm + cell.widthCm - xCm),
      yCm: noteLineRows(rect.yCm, rect.heightCm),
    };
  }
  return slot;
}

function pageNumberBand(page: PageSize): CmRect {
  return { xCm: 0, yCm: round(page.heightCm - HANDOUT_MARGIN_CM), widthCm: page.widthCm, heightCm: HANDOUT_MARGIN_CM };
}

/** The notes line height, and how many characters a line of `widthCm` holds
 * (a conservative estimate — see NOTES_AVERAGE_ADVANCE_EM). */
function notesMetrics(widthCm: number): { lineCm: number; charsPerLine: number } {
  const lineCm = NOTES_FONT_SIZE_PT * NOTES_LINE_HEIGHT * CM_PER_PT;
  const advanceCm = NOTES_FONT_SIZE_PT * NOTES_AVERAGE_ADVANCE_EM * CM_PER_PT;
  return { lineCm, charsPerLine: Math.max(1, Math.floor(widthCm / advanceCm)) };
}

/** Greedy word wrap of one paragraph into estimated lines; an empty
 * paragraph is one blank line, an over-long word is broken. */
function wrapParagraph(text: string, charsPerLine: number): string[] {
  if (text.trim().length === 0) {
    return [""];
  }
  const lines: string[] = [];
  let line = "";
  for (const word of text.split(/\s+/).filter((w) => w.length > 0)) {
    let rest = word;
    while (rest.length > charsPerLine) {
      if (line.length > 0) {
        lines.push(line);
        line = "";
      }
      lines.push(rest.slice(0, charsPerLine));
      rest = rest.slice(charsPerLine);
    }
    if (line.length === 0) {
      line = rest;
    } else if (line.length + 1 + rest.length <= charsPerLine) {
      line += ` ${rest}`;
    } else {
      lines.push(line);
      line = rest;
    }
  }
  if (line.length > 0) {
    lines.push(line);
  }
  return lines;
}

/** Splits a slide's notes into per-page paragraph lists: the first page's
 * box holds `firstLines` lines, every continuation `restLines`. A paragraph
 * that straddles a page break continues as its own paragraph on the next. */
export function paginateNotes(
  notes: readonly string[],
  charsPerLine: number,
  firstLines: number,
  restLines: number,
): string[][] {
  const pages: string[][] = [[]];
  let capacity = Math.max(1, firstLines);
  let used = 0;
  for (const paragraph of notes) {
    const lines = wrapParagraph(paragraph, charsPerLine);
    let chunk: string[] = [];
    for (const line of lines) {
      if (used >= capacity) {
        if (chunk.length > 0) {
          pages[pages.length - 1].push(chunk.join(" "));
          chunk = [];
        }
        pages.push([]);
        capacity = Math.max(1, restLines);
        used = 0;
      }
      chunk.push(line);
      used += 1;
    }
    pages[pages.length - 1].push(chunk.join(" "));
  }
  return pages;
}

/** The sheet: one page per slide (Slides), per slide plus notes continuation
 * pages (Notes pages), or per `columns × rows` slides (handouts), over the
 * resolved range. Empty when the range resolves to no slides. */
export function planSheet(
  model: Pick<PresentationModel, "slides" | "slideWidthCm" | "slideHeightCm">,
  options: DeckPrintOptions,
): SheetPage[] {
  const { slides } = resolveSlideRange(model, options);
  const page = sheetPageSize({ ...options, slideWidthCm: model.slideWidthCm, slideHeightCm: model.slideHeightCm });
  const pages: SheetPage[] = [];
  const push = (fields: Omit<SheetPage, "index" | "widthCm" | "heightCm">): void => {
    const sheetPage: SheetPage = { index: pages.length, widthCm: page.widthCm, heightCm: page.heightCm, ...fields };
    if (options.layout !== "slides") {
      sheetPage.pageNumber = String(pages.length + 1);
      sheetPage.pageNumberBox = pageNumberBand(page);
    }
    pages.push(sheetPage);
  };

  if (options.layout === "slides") {
    for (const slideIndex of slides) {
      const cell = { xCm: 0, yCm: 0, widthCm: page.widthCm, heightCm: page.heightCm };
      push({ slots: [slotIn(cell, slideIndex, model, false, options.frameSlides)] });
    }
    return pages;
  }

  const margin = HANDOUT_MARGIN_CM;
  const content: CmRect = {
    xCm: margin,
    yCm: margin,
    widthCm: page.widthCm - 2 * margin,
    heightCm: page.heightCm - 2 * margin,
  };

  if (options.layout === "notes") {
    const half = page.heightCm / 2;
    const slideArea: CmRect = {
      xCm: content.xCm,
      yCm: content.yCm,
      widthCm: content.widthCm,
      heightCm: half - content.yCm - HANDOUT_GUTTER_CM / 2,
    };
    const firstBox: CmRect = {
      xCm: content.xCm,
      yCm: half + HANDOUT_GUTTER_CM / 2,
      widthCm: content.widthCm,
      heightCm: content.yCm + content.heightCm - (half + HANDOUT_GUTTER_CM / 2),
    };
    const { lineCm, charsPerLine } = notesMetrics(content.widthCm);
    for (const slideIndex of slides) {
      const notes = model.slides[slideIndex]?.notes ?? [];
      const chunks = paginateNotes(
        notes,
        charsPerLine,
        Math.floor(firstBox.heightCm / lineCm),
        Math.floor(content.heightCm / lineCm),
      );
      chunks.forEach((paragraphs, chunkIndex) => {
        const continued = chunkIndex > 0;
        push({
          slots: continued ? [] : [slotIn(slideArea, slideIndex, model, false, options.frameSlides)],
          notes: { slideIndex, box: roundRect(continued ? content : firstBox), paragraphs, continued },
        });
      });
    }
    return pages;
  }

  const info = layoutInfo(options.layout);
  const perPage = info.columns * info.rows;
  const lines = noteLinesApply(options.layout, options.noteLines);
  const cellWidth = (content.widthCm - (info.columns - 1) * HANDOUT_GUTTER_CM) / info.columns;
  const cellHeight = (content.heightCm - (info.rows - 1) * HANDOUT_GUTTER_CM) / info.rows;
  for (let start = 0; start < slides.length; start += perPage) {
    const slots = slides.slice(start, start + perPage).map((slideIndex, offset) => {
      const column = offset % info.columns;
      const row = Math.floor(offset / info.columns);
      const cell: CmRect = {
        xCm: content.xCm + column * (cellWidth + HANDOUT_GUTTER_CM),
        yCm: content.yCm + row * (cellHeight + HANDOUT_GUTTER_CM),
        widthCm: cellWidth,
        heightCm: cellHeight,
      };
      return slotIn(cell, slideIndex, model, lines, options.frameSlides);
    });
    push({ slots });
  }
  return pages;
}

/** The deck slides a plan places, in order (each once — a notes
 * continuation page repeats none). */
export function plannedSlides(pages: readonly SheetPage[]): number[] {
  return pages.flatMap((page) => page.slots.map((slot) => slot.slideIndex));
}
