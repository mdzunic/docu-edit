// The deck field model's logic (SPEC-093): the resolver every render and save
// surface reads a footer/slide-number/date value through, plus the deck-level
// apply/visibility mutators the Header & Footer dialog drives. Framework-free
// per CLAUDE.md — imports only presentation siblings and lib/core — so the
// stage (`SlideStage.vue`), the HTML export (`htmlExport.ts`, which lives in
// lib/ and cannot import a component), the build preview and the save writers
// all resolve one identical value.
//
// Determinism: nothing here reads the wall clock; the instant is always passed
// in (rendering passes `now()`, a save passes `WritePptxOptions.now`).
import { DEFAULT_LANGUAGE } from "../core/language";
import { renderDateTime } from "../core/dateTimeFormats";
import type {
  DeckHeaderFooter,
  PresentationModel,
  Slide,
  SlideFieldKind,
  SlideFieldPlaceholder,
} from "./model";
import { markSlideEdited } from "./text";

/** The three deck field kinds, in model order. */
export const SLIDE_FIELD_KINDS: readonly SlideFieldKind[] = ["footer", "slideNumber", "date"];

/** The auto-date default when `DeckHeaderFooter.dateFormat` is absent. */
export const DEFAULT_DATE_FORMAT = "d MMMM yyyy";

/** The bottom-band geometry (cm) and look a field placeholder falls back to
 * when neither the slide nor the layout/master chain gives one (SPEC-093
 * Reference "DEFAULT_FIELD_BAND"): a band `heightCm` tall sitting `bottomInset`
 * above the slide's bottom edge, inset `sideInset` each side, split into three
 * equal boxes with `gap` between them — visually date | footer | number, left
 * | center | right aligned, 12pt #595959. */
export const DEFAULT_FIELD_BAND = {
  heightCm: 0.85,
  bottomInsetCm: 1.27,
  sideInsetCm: 1.27,
  gapCm: 0.42,
  sizePt: 12,
  color: "#595959",
} as const;

/** The band's left-to-right slot order and per-slot alignment. */
const FIELD_BAND_ORDER: readonly SlideFieldKind[] = ["date", "footer", "slideNumber"];
const FIELD_BAND_ALIGN: Record<SlideFieldKind, "left" | "center" | "right"> = {
  date: "left",
  footer: "center",
  slideNumber: "right",
};

/** The default rectangle + alignment a field of `kind` takes in the bottom
 * band of a `slideWidthCm × slideHeightCm` slide (SPEC-093 AC-20). */
export function defaultFieldRect(
  kind: SlideFieldKind,
  slideWidthCm: number,
  slideHeightCm: number,
): { xCm: number; yCm: number; widthCm: number; heightCm: number; align: "left" | "center" | "right" } {
  const { sideInsetCm, gapCm, heightCm, bottomInsetCm } = DEFAULT_FIELD_BAND;
  const bandWidthCm = slideWidthCm - 2 * sideInsetCm;
  const boxWidthCm = (bandWidthCm - 2 * gapCm) / 3;
  const slot = FIELD_BAND_ORDER.indexOf(kind);
  return {
    xCm: sideInsetCm + slot * (boxWidthCm + gapCm),
    yCm: slideHeightCm - bottomInsetCm,
    widthCm: boxWidthCm,
    heightCm,
    align: FIELD_BAND_ALIGN[kind],
  };
}

/** The live slide number for a slide index (SPEC-093 AC-19): the deck's
 * `firstSlideNumber` origin (SPEC-090, absent = 1) plus the slide's position,
 * so numbers stay live under add/delete/move/duplicate without touching any
 * field data. */
export function slideNumberFor(model: PresentationModel, slideIndex: number): number {
  return (model.firstSlideNumber ?? 1) + slideIndex;
}

/** Whether a slide is the deck's title slide (SPEC-093 AC-25): its SPEC-089
 * layout type is `title`, or — when no layout model resolves — index 0 alone. */
export function isTitleSlide(model: PresentationModel, slideIndex: number): boolean {
  const slide = model.slides[slideIndex];
  if (!slide) {
    return false;
  }
  const layout =
    slide.layoutKey !== undefined ? model.layouts?.find((entry) => entry.key === slide.layoutKey) : undefined;
  if (layout) {
    return layout.kind === "title";
  }
  return slideIndex === 0;
}

/** Whether a date field entry is out-of-subset (SPEC-093 "Date formats"): the
 * reader kept its wire type/style verbatim because it maps to no catalogue
 * pattern, so it is rendered from its cache and never recomputed. */
function isPreservedDate(entry: SlideFieldPlaceholder): boolean {
  return entry.kind === "date" && (entry.wireType !== undefined || entry.wireStyleName !== undefined);
}

/** The date string a slide shows, honouring fixed vs auto and the
 * preserve-don't-parse rule for out-of-subset formats. */
function dateText(hf: DeckHeaderFooter | undefined, entry: SlideFieldPlaceholder | undefined, instant: Date): string {
  if (hf?.dateSource === "fixed") {
    return entry?.text ?? hf.dateFixedText ?? "";
  }
  if (entry && isPreservedDate(entry)) {
    // Out-of-subset: pptx cached its `a:t` (kept in `text`); ODF caches nothing,
    // so it falls back to the default pattern. Never recomputed (AC-63/64).
    return entry.text ?? renderDateTime(DEFAULT_DATE_FORMAT, instant, DEFAULT_LANGUAGE) ?? "";
  }
  const format = hf?.dateFormat ?? DEFAULT_DATE_FORMAT;
  return renderDateTime(format, instant, DEFAULT_LANGUAGE) ?? "";
}

/** The string a field of `kind` displays on slide `slideIndex` at `instant`
 * (SPEC-093 AC-18) — the single seam every surface renders through. Returns ""
 * for every kind on a suppressed title slide (AC-24). This resolves the VALUE;
 * whether the field shows at all is the slide's `fields` membership. */
export function fieldText(
  model: PresentationModel,
  slideIndex: number,
  kind: SlideFieldKind,
  instant: Date,
): string {
  const slide = model.slides[slideIndex];
  if (!slide) {
    return "";
  }
  const hf = model.headerFooter;
  if (hf?.showOnTitleSlide === false && isTitleSlide(model, slideIndex)) {
    return "";
  }
  const entry = slide.fields?.find((field) => field.kind === kind);
  if (kind === "slideNumber") {
    return String(slideNumberFor(model, slideIndex));
  }
  if (kind === "footer") {
    return entry?.text ?? hf?.footerText ?? "";
  }
  return dateText(hf, entry, instant);
}

/** The visible field entries a surface renders for a slide, each paired with
 * its resolved text (SPEC-093 AC-27/28). Empty on a suppressed title slide. */
export function visibleFields(
  model: PresentationModel,
  slideIndex: number,
  instant: Date,
): { field: SlideFieldPlaceholder; text: string }[] {
  const slide = model.slides[slideIndex];
  if (!slide?.fields) {
    return [];
  }
  return slide.fields.map((field) => ({ field, text: fieldText(model, slideIndex, field.kind, instant) }));
}

/** The stable `data-field-kind` token for a kind (AC-28): the DOM/CSS spelling
 * of the model kind (`slideNumber` → `slide-number`). */
export function fieldKindToken(kind: SlideFieldKind): "footer" | "slide-number" | "date" {
  return kind === "slideNumber" ? "slide-number" : kind;
}

/** A fresh field entry for a newly-enabled kind: default band geometry + slot
 * alignment, no explicit look (so it round-trips to the same absent-look entry
 * the canonical `p:sp` reads back as). Left alignment is the default, so — like
 * the readers' `alignOf` — only center/right are stored. */
function newFieldEntry(model: PresentationModel, kind: SlideFieldKind): SlideFieldPlaceholder {
  const rect = defaultFieldRect(kind, model.slideWidthCm, model.slideHeightCm);
  const entry: SlideFieldPlaceholder = { kind, xCm: rect.xCm, yCm: rect.yCm, widthCm: rect.widthCm, heightCm: rect.heightCm };
  if (rect.align !== "left") {
    entry.align = rect.align;
  }
  return entry;
}

/** Turns a kind on or off for one slide (SPEC-093 AC-21): presence in
 * `slide.fields` IS visibility. Enabling mints a default-band entry when the
 * slide has none; disabling drops it (and the array when it empties). An
 * existing entry's resolved geometry/look/text is left untouched. */
export function setFieldVisibility(
  model: PresentationModel,
  slide: Slide,
  kind: SlideFieldKind,
  visible: boolean,
): void {
  const fields = slide.fields ?? [];
  const existing = fields.find((field) => field.kind === kind);
  if (visible) {
    if (!existing) {
      slide.fields = [...fields, newFieldEntry(model, kind)];
    }
    return;
  }
  if (existing) {
    const next = fields.filter((field) => field.kind !== kind);
    if (next.length > 0) {
      slide.fields = next;
    } else {
      delete slide.fields;
    }
  }
}

/** The Header & Footer dialog's flat draft (SPEC-093 AC-32..44): every control
 * as a plain value, independent of the model's absent-is-default encoding. */
export interface HeaderFooterValues {
  footer: boolean;
  footerText: string;
  date: boolean;
  dateSource: "auto" | "fixed";
  dateFormat: string;
  dateFixedText: string;
  slideNumber: boolean;
  showOnTitleSlide: boolean;
}

/** Which slides an Apply touches (SPEC-093 AC-47/48). */
export type HeaderFooterScope = { mode: "slide"; index: number } | { mode: "all" };

/** Reads the deck values back into a dialog draft (the model's absent-is-
 * default encoding expanded to explicit control state). Visibility is read
 * from a specific slide's `fields`. */
export function headerFooterDraft(model: PresentationModel, slide: Slide | null): HeaderFooterValues {
  const hf = model.headerFooter;
  const has = (kind: SlideFieldKind): boolean => slide?.fields?.some((field) => field.kind === kind) ?? false;
  return {
    footer: has("footer"),
    footerText: hf?.footerText ?? "",
    date: has("date"),
    dateSource: hf?.dateSource === "fixed" ? "fixed" : "auto",
    dateFormat: hf?.dateFormat ?? DEFAULT_DATE_FORMAT,
    dateFixedText: hf?.dateFixedText ?? "",
    slideNumber: has("slideNumber"),
    showOnTitleSlide: hf?.showOnTitleSlide !== false,
  };
}

/** Builds the deck-level `DeckHeaderFooter` from a dialog draft, storing only
 * the non-default values (the absent-is-default rule). */
function deckValuesFrom(values: HeaderFooterValues): DeckHeaderFooter {
  const hf: DeckHeaderFooter = {};
  if (values.footerText !== "") {
    hf.footerText = values.footerText;
  }
  if (values.dateSource === "fixed") {
    hf.dateSource = "fixed";
  }
  if (values.dateFixedText !== "") {
    hf.dateFixedText = values.dateFixedText;
  }
  if (values.dateFormat !== DEFAULT_DATE_FORMAT) {
    hf.dateFormat = values.dateFormat;
  }
  if (!values.showOnTitleSlide) {
    hf.showOnTitleSlide = false;
  }
  return hf;
}

/** Applies the dialog's values (SPEC-093 AC-22/47/48): writes the deck values,
 * then the per-slide visibility for the scope — this slide, or every slide
 * (also dropping per-slide text overrides). Marks every touched slide edited
 * and sets `headerFooterDirty`. The caller wraps this in record()/commit(). */
export function applyHeaderFooter(
  model: PresentationModel,
  values: HeaderFooterValues,
  scope: HeaderFooterScope,
): void {
  const deck = deckValuesFrom(values);
  if (Object.keys(deck).length > 0) {
    model.headerFooter = deck;
  } else {
    delete model.headerFooter;
  }
  const targets = scope.mode === "all" ? model.slides : model.slides.slice(scope.index, scope.index + 1);
  for (const slide of targets) {
    setFieldVisibility(model, slide, "footer", values.footer);
    setFieldVisibility(model, slide, "slideNumber", values.slideNumber);
    setFieldVisibility(model, slide, "date", values.date);
    if (scope.mode === "all") {
      // Apply to all normalises every slide to the deck values — drop any
      // per-slide text overrides (AC-48).
      for (const field of slide.fields ?? []) {
        delete field.text;
      }
    }
    markSlideEdited(slide, model.format);
  }
  model.headerFooterDirty = true;
}
