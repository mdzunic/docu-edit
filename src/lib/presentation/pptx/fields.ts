// The pptx wire mapping for deck field placeholders (SPEC-093): the
// `p:ph@type` ⇄ SlideFieldKind correspondence, the provisional (golden-pinned)
// `a:fld@type` ⇄ catalogue-pattern table, and the `p:ph@idx` conventions the
// canonical `p:sp` uses. One source of truth so the reader, the serializer and
// the golden check agree — and so AC-84's "one sanctioned datetime* delta" is
// pinned in exactly one place.
import type { SlideFieldKind } from "../model";

/** `p:ph@type` → deck field kind (`ftr`/`sldNum`/`dt`). */
export const KIND_OF_PH_TYPE: Record<string, SlideFieldKind> = {
  ftr: "footer",
  sldNum: "slideNumber",
  dt: "date",
};

/** Deck field kind → `p:ph@type`. */
export const PH_TYPE_OF_KIND: Record<SlideFieldKind, "ftr" | "sldNum" | "dt"> = {
  footer: "ftr",
  slideNumber: "sldNum",
  date: "dt",
};

/** The `p:ph@idx` the canonical `p:sp` writes per kind (golden-confirmed):
 * footer 10, date 11, slide number 12 — PowerPoint's usual quarter-placeholder
 * indices. */
export const PH_IDX_OF_KIND: Record<SlideFieldKind, string> = {
  footer: "10",
  date: "11",
  slideNumber: "12",
};

/** The `p:cNvPr` name the canonical `p:sp` writes per kind. */
export const PH_NAME_OF_KIND: Record<SlideFieldKind, string> = {
  footer: "Footer Placeholder",
  date: "Date Placeholder",
  slideNumber: "Slide Number Placeholder",
};

/** `a:fld@type` (datetime*) → catalogue pattern. Provisional and
 * locale-dependent — the golden fixture is normative (SPEC-093 "Date formats").
 * `MM/dd/yyyy` and `yyyy-MM-dd` have no faithful datetime* equivalent, so on
 * pptx an auto date in those formats writes a plain `a:r` this editor still
 * recomputes. */
export const PPTX_DATETIME_TO_FORMAT: Record<string, string> = {
  datetime1: "dd/MM/yyyy",
  datetime2: "EEEE, d MMMM yyyy",
  datetime3: "d MMMM yyyy",
  datetime4: "d MMM yyyy",
};

/** Catalogue pattern → `a:fld@type`, or undefined for a format with no
 * faithful datetime* (write a plain recomputed `a:r` instead). */
export const FORMAT_TO_PPTX_DATETIME: Record<string, string> = {
  "dd/MM/yyyy": "datetime1",
  "EEEE, d MMMM yyyy": "datetime2",
  "d MMMM yyyy": "datetime3",
  "d MMM yyyy": "datetime4",
};
