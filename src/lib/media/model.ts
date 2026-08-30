// The in-memory media model — the third product's counterpart of
// documentModel.ts and presentation/model.ts (SPEC-107). A media document is
// a byte payload plus what can be learned about it without rendering it: an
// image's pixel dimensions, a PDF's structural summary. THIS FILE is the
// normative source for `MediaModel` and `MediaFormat`.
//
// Layering rule (enforced by boundary.spec.ts): this directory imports only
// src/lib/core — never documentModel.ts, the commands, the format layers,
// the presentation layer, the store or Vue. Plain serialisable data only, so
// the document store persists all three products' models through the same
// IndexedDB machinery and `JSON.parse(JSON.stringify(model))` `toEqual`s the
// model.
import type { DocumentProperties } from "../core/documentProperties";
import type { MediaPart } from "../core/media";

/** The media formats the library holds: the five image formats every target
 * browser decodes natively, plus PDF. */
export type MediaFormat = "pdf" | "png" | "jpeg" | "gif" | "svg" | "webp";

/** A PDF's structural summary, built by core/pdfDocument's `parsePdf` at
 * open time so the library and viewer never re-parse the bytes. */
export interface PdfSummary {
  /** Header version, e.g. "1.7". */
  version: string;
  pageCount: number;
  /** Per page: the effective CropBox size (points) and normalised rotation. */
  pages: Array<{ widthPt: number; heightPt: number; rotate: 0 | 90 | 180 | 270 }>;
  linearised: boolean;
  /** True when the xref was unusable and the object table was rebuilt by
   * scanning the whole buffer (core/pdfDocument's reconstruction path). */
  reconstructed: boolean;
}

/** The media product's model. The payload lives in `media` as the SPEC-024
 * `MediaPart` the store already splits into its own IndexedDB record — so
 * listing a library of 200 images structured-clones no pixels, and the
 * SPEC-047 reference-identity latch means a rename or a star never rewrites
 * the payload. */
export interface MediaModel {
  /** The store's discriminant, as PresentationModel's `kind` is. */
  kind: "media";
  id: string;
  /** File name without the extension (titleFromFileName). */
  title: string;
  format: MediaFormat;
  /** "application/pdf", "image/png", … */
  contentType: string;
  /** The MediaPart in `media` holding the source bytes, verbatim. */
  sourceMediaId: string;
  /** Exactly one part here; SPEC-112 adds a second. */
  media: MediaPart[];
  /** The source length in bytes, so a row renders without the payload. */
  byteLength: number;
  /** Absent for pdf, and for an image whose dimensions could not be read. */
  image?: { widthPx: number; heightPx: number };
  /** pdf only. */
  pdf?: PdfSummary;
  /** From a PDF's /Info, mapped onto core/documentProperties.ts. */
  properties?: DocumentProperties;
}

/** MIME type for each media format — the `contentType` a fresh MediaModel
 * carries and the type its download is served under. */
export const MEDIA_CONTENT_TYPES: Record<MediaFormat, string> = {
  pdf: "application/pdf",
  png: "image/png",
  jpeg: "image/jpeg",
  gif: "image/gif",
  svg: "image/svg+xml",
  webp: "image/webp",
};

/** The file extension a media document downloads with — the format name
 * itself for every format ("jpeg" being the canonical JPEG spelling, as
 * core/media's writers already pick). */
export function mediaFileExtension(format: MediaFormat): string {
  return format;
}
