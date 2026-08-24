// The in-memory presentation model — the second product's counterpart of
// documentModel.ts, deliberately minimal for now. The preparation milestone
// (open `.pptx`/`.odp`, list it in the library, preview its slides, save it
// back with nothing lost) models only what that needs: a read-only text
// outline per slide, plus the entire source package retained verbatim.
// Editing specs will grow this model the way SPEC-011+ grew the text one.
//
// Layering rule (enforced by boundary.spec.ts): this directory imports only
// src/lib/core — never documentModel.ts, the commands, or the .docx/.odt
// format layers. Plain serialisable data only, like documentModel.ts, so
// the document store persists both products' models through the same
// IndexedDB machinery.
import type { DocumentProperties } from "../core/documentProperties";
import type { MediaPart } from "../core/media";
import type { OpaquePart } from "../core/packageParts";

/** The presentation file formats, by the package family they ride in:
 * `.pptx` is OPC/PresentationML, `.odp` is an ODF package. */
export type PresentationFormat = "pptx" | "odp";

/** One slide's read-only preview: what the stub presentation view renders. */
export interface SlidePreview {
  /** The slide's title text, when a title placeholder (`.pptx`
   * `p:ph type="title"|"ctrTitle"`, `.odp` `presentation:class="title"`)
   * carries one. */
  title?: string;
  /** The slide's remaining visible text, one entry per paragraph, in
   * document order (title text excluded). */
  lines: string[];
}

/** The root of a presentation. `kind` is the literal discriminant that
 * tells it apart from a text `DocumentModel` (which has no `kind` field)
 * wherever the document store handles both. */
export interface PresentationModel {
  kind: "presentation";
  id: string;
  /** The app-bar/library title (file name without extension), same
   * convention as `DocumentModel.title`. */
  title: string;
  format: PresentationFormat;
  slides: SlidePreview[];
  /** Every part of the source package, verbatim and in package order —
   * SPEC-008's retention rule applied wholesale: nothing is editable yet,
   * so nothing is rebuilt on save and the writer re-emits exactly these.
   * (`.odp`'s `mimetype` entry is included and written first + STOREd.) */
  opaque: OpaquePart[];
  /** Reserved for the editing milestone (image bytes split out the way
   * `DocumentModel.media` is); never set by the preparation readers. The
   * field exists so the store's media split/rejoin machinery types over
   * both products' models. */
  media?: MediaPart[];
  /** Document metadata read from `docProps/core.xml` (.pptx) / `meta.xml`
   * (.odp); absent when the package declares none. */
  properties?: DocumentProperties;
}

export const DEFAULT_PRESENTATION_TITLE = "Untitled presentation";

/** What a presentation reader hands back — the shared assembly input of
 * `assemblePresentation` (presentation/open.ts), the same split as the text
 * side's `ParsedDocxDocument`/`ParsedOdtDocument` vs `assembleDocumentModel`. */
export interface ParsedPresentation {
  slides: SlidePreview[];
  /** Every source package part, verbatim, in package order. */
  opaque: OpaquePart[];
  /** Metadata from `docProps/core.xml`/`meta.xml`; `{}` when none. */
  properties: DocumentProperties;
}

/** How many slides the library/preview surfaces report for a model. */
export function slideCount(model: PresentationModel): number {
  return model.slides.length;
}
