// The document-kind registry: the one place the app layer knows which
// products exist and how a library record maps onto them. A "kind" is a
// product (text editor vs presentation viewer/editor vs media viewer vs
// workbook viewer); a
// format is a file type within one. Everything kind-dispatched — the landing
// page's routing, the open funnel, the download dispatch, the store's
// accessors — derives from these functions rather than re-encoding "pptx
// means presentation" locally, so adding a future kind is an edit here plus
// a new product layer, not a hunt across the app (SPEC-107 added "media",
// SPEC-116 "spreadsheet").
import type { DocumentFormat as TextDocumentFormat } from "./documentModel";
import type { PresentationFormat } from "./presentation/model";
import { MEDIA_CONTENT_TYPES, type MediaFormat } from "./media/model";
import type { SpreadsheetFormat } from "./spreadsheet/model";
import { SPREADSHEET_CONTENT_TYPES } from "./spreadsheet/fileTypes";

export type DocumentKind = "text" | "presentation" | "media" | "spreadsheet";

/** Every format the library can hold: the text formats, the presentation
 * formats, the media formats, the spreadsheet formats, and "new" (a text
 * document created in the app, never opened from or saved to a file). Blank
 * presentations do NOT use "new": they carry `"pptx"` from birth
 * (`createBlankPresentation`, SPEC-080), so `kindOfFormat` routes and the
 * library badges them as presentations with no new format. A workbook's
 * template and macro-enabled variants are package variants of "xlsx"/"ods",
 * not formats of their own. */
export type LibraryFormat = TextDocumentFormat | PresentationFormat | MediaFormat | SpreadsheetFormat | "new";

const PRESENTATION_FORMATS: ReadonlySet<string> = new Set(["pptx", "odp"]);

/** The media product's formats (SPEC-107), in display order. */
export const MEDIA_FORMATS: readonly MediaFormat[] = ["pdf", "png", "jpeg", "gif", "svg", "webp"];

const MEDIA_FORMAT_SET: ReadonlySet<string> = new Set(MEDIA_FORMATS);

/** The content types a media file is served under (application/pdf,
 * image/png, …), so the open-from-URL check admits the media kind without
 * knowing any of them itself. */
export const MEDIA_FORMAT_CONTENT_TYPES: readonly string[] = MEDIA_FORMATS.map(
  (format) => MEDIA_CONTENT_TYPES[format],
);

/** The spreadsheet product's formats (SPEC-116), in display order. */
export const SPREADSHEET_FORMATS: readonly SpreadsheetFormat[] = ["xlsx", "ods"];

const SPREADSHEET_FORMAT_SET: ReadonlySet<string> = new Set(SPREADSHEET_FORMATS);

/** The content types a workbook is served under (every format and package
 * variant), so the open-from-URL check admits the spreadsheet kind without
 * knowing any of them itself. */
export const SPREADSHEET_FORMAT_CONTENT_TYPES: readonly string[] = SPREADSHEET_CONTENT_TYPES;

/** The product a library record with this format belongs to. Accepts a
 * plain string (formats often arrive via runtime-typed `.vue` props). */
export function kindOfFormat(format: string): DocumentKind {
  if (PRESENTATION_FORMATS.has(format)) {
    return "presentation";
  }
  if (MEDIA_FORMAT_SET.has(format)) {
    return "media";
  }
  if (SPREADSHEET_FORMAT_SET.has(format)) {
    return "spreadsheet";
  }
  return "text";
}

/** The editor route for a library entry: text documents open in the word
 * processor at /doc/:id, presentations in the presentation view at
 * /slides/:id, media in the media viewer at /media/:id, workbooks in the
 * workbook view at /sheets/:id. */
export function editorPathFor(kind: DocumentKind, id: string): string {
  if (kind === "presentation") {
    return `/slides/${id}`;
  }
  if (kind === "media") {
    return `/media/${id}`;
  }
  if (kind === "spreadsheet") {
    return `/sheets/${id}`;
  }
  return `/doc/${id}`;
}

/** The file extensions the "Open document" pickers and the drag-drop
 * name-check accept, across every kind, in display order. Markdown's two
 * spellings are the SPEC-050 exception documented in openDocument.ts; jpg
 * and jpeg are one format ("jpeg", the canonical spelling) under two
 * conventional file extensions. "key" (SPEC-087) is an import-only
 * extension: a Keynote deck converts one-way into a "pptx" presentation, so
 * it appears here (the open surface) but never as a LibraryFormat. A
 * workbook's six extensions (SPEC-116) are two formats: xlsm/xltx/xltm are
 * "xlsx" package variants and ots an "ods" one — detected from the package,
 * never the name. */
export const OPENABLE_EXTENSIONS = [
  "docx",
  "odt",
  "md",
  "markdown",
  "pptx",
  "odp",
  "key",
  "xlsx",
  "xlsm",
  "xltx",
  "xltm",
  "ods",
  "ots",
  "pdf",
  "png",
  "jpg",
  "jpeg",
  "gif",
  "svg",
  "webp",
] as const;

/** The `accept` attribute for the hidden open-document file inputs. */
export const OPEN_FILE_ACCEPT = OPENABLE_EXTENSIONS.map((ext) => `.${ext}`).join(",");
