// The document-kind registry: the one place the app layer knows which
// products exist and how a library record maps onto them. A "kind" is a
// product (text editor vs presentation viewer/editor); a format is a file
// type within one. Everything kind-dispatched — the landing page's routing,
// the open funnel, the download dispatch, the store's accessors — derives
// from these functions rather than re-encoding "pptx means presentation"
// locally, so adding a future kind (e.g. spreadsheets) is an edit here plus
// a new product layer, not a hunt across the app.
import type { DocumentFormat as TextDocumentFormat } from "./documentModel";
import type { PresentationFormat } from "./presentation/model";

export type DocumentKind = "text" | "presentation";

/** Every format the library can hold: the text formats, the presentation
 * formats, and "new" (a text document created in the app, never opened from
 * or saved to a file). Blank presentations do NOT use "new": they carry
 * `"pptx"` from birth (`createBlankPresentation`, SPEC-080), so `kindOfFormat`
 * routes and the library badges them as presentations with no new format. */
export type LibraryFormat = TextDocumentFormat | PresentationFormat | "new";

const PRESENTATION_FORMATS: ReadonlySet<string> = new Set(["pptx", "odp"]);

/** The product a library record with this format belongs to. Accepts a
 * plain string (formats often arrive via runtime-typed `.vue` props). */
export function kindOfFormat(format: string): DocumentKind {
  return PRESENTATION_FORMATS.has(format) ? "presentation" : "text";
}

/** The editor route for a library entry: text documents open in the word
 * processor at /doc/:id, presentations in the presentation view at
 * /slides/:id. */
export function editorPathFor(kind: DocumentKind, id: string): string {
  return kind === "presentation" ? `/slides/${id}` : `/doc/${id}`;
}

/** The file extensions the "Open document" pickers and the drag-drop
 * name-check accept, across every kind, in display order. Markdown's two
 * spellings are the SPEC-050 exception documented in openDocument.ts. */
export const OPENABLE_EXTENSIONS = [
  "docx",
  "odt",
  "md",
  "markdown",
  "pptx",
  "odp",
] as const;

/** The `accept` attribute for the hidden open-document file inputs. */
export const OPEN_FILE_ACCEPT = OPENABLE_EXTENSIONS.map((ext) => `.${ext}`).join(",");
