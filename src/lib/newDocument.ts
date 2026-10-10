// The landing page's create actions (SPEC-121 AC-1): the ONE place a blank
// document of each product is made and stored. "New document", "New
// presentation" and "New spreadsheet" — and the template gallery's "Blank
// document" and the library palette's library.new-document, which are "New
// document" by another door — all call createBlankDocument, so HomeView
// keeps only the template hookups (it navigates to the path returned here)
// and stays a thin wiring hub (HomeView.size.spec.ts).
//
// App layer, like openDocument.ts and documentStore.ts: it composes every
// product's blank model with the store, which is why it lives here and not
// in a product layer (products never import each other). Each kind is stored
// through its own store accessor carrying its format from birth — a text
// document "new" (getOrCreateDocument, which also applies the "new
// documents" preferences), a deck "pptx" (createBlankPresentation, SPEC-080),
// a workbook "xlsx" (createBlankWorkbook, SPEC-121) — so the library lists,
// badges and routes it like any other record of its kind, and the route
// comes from documentKinds.editorPathFor, never re-encoded here.
import { editorPathFor } from "./documentKinds";
import { getOrCreateDocument, storePresentation, storeSpreadsheet } from "./documentStore";
import { createBlankPresentation } from "./presentation/slides";
import { createBlankWorkbook } from "./spreadsheet/sheets";

/** The kinds the landing page creates from nothing (media is open-only). */
export type BlankDocumentKind = "text" | "presentation" | "spreadsheet";

/** Every BlankDocumentKind, in the landing page's button order. */
export const BLANK_DOCUMENT_KINDS: readonly BlankDocumentKind[] = ["text", "presentation", "spreadsheet"];

/** Stores a new blank document of `kind` under a fresh id and returns the
 * editor path that opens it (`/doc/:id`, `/slides/:id` or `/sheets/:id`). */
export function createBlankDocument(kind: BlankDocumentKind): string {
  if (kind === "presentation") {
    const model = createBlankPresentation();
    storePresentation(model);
    return editorPathFor(kind, model.id);
  }
  if (kind === "spreadsheet") {
    const model = createBlankWorkbook();
    storeSpreadsheet(model);
    return editorPathFor(kind, model.id);
  }
  const doc = getOrCreateDocument(crypto.randomUUID());
  return editorPathFor(kind, doc.id);
}
