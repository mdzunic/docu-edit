// The JSON document-description format (SPEC-045 AC-4): a plain-JSON
// object that maps 1:1 to the builder API's operations — `body` is an array
// of the same block-spec objects `blocks.ts`'s dispatcher already
// understands (the fluent `DocumentBuilder.addHeading`/`addParagraph`/…
// methods call the exact same builder functions), so a document can be
// generated from data with no imperative code at all. The published JSON
// Schema (schema/document-description.json, AC-5) describes this same
// shape for external tooling; this file is its runtime validator.
import { createDocument, type DocumentModel } from "../documentModel";
import {
  BODY_BLOCK_TYPES,
  buildBlocksFromSpecs,
  buildHeaderFooterContent,
  finalizeTocEntries,
  type BuildContext,
} from "./blocks";
import { expectPlainObject, expectString, rejectUnknownKeys } from "./validate";

const DESCRIPTION_KEYS = ["title", "body", "header", "footer"] as const;

/** A JSON document description (AC-4): validated and interpreted by
 * `docFromDescription` into the same canonical `DocumentModel` the fluent
 * builder produces. Every field is optional — an object with none of them
 * describes an empty document (AC-10). */
export interface DocumentDescription {
  title?: string;
  body?: unknown[];
  /** A header block-spec array, or the paragraph-content shorthand (a
   * string or run array) for a plain single-line header. */
  header?: unknown;
  footer?: unknown;
}

/** Validates and interprets a JSON document description into the canonical
 * model (AC-4, AC-6): throws `DocumentGenerationError` on the first invalid
 * field, path-qualified from the description's own root (AC-7, AC-9) —
 * never writes a byte first. */
export function docFromDescription(description: unknown): DocumentModel {
  const obj = expectPlainObject(description, "");
  rejectUnknownKeys(obj, DESCRIPTION_KEYS, "");
  const doc = createDocument();
  doc.sections[0].blocks = [];
  if (obj.title !== undefined) {
    doc.title = expectString(obj.title, "title");
  }
  const ctx: BuildContext = { doc };
  if (obj.body !== undefined) {
    doc.sections[0].blocks = buildBlocksFromSpecs(ctx, obj.body, BODY_BLOCK_TYPES, "body");
  }
  if (obj.header !== undefined) {
    doc.sections[0].headers = { default: buildHeaderFooterContent(ctx, obj.header, "header") };
  }
  if (obj.footer !== undefined) {
    doc.sections[0].footers = { default: buildHeaderFooterContent(ctx, obj.footer, "footer") };
  }
  finalizeTocEntries(doc);
  return doc;
}
