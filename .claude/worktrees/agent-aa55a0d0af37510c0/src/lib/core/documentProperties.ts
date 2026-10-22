// Document metadata types (SPEC-041), product-neutral: the same
// title/subject/author/custom-property vocabulary is carried by every
// package format this repo opens — `.docx`/`.pptx` in `docProps/core.xml`
// and `docProps/custom.xml` (OPC), `.odt`/`.odp` in `meta.xml` (ODF) — so
// the types live in core and each package family's parser/serialiser
// (`core/opcProperties.ts`, `core/odfMeta.ts`) is typed against them.
// Editor-facing helpers (the Properties dialog's draft shape and its
// `.vue`-friendly pass-throughs) stay in `documentModel.ts`.

/** Supported custom-property value types (AC-9): the dialog can add, edit
 * and validate these. `value`'s string encoding depends on `type` — text
 * verbatim, number as its decimal string, date as "YYYY-MM-DD", boolean as
 * "true"/"false" — so the model never needs a union of JS value types and
 * the dialog's inline validation (AC-11) has one shape to parse regardless
 * of type. */
export type CustomPropertyType = "text" | "number" | "date" | "boolean";

/** One custom document property (AC-7..11). */
export interface CustomProperty {
  name: string;
  type: CustomPropertyType;
  value: string;
}

/** A custom property whose source value type this editor does not support
 * editing (AC-10) — e.g. `.docx`'s `vt:i8`/`vt:r4`/`vt:cy` or `.odt`'s
 * `time`/`void` value-types. Its whole source element is kept verbatim (like
 * `OpaquePart`, but scoped to one `<property>`/`<meta:user-defined>` element
 * rather than a whole package part) so it round-trips unchanged without this
 * editor ever having to understand it. */
export interface UnsupportedCustomProperty {
  name: string;
  /** The original type label, for read-only display (e.g. "i8", "time"). */
  rawType: string;
  /** The source `<property>` (.docx) or `<meta:user-defined>` (.odt)
   * element, serialised verbatim. */
  rawXml: string;
}

/** A source metadata field this editor has no dedicated slot for (AC-6:
 * shown "as they exist in the file, including fields the editor does not
 * expose for editing") — one child element of `.docx` `docProps/core.xml`
 * (e.g. `cp:category`, `cp:revision`) or `.odt` `meta.xml`'s `office:meta`
 * (e.g. `meta:generator`) outside the fields this editor models
 * individually. Read-only and preserved unchanged on save. */
export interface ExtraMetadataField {
  /** The source element's local name, used to look up a display label; an
   * unrecognised name falls back to the name itself. */
  name: string;
  /** The source element, serialised verbatim. */
  xml: string;
  /** The source element's text content, for display only (AC-6: "displayed
   * as they exist in the file") — `xml` remains the source of truth written
   * back on save. */
  value: string;
}

/** Document metadata (SPEC-041): the Properties dialog's Description-tab
 * fields, custom properties, and everything else read from the source
 * file's metadata parts (`.docx` `docProps/core.xml`/`docProps/custom.xml`,
 * `.odt` `meta.xml`). Distinct from `DocumentModel.title`, which is the
 * app-bar/library/file-name title (SPEC-002/SPEC-003) — `properties.title`
 * is the separate `dc:title` metadata property a user can set independently,
 * exactly like Word/LibreOffice's own "Title" property vs. file name. */
export interface DocumentProperties {
  title?: string;
  subject?: string;
  keywords?: string;
  /** The Description tab's "Comments" field — `dc:description` (.docx) /
   * `dc:description` (.odt). Unrelated to `DocumentModel.comments`
   * (SPEC-038 review comments). */
  comments?: string;
  /** `dc:creator` (.docx) / `meta:initial-creator` (.odt). */
  author?: string;
  /** `cp:lastModifiedBy` (.docx) / `dc:creator` (.odt — ODF's `dc:creator`
   * names the last person to modify the document, not the original author;
   * `meta:initial-creator` is the original author). */
  lastModifiedBy?: string;
  /** ISO 8601 timestamp, read-only (AC-2): when the file's own metadata says
   * it was created. Set the first time a document is saved to a file and
   * kept stable afterwards; absent for a document that has never been saved
   * (AC-18a). */
  created?: string;
  /** ISO 8601 timestamp, read-only (AC-2): the file's own last-modified
   * stamp, refreshed on every save. */
  modified?: string;
  customProperties?: CustomProperty[];
  unsupportedCustomProperties?: UnsupportedCustomProperty[];
  extraFields?: ExtraMetadataField[];
  /** "Remove personal information on save" (AC-12a..14): persisted with the
   * document so it stays set across sessions/reopens. */
  removePersonalInfoOnSave?: boolean;
}
