// The in-memory document model: a serialisable tree every later spec reads
// and writes. Plain data only (no classes, no functions on the objects) so
// `JSON.stringify`/`JSON.parse` round-trips it without loss — that property
// is what SPEC-006/007 (save) and SPEC-004/005 (open) will rely on.

/** Sub/superscript (`w:vertAlign` / `style:text-position`). Mutually
 * exclusive by construction: one field holds either value (SPEC-013). */
export type VerticalAlignment = "subscript" | "superscript";

/** Character-level formatting applied to a run's text. Empty until SPEC-012
 * (character formatting) and later specs add fields such as bold/italic. */
export interface RunFormatting {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strikethrough?: boolean;
  /** e.g. "Calibri" */
  fontFamily?: string;
  fontSizePt?: number;
  /** CSS-ready colour, e.g. "#ff0000". */
  color?: string;
  /** OOXML highlight colour name ("yellow") for the 16-colour palette, a
   * lowercase "#rrggbb" for any other colour, or `false` for an explicit
   * "No colour" — distinct from undefined (not set) the same way the
   * boolean toggles distinguish explicit-off from unset (SPEC-013). */
  highlight?: string | false;
  /** Run shading (`w:shd` fill) as lowercase "#rrggbb" — a preservation
   * slot set only when a source run carried both `w:highlight` and `w:shd`
   * (SPEC-013). Never shown in the picker; the highlight wins on screen. */
  shadingFill?: string;
  vertAlign?: VerticalAlignment;
  /** External hyperlink target (SPEC-025) — an absolute URL. Two scalar
   * fields rather than one `{href}` object so run coalescing and formatting
   * equality keep working on `===` per key (see textEditing's
   * sameFormatting). At most one of `linkHref`/`linkAnchor` is set. */
  linkHref?: string;
  /** Internal link target (SPEC-025): a bookmark name or a derived heading
   * anchor (see lib/links.ts) in this document. */
  linkAnchor?: string;
}

/** The run-formatting fields SPEC-012's toolbar toggles: bold, italic,
 * underline and strikethrough. */
export type CharacterFormatKey = "bold" | "italic" | "underline" | "strikethrough";

/** A collapsed caret's not-yet-typed format override (SPEC-012 AC-9..11):
 * a key present here forces that format for the next inserted text
 * regardless of the run the caret sits in; an absent key falls through to
 * that run's own formatting. Lives outside the document model — there is
 * nothing to persist until text is actually typed with it.
 *
 * SPEC-013 widens it to a general run-format patch, also the argument to
 * the selection-formatting command. Per-key off semantics: `false` on a
 * boolean toggle or on `vertAlign` removes the key (the text ends up as if
 * never formatted); `false` on `highlight` is stored — "No colour" is an
 * explicit off that must survive saving (SPEC-013). */
export interface PendingFormat {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strikethrough?: boolean;
  fontFamily?: string;
  fontSizePt?: number;
  color?: string;
  highlight?: string | false;
  vertAlign?: VerticalAlignment | false;
}

/** Typed pass-through for a character-format key coming from a `.vue`
 * template's plain string (a click handler's event payload) — .vue scripts
 * are parsed by espree, so a type guard/annotation isn't available there
 * (see lib/opaque.ts's `asDropRecords` for the same pattern). */
export function asCharacterFormatKey(value: string): CharacterFormatKey {
  return value as CharacterFormatKey;
}

/** Typed pass-through for an alignment coming from a `.vue` template's
 * plain string (same espree constraint as asCharacterFormatKey). */
export function asParagraphAlignment(value: string): ParagraphAlignment {
  return value as ParagraphAlignment;
}

/** Typed pass-through for a border line style coming from a `.vue`
 * template's plain string (same espree constraint as asCharacterFormatKey). */
export function asBorderLineStyle(value: string): BorderLineStyle {
  return value as BorderLineStyle;
}

/** Typed pass-through for a cell vertical alignment coming from a `.vue`
 * template's plain string (same espree constraint as asCharacterFormatKey). */
export function asCellVerticalAlign(value: string): CellVerticalAlign {
  return value as CellVerticalAlign;
}

/** Typed pass-through for a table alignment coming from a `.vue` template's
 * plain string (same espree constraint as asCharacterFormatKey). */
export function asTableAlignment(value: string): TableAlignment {
  return value as TableAlignment;
}

/** Typed pass-through for a table width mode coming from a `.vue`
 * template's plain string (same espree constraint as asCharacterFormatKey). */
export function asTableWidthMode(value: string): TableWidthMode {
  return value as TableWidthMode;
}

/** The smallest unit of text: a string plus the formatting applied to it.
 * Tabs and forced line breaks are preserved as literal "\t" / "\n"
 * characters in `text` — renderers use `white-space: pre-wrap` so both
 * display exactly as the source document intended. */
export interface Run {
  text: string;
  formatting: RunFormatting;
}

export type ParagraphAlignment = "left" | "center" | "right" | "justify";

/** Paragraph writing direction; absent means `"ltr"`. Readers only ever
 * record `"rtl"` (SPEC-014) — an explicit `"ltr"` never reaches the model,
 * which keeps the field round-trip-stable across both writers. */
export type ParagraphDirection = "ltr" | "rtl";

/** Block-level formatting applied to a paragraph (SPEC-014). Indents are
 * held as *logical* start/end values: `direction` decides which physical
 * edge each paints on — in the renderer, the ruler and both writers. */
export interface ParagraphFormatting {
  alignment?: ParagraphAlignment;
  /** Indent on the edge the text starts from (left for LTR, right for RTL). */
  indentStartCm?: number;
  /** Indent on the edge the text runs toward. */
  indentEndCm?: number;
  /** Positive indents the first line; negative produces a hanging indent.
   * Start-relative in both file formats, so no direction mapping applies. */
  indentFirstLineCm?: number;
  spaceBeforePt?: number;
  spaceAfterPt?: number;
  /** Line height as a multiplier of the font size (CSS `line-height`-style),
   * used when the source expresses spacing as a proportional line rule. */
  lineSpacing?: number;
  /** Line height as an absolute point size, used when the source expresses
   * spacing as a fixed/minimum line rule. Takes precedence over
   * `lineSpacing` when both would otherwise apply. */
  lineSpacingPt?: number;
  /** How `lineSpacingPt` binds: "exact" is a fixed line height ("Exactly" in
   * the spacing dialog); absent means "at least" — the historical default
   * both writers emitted, kept as the absent value so models saved before
   * SPEC-014 keep their meaning. Meaningless without `lineSpacingPt`. */
  lineSpacingRule?: "exact";
  direction?: ParagraphDirection;
  /** Keep this paragraph on the same page as the start of the next block
   * (SPEC-028 AC-22/23, `w:keepNext` / `fo:keep-with-next`). */
  keepNext?: boolean;
  /** Never split this paragraph's lines across a page boundary (SPEC-028
   * AC-24/25, `w:keepLines` / `style:keep-together`). */
  keepLines?: boolean;
  /** Start a new page before this paragraph (SPEC-028 AC-26/27,
   * `w:pageBreakBefore` / `fo:break-before` on the paragraph style). This is
   * paragraph formatting, distinct from a manual break (BreakBlock). */
  pageBreakBefore?: boolean;
  /** Widow/orphan control (SPEC-028 AC-28/29). Absent means ON — both file
   * formats default it on, so only the explicit-off (`false`) and an
   * explicit re-enable after an off (`true`) are ever stored. */
  widowControl?: boolean;
}

// --- lists (SPEC-018) -------------------------------------------------------

export type ListType = "bullet" | "number";

/** The marker styles the editor offers (SPEC-018 AC-12): three bullet
 * characters and five number formats. */
export type ListMarkerStyle =
  | "disc"
  | "circle"
  | "square"
  | "decimal"
  | "lowerAlpha"
  | "upperAlpha"
  | "lowerRoman"
  | "upperRoman";

/** A paragraph's membership in a bulleted or numbered list (SPEC-018).
 * List *identity* is adjacency: consecutive list paragraphs form one list,
 * and a non-list paragraph between two list items splits them into two
 * lists whose numbering restarts (AC-10, AC-11) — so no list id is stored.
 * Displayed numbers are always derived by scanning the block sequence
 * (lib/lists.ts), never persisted. */
export interface ListFormatting {
  type: ListType;
  /** 1-based nesting depth, 1..9 (AC-3, AC-4). */
  level: number;
  /** Explicit marker style for this item's level (AC-12, AC-13); absent
   * means the level's default cycle (AC-5, see lists.ts). */
  marker?: ListMarkerStyle;
  /** The number the item's list counts from at this level; absent means 1.
   * Carried on every item of an imported list so a restart after an
   * interruption returns to the list's own start value (AC-11, AC-22). */
  startAt?: number;
  /** The level's text indent in cm as defined by an imported numbering
   * definition (AC-20, AC-21); absent means the editor default
   * (level × 1.27 cm, see lists.ts). */
  indentCm?: number;
}

/** Typed pass-through for a list type coming from a `.vue` template's plain
 * string (same espree constraint as asCharacterFormatKey). */
export function asListType(value: string): ListType {
  return value as ListType;
}

/** Typed pass-through for a marker style coming from a `.vue` template's
 * plain string (same espree constraint as asCharacterFormatKey). */
export function asListMarkerStyle(value: string): ListMarkerStyle {
  return value as ListMarkerStyle;
}

// --- paragraph styles (SPEC-019) --------------------------------------------

/** One paragraph style of the document's style catalogue (SPEC-019): its
 * *own* (non-inherited) formatting plus the parent link resolution follows
 * (`w:basedOn` / `style:parent-style-name`). For a `.docx` these mirror the
 * paragraph `w:style` elements of styles.xml; for a `.odt`, the named
 * paragraph styles of its styles.xml. The model keeps paragraph/run
 * formatting *baked* (effective) on each paragraph, so the catalogue is not
 * needed to render — it drives the style picker, the diff that separates
 * direct formatting from style formatting, and the writers' style
 * definitions. */
export interface NamedStyle {
  /** Style identity as referenced by `Paragraph.styleId`. */
  id: string;
  /** Display name shown in the picker. */
  name?: string;
  /** Parent style id the style inherits from. */
  basedOn?: string;
  /** The style's own paragraph formatting; absent when it sets none. */
  paragraphFormatting?: ParagraphFormatting;
  /** The style's own run formatting; absent when it sets none. */
  runFormatting?: RunFormatting;
  /** ODF-only: a proportional font size ("130%" in the source), resolved
   * against the inherited size during chain resolution. */
  fontSizePercent?: number;
  /** 1-based outline level; presence makes this a heading style. */
  headingLevel?: number;
}

/** An inline XML element the reader does not model (SPEC-008) — e.g. a
 * `w:drawing`-carrying run or an ODF `draw:frame` — held verbatim and
 * anchored to its position between the paragraph's runs so the writer can
 * re-emit it in place. `xml` is a self-contained serialisation (namespace
 * declarations included); it is compared with XML canonicalization, never
 * byte-for-byte. */
export interface OpaqueInline {
  xml: string;
  /** The node sits immediately before `runs[runIndex]`; `runs.length` means
   * after the last run. */
  runIndex: number;
}

/** A named point location in a paragraph's text (SPEC-025) — the model form
 * of `w:bookmarkStart/End` (.docx) and `text:bookmark` (.odt). Bookmarks are
 * points, not ranges: the editor's "Insert bookmark" names a location, and a
 * ranged source bookmark is anchored at its start. Deliberately not named
 * `id`: lib/equality.ts's IGNORED_KEYS drops bare "id" keys from structural
 * comparison, which would let a wrong reference pass a round-trip test. */
export interface BookmarkAnchor {
  /** Unique document-wide, no spaces (validated in lib/links.ts). */
  name: string;
  /** Character offset into the paragraph's concatenated run text. */
  offset: number;
}

/** A paragraph is a block-level element made of inline runs. */
export interface Paragraph {
  type: "paragraph";
  runs: Run[];
  formatting?: ParagraphFormatting;
  /** Bulleted/numbered list membership (SPEC-018); absent for a normal
   * paragraph. */
  list?: ListFormatting;
  /** Raw style reference (e.g. `w:pStyle`'s value), kept alongside the
   * resolved `formatting`/run formatting so later specs (save) can re-emit
   * a style reference instead of only direct formatting. */
  styleId?: string;
  /** 1-based outline level resolved from the paragraph's style chain;
   * undefined for body text. */
  headingLevel?: number;
  /** Unmodelled inline XML preserved through a same-format save (SPEC-008).
   * A side-list (rather than a run union) so `runs` stays a plain `Run[]`
   * for every renderer and editing consumer; order follows `runIndex`. */
  opaqueInlines?: OpaqueInline[];
  /** Named point locations in this paragraph (SPEC-025), ordered by
   * `offset`. A side-list like `opaqueInlines` so `runs` stays plain;
   * offsets are maintained by the textEditing primitives. */
  bookmarks?: BookmarkAnchor[];
}

// --- images (SPEC-024) -------------------------------------------------------

/** How an image sits relative to the surrounding text (SPEC-024 AC-13,
 * AC-14): `"inline"` flows with the text like a character; `"square"` floats
 * with text wrapping around its rectangle; `"behind"`/`"infront"` float
 * without wrapping, stacked behind or in front of the text. */
export type ImageWrapMode = "inline" | "square" | "behind" | "infront";

/** An image block (SPEC-024): a block-level object addressed by `CaretPosition`
 * the same way `Table` is — a length-1 atom (see textEditing.blockLength) a
 * caret sits before or after, never inside. Real .docx/.odt inline images
 * live mid-run in the source XML; the reader splits the paragraph around one
 * on open (mirroring how a body-level `w:tbl` already breaks the block
 * sequence) so every image, inline or floating, is one section block here —
 * one caret/selection/rendering model for both, instead of a second
 * inline-anchor system alongside `OpaqueInline`. `"inline"` wrap renders as a
 * plain block (no float); `"square"` floats with text wrapping around it;
 * `"behind"`/`"infront"` float out of flow, stacked behind or in front of
 * surrounding text (AC-13, AC-14). */
export interface Image {
  type: "image";
  /** References a `DocumentModel.media` entry's `mediaId`. */
  mediaId: string;
  /** The source image's pixel dimensions, fixed at insertion time — the
   * "Reset size" button's target (AC-21) and the fit-to-text-area check's
   * input (AC-5, AC-6), both independent of the current display size below. */
  naturalWidthPx: number;
  naturalHeightPx: number;
  /** Current display size in cm (AC-16, AC-19, AC-20). */
  widthCm: number;
  heightCm: number;
  wrap: ImageWrapMode;
  /** `wp:docPr/@descr` (.docx) / `svg:desc` (.odt) — AC-7, AC-23..25. */
  altText?: string;
}

/** Typed pass-through for a wrap mode coming from a `.vue` template's plain
 * string (same espree constraint as asCharacterFormatKey). */
export function asImageWrapMode(value: string): ImageWrapMode {
  return value as ImageWrapMode;
}

/** Whether a block is an image — the isTable-style narrowing helper (also
 * usable from espree-parsed `.vue` scripts, which cannot write the type
 * predicate themselves). */
export function isImage(block: Block | undefined): block is Image {
  return block?.type === "image";
}

/** Typed pass-through for an Image arriving as a runtime-declared Object prop
 * (espree bars the type annotation — see tables.ts's asTable for the same
 * pattern). */
export function asImage(value: unknown): Image {
  return value as Image;
}

/** A binary image payload stored once and referenced by `Image.mediaId` — so
 * an image used twice in the document (AC-32) shares one entry, and deleting
 * one use does not by itself remove it (see lib/images.ts's reference-counted
 * pruning, AC-31). `mediaId` is deliberately not named `id`: lib/equality.ts's
 * IGNORED_KEYS drops any bare "id" key from structural comparison, which
 * would let a wrong mediaId reference pass a round-trip test unnoticed. */
export interface MediaPart {
  mediaId: string;
  /** MIME type — "image/png", "image/jpeg", "image/gif", or "image/svg+xml"
   * for a `.odt` (never `.docx`: SVG is converted to PNG on insert, AC-2). */
  contentType: string;
  /** The image bytes, base64-encoded so the model stays JSON-serialisable
   * (see opaque.ts's OpaquePart, the same convention). */
  data: string;
}

// --- tables (SPEC-022, cell/border formatting SPEC-023) ---------------------

/** A border line's rendering (SPEC-023): the standard presets plus width and
 * colour. `"none"` paints nothing, distinct from an absent `BorderSide`
 * (unset — falls back to whatever the renderer/writer defaults to). */
export type BorderLineStyle = "none" | "single" | "double" | "dashed" | "dotted";

/** One edge of a cell's border (SPEC-023 AC-10..12): style, width in points
 * and a CSS-ready colour ("#rrggbb", or "auto" for the source's automatic
 * colour). */
export interface BorderSide {
  style: BorderLineStyle;
  widthPt: number;
  color: string;
}

/** A cell's four borders, each independently set or left unset (SPEC-023).
 * Table-wide and per-selected-cell border operations (AC-13, AC-14) both
 * write into this same per-cell shape — "apply to the whole table" is just
 * "apply to every cell" (see lib/tables.ts's border-preset commands), which
 * is also why the round trip is always per-cell (`w:tcBorders` /
 * .odt cell styles — AC-25, AC-29), never a table-level default. */
export interface CellBorders {
  top?: BorderSide;
  bottom?: BorderSide;
  left?: BorderSide;
  right?: BorderSide;
}

/** A cell's content vertical alignment within its (possibly taller, due to
 * sibling cells) row (SPEC-023 AC-17). */
export type CellVerticalAlign = "top" | "center" | "bottom";

/** One table cell: a nested block container holding paragraphs only (nested
 * tables are out of scope — SPEC-022). A real (non-covered) cell always has
 * at least one paragraph, so a caret can land in it. */
export interface TableCell {
  blocks: Paragraph[];
  /** Grid columns this cell spans rightward, including itself; absent means
   * 1 (SPEC-023 AC-6, AC-24, AC-28). */
  colSpan?: number;
  /** Grid rows this cell spans downward, including itself; absent means 1
   * (SPEC-023 AC-6, AC-24, AC-28). */
  rowSpan?: number;
  /** True when this grid slot is covered by a neighbouring cell's
   * `colSpan`/`rowSpan` (SPEC-023): the slot has no content of its own
   * (`blocks` is empty) and is never an editing/navigation target — the
   * spanning cell's own `<td>` (or `w:tc`/`table:covered-table-cell`) covers
   * it visually and on disk. */
  covered?: boolean;
  /** This cell's own border overrides (SPEC-023 AC-10..14). */
  borders?: CellBorders;
  /** Background shading as a CSS-ready "#rrggbb" (SPEC-023 AC-15). */
  shadingFill?: string;
  /** Cell padding in cm, all four sides (SPEC-023 AC-16); absent uses the
   * renderer/writer's own default padding. */
  paddingCm?: number;
  /** Content's vertical alignment within the cell (SPEC-023 AC-17). */
  verticalAlign?: CellVerticalAlign;
  /** Verbatim `w:tcPr` from a .docx source, re-emitted on a same-format
   * save so unmodelled cell properties survive (SPEC-008's fidelity rule).
   * Cleared by any SPEC-023 command that edits this cell's modelled
   * properties, so the writer regenerates `w:tcPr` from them instead of
   * reproducing stale XML. Absent on cells created/edited in the editor. */
  docxTcPr?: string;
  /** `table:style-name` of the source .odt cell — re-emitted so the
   * preserved automatic style (sourceXml.odtAutomaticStyles) applies
   * again. Cleared the same way as `docxTcPr` when a command edits the
   * cell's modelled properties. */
  odtStyleName?: string;
}

/** One table row. */
export interface TableRow {
  cells: TableCell[];
  /** The row's minimum height in cm (SPEC-023 AC-4, AC-5): dragging a row
   * boundary sets this floor, but the row still grows taller to fit its
   * content — it is never a fixed height. */
  minHeightCm?: number;
  /** Verbatim `w:trPr` from a .docx source (row height, header flag, …). */
  docxTrPr?: string;
  /** `table:style-name` of the source .odt row. */
  odtStyleName?: string;
}

/** Table alignment on the page (SPEC-023 AC-22). */
export type TableAlignment = "left" | "center" | "right";

/** How a table's overall width is determined (SPEC-023 AC-23): `"fixed"`
 * (absent also means fixed, for models saved before SPEC-023) uses the
 * `columnWidthsCm` sum as-is; `"percent"` scales `widthPercent`% of the
 * text area, redistributing `columnWidthsCm` proportionally; `"auto"` sizes
 * the table to its content, ignoring `columnWidthsCm` for layout (the field
 * is kept as the last explicit widths, reused if the mode changes back). */
export type TableWidthMode = "auto" | "fixed" | "percent";

/** A table block (SPEC-022, formatting SPEC-023): a rectangular grid of
 * cells. Every row has exactly `columnWidthsCm.length` cells — a merged
 * cell's covered neighbours are present as placeholder (`covered: true`)
 * cells, so the grid stays regular for indexing purposes even though it
 * renders and saves as merged. */
export interface Table {
  type: "table";
  rows: TableRow[];
  /** Column widths in cm, one per column — the `w:tblGrid` / ODF
   * table-column definitions both readers resolve and both writers emit. */
  columnWidthsCm: number[];
  /** The table's alignment on the page (SPEC-023 AC-22); absent means left
   * (both formats' default). */
  alignment?: TableAlignment;
  /** How the table's width is determined (SPEC-023 AC-23); absent means
   * `"fixed"`. */
  widthMode?: TableWidthMode;
  /** Percentage (0..100) of the text area the table spans; meaningful only
   * when `widthMode` is `"percent"` (SPEC-023 AC-23). */
  widthPercent?: number;
  /** Verbatim `w:tblPr` from a .docx source (table style ref, look flags,
   * …); re-emitted on a same-format save for whatever SPEC-023 does not
   * model. Cleared by any command that edits `alignment`/`widthMode`/
   * `widthPercent`, so the writer regenerates the affected parts of
   * `w:tblPr` instead of reproducing stale XML. */
  docxTblPr?: string;
  /** `table:name` of the source .odt table (user-facing table name). */
  odtName?: string;
  /** `table:style-name` of the source .odt table. */
  odtStyleName?: string;
  /** Per-column `table:style-name`s of the source .odt table's
   * `table:table-column` elements; entries may be undefined for columns
   * inserted in the editor (the writer generates a style for those). */
  odtColumnStyleNames?: Array<string | undefined>;
}

/** Whether a block is a table — the one narrowing helper every consumer of
 * the widened Block union shares (also usable from espree-parsed `.vue`
 * scripts, which cannot write the type predicate themselves). */
export function isTable(block: Block | undefined): block is Table {
  return block?.type === "table";
}

/** Typed pass-through for a Block a `.vue` template renders down the
 * paragraph branch — the branch itself guards on the block kind, but the
 * espree-parsed script cannot express the narrowing (same constraint as
 * asCharacterFormatKey). */
export function asParagraphBlock(value: Block): Paragraph {
  return value as Paragraph;
}

/** The block as a paragraph, throwing on anything else — the narrowing for
 * call sites that know the block kind by construction (tests and fixture
 * builders), where silently reading a table as a paragraph would hide a
 * real defect. */
export function asParagraph(block: Block | undefined): Paragraph {
  if (!block || block.type !== "paragraph") {
    throw new Error("expected a paragraph block");
  }
  return block;
}

/** A block-level XML element the reader does not model (SPEC-008) — e.g. a
 * block-level `w:sdt` content control, a complex (merged-cell/nested)
 * table, or an ODF `draw:frame` — held verbatim, anchored to its position
 * among the section's blocks, and re-emitted in place on a same-format
 * save. */
export interface OpaqueBlock {
  /** Self-contained XML serialisation, compared canonically (see
   * OpaqueInline.xml). */
  xml: string;
  /** The node sits immediately before `blocks[blockIndex]`; `blocks.length`
   * means after the last block. */
  blockIndex: number;
}

// --- breaks and sections (SPEC-028) -----------------------------------------

/** The manual break kinds Insert → Break offers (SPEC-028 AC-9/10). Column
 * breaks are stored and round-tripped but have no layout effect — they sit
 * inertly in the flow (marker visible when formatting marks are on) until
 * columns exist (SPEC-030). */
export type BreakKind = "page" | "column";

/** A manual break (SPEC-028): a block-level atom (blockLength 1, like Image)
 * sitting between the content before and after it. Pagination ends the
 * current page at a `"page"` break. Round-trips as a `w:br w:type="page"`
 * run (.docx) / `fo:break-before` on the following paragraph (.odt); a
 * source break in the middle of a paragraph splits the paragraph on open,
 * mirroring how inline images become blocks (see Image). */
export interface BreakBlock {
  type: "break";
  kind: BreakKind;
}

/** Whether a block is a manual break — the isTable-style narrowing helper
 * (also usable from espree-parsed `.vue` scripts). */
export function isBreak(block: Block | undefined): block is BreakBlock {
  return block?.type === "break";
}

/** How a section begins relative to the previous one (SPEC-028 AC-8/9):
 * `"nextPage"` starts on a fresh page, `"continuous"` continues on the same
 * page. Mirrors OOXML's `w:type` inside the section's own `w:sectPr`.
 * Meaningless on a document's first section. */
export type SectionStart = "nextPage" | "continuous";

/** Typed pass-through for a section start coming from a `.vue` template's
 * plain string (same espree constraint as asCharacterFormatKey). */
export function asSectionStart(value: string): SectionStart {
  return value as SectionStart;
}

/** Typed pass-through for a break kind coming from a `.vue` template's
 * plain string (same espree constraint as asCharacterFormatKey). */
export function asBreakKind(value: string): BreakKind {
  return value as BreakKind;
}

/** A section's multi-column layout (SPEC-030). `count` is 1..6; a `count`
 * of 1 (or an absent `columns`) is an ordinary single-column section and
 * behaves exactly as before this spec. Widths and the gutter are stored in
 * cm; when `equalWidth` is true (AC-3a) the per-column widths are computed
 * evenly from the section's content width and `widthsCm` is ignored. */
export interface SectionColumns {
  /** Number of columns, 1..6 (AC-2). */
  count: number;
  /** Whether the columns share one computed width (AC-3a). When false,
   * `widthsCm` gives each column's own width. */
  equalWidth: boolean;
  /** The gutter (spacing) between adjacent columns, in cm (AC-3). Uniform
   * across gutters — the dialog and ruler both treat it as one value. */
  gapCm: number;
  /** "Line between columns" (AC-4): a vertical rule drawn in each gutter. */
  separator: boolean;
  /** Per-column widths in cm when `equalWidth` is false (AC-3); one entry
   * per column. Ignored (and normally absent) for equal columns. */
  widthsCm?: number[];
}

/** The inclusive column-count bounds the Columns dialog accepts (AC-2). */
export const MIN_COLUMN_COUNT = 1;
export const MAX_COLUMN_COUNT = 6;

/** Word's default column gutter (0.5", 720 twips ≈ 1.27cm) — the spacing a
 * freshly applied multi-column layout starts with (AC-3). */
export const DEFAULT_COLUMN_GAP_CM = 1.27;

/** The smallest width a section column may be dragged/entered to (AC-14, and
 * the dialog's per-column validation). */
export const MIN_SECTION_COLUMN_WIDTH_CM = 1.27;

/** A section's effective column layout: its own, else a single column. */
export function sectionColumns(model: DocumentModel, sectionIndex: number): SectionColumns {
  return (
    model.sections[sectionIndex]?.columns ?? {
      count: 1,
      equalWidth: true,
      gapCm: DEFAULT_COLUMN_GAP_CM,
      separator: false,
    }
  );
}

/** The content width (cm) available to a section's text: its page width less
 * its horizontal margins. */
export function sectionContentWidthCm(model: DocumentModel, sectionIndex: number): number {
  const pageSize = sectionPageSize(model, sectionIndex);
  const margins = sectionMargins(model, sectionIndex);
  return Math.max(0.1, pageSize.widthCm - margins.leftCm - margins.rightCm);
}

/** The resolved width (cm) of every column in a layout, given the section's
 * content width. Equal columns split the content width (minus gutters)
 * evenly; unequal columns keep their stored proportions rescaled to fill the
 * content width exactly, so the total is always preserved (AC-14) even after
 * a page-size or margin change. */
export function resolveColumnWidthsCm(
  columns: SectionColumns,
  contentWidthCm: number,
): number[] {
  const count = Math.max(1, Math.min(MAX_COLUMN_COUNT, Math.round(columns.count)));
  const gutters = Math.max(0, columns.gapCm) * (count - 1);
  const available = Math.max(0.1, contentWidthCm - gutters);
  if (columns.equalWidth || !columns.widthsCm || columns.widthsCm.length !== count) {
    const each = available / count;
    return Array.from({ length: count }, () => each);
  }
  const total = columns.widthsCm.reduce((sum, w) => sum + Math.max(0, w), 0) || 1;
  return columns.widthsCm.map((w) => (Math.max(0, w) / total) * available);
}

/** The width (cm) of a section's caret column for content sized to fit a
 * single column — the smallest column, so new tables/images fit any column
 * (AC-13, AC-15). For a single-column section this is just the content
 * width. */
export function sectionColumnWidthCm(model: DocumentModel, sectionIndex: number): number {
  const columns = sectionColumns(model, sectionIndex);
  if (columns.count <= 1) {
    return sectionContentWidthCm(model, sectionIndex);
  }
  return Math.min(...resolveColumnWidthsCm(columns, sectionContentWidthCm(model, sectionIndex)));
}

/** Typed pass-through for a SectionColumns arriving from a `.vue` template as
 * a runtime-declared Object (same espree constraint as asParagraphBlock). */
export function asSectionColumns(value: unknown): SectionColumns {
  return value as SectionColumns;
}

/** Block-level children of a section: paragraphs, tables (SPEC-022),
 * images (SPEC-024) and manual breaks (SPEC-028). */
export type Block = Paragraph | Table | Image | BreakBlock;

/** A section groups block-level content; documents have one or more. A
 * section boundary IS the section break (SPEC-028): inserting a section
 * break splits a section in two, deleting one merges neighbours back. */
export interface Section {
  blocks: Block[];
  /** Unmodelled block-level XML preserved through a same-format save
   * (SPEC-008). A side-list (rather than a Block union member) so `blocks`
   * stays renderable/editable without narrowing; order follows
   * `blockIndex`. */
  opaqueBlocks?: OpaqueBlock[];
  /** How this section begins (SPEC-028); absent means `"nextPage"`. */
  start?: SectionStart;
  /** This section's own page size; absent falls back to
   * `defaultStyles.pageSize` (SPEC-028 AC-17/18). */
  pageSize?: PageSize;
  /** This section's own margins; absent falls back to
   * `defaultStyles.margins` (SPEC-028 AC-19). */
  margins?: Margins;
  /** This section's multi-column layout (SPEC-030); absent means a single
   * column. Round-trips as `w:cols` (.docx) / `style:columns` (.odt). */
  columns?: SectionColumns;
  /** Verbatim `w:sectPr` of a non-final .docx section (SPEC-028) — the one
   * inside its last paragraph's `w:pPr` — re-emitted on a same-format save
   * so unmodelled section properties survive (the SPEC-008 fidelity rule;
   * the final section's sectPr lives in `SourceXml.docxSectPr`). Cleared by
   * any command that edits this section's modelled page setup, so the
   * writer regenerates it instead of reproducing stale XML. */
  docxSectPr?: string;
}

/** Page dimensions in centimetres. */
export interface PageSize {
  widthCm: number;
  heightCm: number;
}

/** Page margins in centimetres. */
export interface Margins {
  topCm: number;
  rightCm: number;
  bottomCm: number;
  leftCm: number;
}

/** The unit the Page setup dialog displays sizes/margins in (SPEC-027
 * AC-4, AC-6). The model itself always stores cm — this only controls
 * display/entry — so no reader/writer maps it to anything in a `.docx` or
 * `.odt` file; it is a document-level app preference persisted the same way
 * as `suppressConversionReport`. */
export type MeasurementUnit = "cm" | "mm" | "in" | "pt";

/** Typed pass-through for a measurement unit coming from a `.vue`
 * template's plain string (same espree constraint as asCharacterFormatKey). */
export function asMeasurementUnit(value: string): MeasurementUnit {
  return value as MeasurementUnit;
}

/** Document-wide defaults new content is created with. */
export interface DefaultStyles {
  pageSize: PageSize;
  margins: Margins;
  fontFamily: string;
  fontSizePt: number;
}

/** The source format a document was opened from, if any — drives the format
 * badge in the app bar. Undefined for documents created blank (SPEC-002). */
export type DocumentFormat = "docx" | "odt";

/** A package part the reader does not model (SPEC-008) — `customXml/`,
 * charts, themes, embedded objects, images, `word/webSettings.xml`, … —
 * retained for the lifetime of the document and written back byte-identical
 * on a save to the same format. */
export interface OpaquePart {
  /** Path within the source package, e.g. "word/charts/chart1.xml". */
  name: string;
  /** The part's original (decompressed) bytes, base64-encoded so the model
   * stays JSON-serialisable (see the header comment). */
  data: string;
}

/** Verbatim XML captured from the source package's modelled parts that the
 * model itself cannot express, re-emitted on a same-format save (SPEC-008).
 * Each field is format-specific; all are compared canonically. */
export interface SourceXml {
  /** .docx: the original body-level `w:sectPr`, emitted in place of a
   * generated one so unmodelled section properties (header/footer
   * references, columns, …) survive. Page size/margins cannot yet diverge
   * from `defaultStyles` — no spec edits them. */
  docxSectPr?: string;
  /** .odt: the original `office:font-face-decls` of content.xml. */
  odtFontFaceDecls?: string;
  /** .odt: original automatic-style definitions of content.xml other than
   * paragraph/text `style:style`s (graphic, table, list, … families) —
   * kept because opaque nodes may reference them. Paragraph/text styles are
   * regenerated from the model's resolved formatting instead. */
  odtAutomaticStyles?: string[];
}

/** The root of the document tree. */
export interface DocumentModel {
  id: string;
  title: string;
  sections: Section[];
  defaultStyles: DefaultStyles;
  /** The document's own paragraph-style catalogue (SPEC-019), read from the
   * opened file's styles.xml. Absent for blank documents — the picker then
   * offers the built-in set from lib/styles.ts. */
  styles?: NamedStyle[];
  format?: DocumentFormat;
  /** Image bytes referenced by `Image.mediaId` (SPEC-024), deduplicated —
   * one entry per distinct image regardless of how many times it is used. */
  media?: MediaPart[];
  /** Unmodelled package parts retained from the opened file (SPEC-008). */
  opaque?: OpaquePart[];
  /** Unmodelled XML retained from the opened file's parsed parts. */
  sourceXml?: SourceXml;
  /** True once the user ticked "Don't show again for this document" on the
   * conversion report dialog (SPEC-008); persisted with the document. */
  suppressConversionReport?: boolean;
  /** The Page setup dialog's measurement-unit preference (SPEC-027 AC-6,
   * AC-7); absent means "cm", the product default. */
  measurementUnit?: MeasurementUnit;
  /** Whether the horizontal ruler is shown (SPEC-027 AC-14); absent means
   * shown — the ruler predates this toggle (SPEC-014) and was always
   * visible, so an unset field must keep reading as visible. */
  showRuler?: boolean;
  /** Whether formatting marks (¶, spaces, tabs, break markers) are shown
   * (SPEC-028 AC-12..16); absent means hidden, the product default. A
   * document-level app preference persisted like `showRuler` — neither
   * reader/writer maps it to anything in a `.docx`/`.odt`. */
  showFormattingMarks?: boolean;
}

export const DEFAULT_TITLE = "Untitled document";

const A4_PAGE_SIZE: PageSize = { widthCm: 21, heightCm: 29.7 };

const STANDARD_MARGINS: Margins = {
  topCm: 2.54,
  rightCm: 2.54,
  bottomCm: 2.54,
  leftCm: 2.54,
};

// Calibri 11pt is Word's modern default body font; picked as the product
// default since OOXML (.docx) is this project's primary format.
const DEFAULT_FONT_FAMILY = "Calibri";
const DEFAULT_FONT_SIZE_PT = 11;

/** The product's built-in fallback styles (A4, standard margins, Calibri
 * 11pt) — the starting point for a blank document, and the fallback the
 * `.docx` reader falls back to for anything a source file doesn't specify. */
export function createDefaultStyles(): DefaultStyles {
  return {
    pageSize: { ...A4_PAGE_SIZE },
    margins: { ...STANDARD_MARGINS },
    fontFamily: DEFAULT_FONT_FAMILY,
    fontSizePt: DEFAULT_FONT_SIZE_PT,
  };
}

/** A section's effective page size: its own, else the document default
 * (SPEC-028 AC-17/18). */
export function sectionPageSize(model: DocumentModel, sectionIndex: number): PageSize {
  return model.sections[sectionIndex]?.pageSize ?? model.defaultStyles.pageSize;
}

/** A section's effective margins: its own, else the document default
 * (SPEC-028 AC-19). */
export function sectionMargins(model: DocumentModel, sectionIndex: number): Margins {
  return model.sections[sectionIndex]?.margins ?? model.defaultStyles.margins;
}

function createEmptyParagraph(): Paragraph {
  return { type: "paragraph", runs: [] };
}

/** Creates a new blank document: one section, one empty paragraph, and the
 * product's default styles. Pass an id to control it (e.g. a route param);
 * otherwise one is generated. */
export function createDocument(id: string = crypto.randomUUID()): DocumentModel {
  return {
    id,
    title: DEFAULT_TITLE,
    sections: [{ blocks: [createEmptyParagraph()] }],
    defaultStyles: createDefaultStyles(),
  };
}
