// The in-memory document model: a serialisable tree every later spec reads
// and writes. Plain data only (no classes, no functions on the objects) so
// `JSON.stringify`/`JSON.parse` round-trips it without loss — that property
// is what SPEC-006/007 (save) and SPEC-004/005 (open) will rely on.
import { DEFAULT_LANGUAGE } from "./language";

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
  /** This run's language (SPEC-035), e.g. "en-US" — see lib/language.ts.
   * Absent means the document's default language (`DefaultStyles.language`)
   * applies; explicitly set here only when a run/style diverges from it, so
   * opening a document does not stamp every run with the document default
   * (see docx/styles.ts's and odf/styles.ts's docDefault "lift-out", which
   * keeps this field absent on the common case). Drives dictionary
   * selection, the rendered `lang` attribute for assistive technology, and
   * round-trips as `w:lang` (.docx) / `fo:language`+`fo:country` (.odt). */
  lang?: string;
  /** "Do not check spelling" (SPEC-035 AC-11): when true, the spell checker
   * skips this run entirely regardless of language/dictionary availability.
   * Round-trips as `w:noProof` (.docx); .odt has no equivalent flag, so it
   * maps to the special `fo:language="zxx"` ("no linguistic content", see
   * lib/language.ts's ODF_NO_LINGUISTIC_CONTENT) convention instead. */
  noProof?: boolean;
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
  /** SPEC-035's language selector, applied to a real selection via
   * applyFormat. */
  lang?: string;
  /** SPEC-035's "Do not check spelling"; `false` clears it (same
   * delete-the-key semantics as the boolean character formats — see
   * withFormat), not an explicit off. */
  noProof?: boolean;
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

/** The page-number field kinds Insert → Page number offers (SPEC-029
 * AC-14, AC-19): "page" is the current page's number, "numpages" the
 * document's total page count. "Page X of Y" is two field runs (one of
 * each kind) either side of a literal " of " text run, not a composite
 * field. Round-trips as `PAGE`/`NUMPAGES` fields (.docx) and
 * `text:page-number`/`text:page-count` (.odt). */
export type FieldKind = "page" | "numpages";

/** Mail-merge chip kinds (SPEC-046): `merge` substitutes a named field's
 * value, `if`/`endif` delimit a conditional block, `each` marks a table row
 * that repeats over an array field. See "Merge chips in the model" in
 * specs/046-template-gallery-and-mail-merge.md. */
export type MergeChipKind = "merge" | "if" | "endif" | "each";

/** A mail-merge chip's spec (SPEC-046). `name` holds the field or array
 * name — deliberately not `id`: lib/equality.ts's IGNORED_KEYS drops bare
 * `id` keys from structural comparison, which would let a chip pointing at
 * the wrong field pass a round-trip test (the same reasoning as
 * BookmarkAnchor). `endif` carries no name — it only closes the nearest
 * open `if`. */
export type MergeFieldSpec =
  | { kind: "merge"; name: string }
  | { kind: "if"; name: string }
  | { kind: "endif" }
  | { kind: "each"; name: string };

/** The numbering formats Insert → Page number offers (SPEC-029 AC-16..18)
 * — deliberately the same vocabulary as lib/lists.ts's `ListMarkerStyle`
 * (its `formatOrdinal` is reused to render a field's value), just narrowed
 * to the three the dialog actually offers. */
export type PageNumberFormat = "decimal" | "lowerRoman" | "lowerAlpha";

/** Typed pass-through for a page-number format coming from a `.vue`
 * template's plain string (same espree constraint as asCharacterFormatKey). */
export function asPageNumberFormat(value: string): PageNumberFormat {
  return value as PageNumberFormat;
}

/** Typed pass-through for a `DocumentModel` passed into a `.vue` component
 * as a runtime `Object` prop (espree bars writing `PropType<DocumentModel>`
 * — see CommentsPane.vue's header comment for the same constraint). */
export function asDocumentModel(value: unknown): DocumentModel {
  return value as DocumentModel;
}

/** A run that renders as a computed value instead of literal text
 * (SPEC-029, widened by SPEC-046): a page-number/page-count field, or a
 * mail-merge chip. `Run.text` still holds a one-character placeholder so
 * the field behaves like a normal single-character atom for
 * caret/selection arithmetic (blockLength, splitRuns, …) — every renderer
 * substitutes the live value at paint time instead of reading it. */
export type FieldSpec = { kind: "page" | "numpages"; format: PageNumberFormat } | MergeFieldSpec;

/** Narrows a `FieldSpec` to a merge chip (SPEC-046), as opposed to a
 * SPEC-029 page-number/page-count field. */
export function isMergeChipField(field: FieldSpec): field is MergeFieldSpec {
  return field.kind === "merge" || field.kind === "if" || field.kind === "endif" || field.kind === "each";
}

/** Which note stream a reference belongs to (SPEC-031): footnotes render at
 * the bottom of the page holding their reference, endnotes in one area at
 * the end of the document. */
export type NoteKind = "footnote" | "endnote";

/** An inline footnote/endnote reference mark (SPEC-031) carried on a run —
 * the same convention as FieldSpec: the run's text stays a one-character
 * placeholder so the mark behaves as a single caret atom, and every
 * renderer substitutes the computed note number at paint time. `noteId`
 * references a `DocumentModel.notes` entry; deliberately not named `id`
 * (see BookmarkAnchor for the equality.ts IGNORED_KEYS rule). */
export interface NoteRef {
  noteId: string;
  kind: NoteKind;
}

/** The smallest unit of text: a string plus the formatting applied to it.
 * Tabs and forced line breaks are preserved as literal "\t" / "\n"
 * characters in `text` — renderers use `white-space: pre-wrap` so both
 * display exactly as the source document intended. */
export interface Run {
  text: string;
  formatting: RunFormatting;
  /** Marks this run as a computed page-number/page-count field (SPEC-029)
   * rather than literal text; see FieldSpec. */
  field?: FieldSpec;
  /** Marks this run as a footnote/endnote reference mark (SPEC-031); see
   * NoteRef. Mutually exclusive with `field` by construction. */
  note?: NoteRef;
  /** Character-style reference (SPEC-042), independent of the containing
   * paragraph's `styleId`: `w:rStyle` (.docx) / a `text:span`'s
   * `text:style-name` resolving to a family="text" style (.odt). Baked
   * formatting still carries the resolved effect, the same "reference kept
   * alongside the resolved formatting" convention as `Paragraph.styleId`. */
  styleId?: string;
  /** Pending tracked-change state (SPEC-039); see RunTrackedChanges. */
  changes?: RunTrackedChanges;
}

/** A field run's placeholder text — a single character so it occupies
 * exactly one caret position, like Image/BreakBlock's length-1 atom
 * convention, while still sitting inline mid-paragraph. */
export const FIELD_PLACEHOLDER_TEXT = "1";

/** Builds a page-number/page-count field run (SPEC-029): the "Insert page
 * number" dialog's per-format insertion and the .docx/.odt readers' PAGE /
 * NUMPAGES / text:page-number / text:page-count parsing both create runs
 * this way, so every field run in the model has the same placeholder text. */
export function createFieldRun(
  kind: FieldKind,
  format: PageNumberFormat,
  formatting: RunFormatting = {},
): Run {
  return { text: FIELD_PLACEHOLDER_TEXT, formatting, field: { kind, format } };
}

/** Builds a mail-merge chip run (SPEC-046) — the same one-character
 * placeholder convention as createFieldRun, so a chip is atomic for
 * caret/selection arithmetic. */
export function createChipRun(field: MergeFieldSpec, formatting: RunFormatting = {}): Run {
  return { text: FIELD_PLACEHOLDER_TEXT, formatting, field };
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
  /** Custom tab stops (SPEC-032 AC-19..25), ordered by position. Absent
   * means the default tab grid only. Round-trips as `w:tabs` (.docx) and
   * `style:tab-stops` (.odt). */
  tabStops?: TabStop[];
}

// --- lists (SPEC-018; multilevel definitions SPEC-032) ----------------------

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
  /** References a `DocumentModel.listDefinitions` entry (SPEC-032).
   * Presence makes this a *defined* list item: its marker text, per-level
   * styles and indents all come from the definition, and its numbering is
   * counted per (defId, listId) instance rather than by adjacency.
   * Deliberately not named `id` (see BookmarkAnchor for the equality.ts
   * IGNORED_KEYS rule). */
  defId?: string;
  /** The numbering instance this item counts in (SPEC-032): items sharing
   * (defId, listId) form one sequence that continues across interruptions;
   * a fresh listId is what "Restart numbering" mints. Meaningless without
   * `defId`. */
  listId?: string;
}

/** How a defined level's number is aligned within its marker box (SPEC-032
 * AC "alignment"; `w:lvlJc` / `fo:text-align` on the level). */
export type ListLevelAlignment = "left" | "center" | "right";

/** One level (1..9) of a multilevel list definition (SPEC-032). Every field
 * except `style` is optional with a documented default, so a definition
 * only stores what differs — which keeps save/re-open fixed points exact. */
export interface ListLevelDefinition {
  /** The level's marker style: a number format or a bullet glyph. */
  style: ListMarkerStyle;
  /** Literal text before the number (e.g. "("); absent means "". */
  prefix?: string;
  /** Literal text after the number (e.g. ")" or "."); absent means "". */
  suffix?: string;
  /** How many level components the number shows, ending at this level
   * (AC-8's "1.1" vs "1"): 3 at level 3 renders "1.1.1"; absent means 1. */
  includeLevels?: number;
  /** The level's start value; absent means 1. */
  startAt?: number;
  /** Number alignment; absent means "left". */
  alignment?: ListLevelAlignment;
  /** The marker's position (indent) in cm; absent means the editor default
   * (text position − 0.635 cm, see lib/listDefinitions.ts). */
  numberIndentCm?: number;
  /** The item text's position in cm; absent means level × 1.27 cm. */
  textIndentCm?: number;
}

/** A multilevel list definition (SPEC-032): nine levels of numbering
 * format, referenced by `ListFormatting.defId`. Each list owns its own
 * definition (gallery presets are copied on apply), which is what makes
 * editing a level's format update that list only (AC-17, AC-18). */
export interface ListDefinition {
  /** Identity referenced by `ListFormatting.defId` — not named `id` for
   * the same equality.ts reason as `BookmarkAnchor.name`. */
  defId: string;
  /** Exactly MAX_LIST_LEVEL entries, level 1 first. */
  levels: ListLevelDefinition[];
  /** Paragraph styles linked to this definition's levels (AC-13):
   * `styleId -> level`. Applying a linked style makes the paragraph a
   * defined item of this definition at that level automatically. */
  styleLinks?: Record<string, number>;
}

/** Typed pass-through for a `ListLevelDefinition[]` arriving from a `.vue`
 * component's untyped Array prop (same espree constraint as
 * asCharacterFormatKey). */
export function asListLevelDefinitions(value: unknown): ListLevelDefinition[] {
  return value as ListLevelDefinition[];
}

/** Typed pass-through for a `ListDefinition[]` prop (espree constraint). */
export function asListDefinitions(value: unknown): ListDefinition[] {
  return value as ListDefinition[];
}

/** Typed empty seed for list-definition arrays in `.vue` scripts. */
export function noListDefinitions(): ListDefinition[] {
  return [];
}

// --- tab stops (SPEC-032) ---------------------------------------------------

/** A custom tab stop's alignment (AC-19..22): where the tabbed text sits
 * relative to the stop — "decimal" aligns on the decimal separator. */
export type TabStopAlignment = "left" | "right" | "center" | "decimal";

/** The leader drawn in the space a tab jumps over (AC-23); absent means
 * none. */
export type TabLeader = "dot" | "hyphen" | "underscore";

/** One custom tab stop on a paragraph (SPEC-032 AC-19..25). `positionCm`
 * is measured from the text area's start edge (the left margin for LTR),
 * matching `w:tab/@w:pos` and ODF's `style:position`. */
export interface TabStop {
  positionCm: number;
  alignment: TabStopAlignment;
  leader?: TabLeader;
}

/** Typed pass-through for a `TabStop[]` arriving from a `.vue` component's
 * untyped Array prop (same espree constraint as asCharacterFormatKey). */
export function asTabStops(value: unknown): TabStop[] {
  return value as TabStop[];
}

/** Typed empty seed for tab-stop arrays in `.vue` scripts. */
export function noTabStops(): TabStop[] {
  return [];
}

/** Typed pass-through for a tab alignment coming from a `.vue` template's
 * plain string (same espree constraint as asCharacterFormatKey). */
export function asTabStopAlignment(value: string): TabStopAlignment {
  return value as TabStopAlignment;
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

/** A style's family (SPEC-042): which kind of thing it applies to and how
 * it groups in the Styles pane. Absent on `NamedStyle.type` means
 * `"paragraph"` — the only family that existed before SPEC-042, kept as the
 * unmarked default so every style built by SPEC-019 code stays exactly as
 * it was. */
export type StyleType = "paragraph" | "character" | "list" | "table";

/** One style of the document's style catalogue (SPEC-019 paragraph styles;
 * SPEC-042 widens the catalogue to character/list/table styles too): its
 * *own* (non-inherited) formatting plus the parent link resolution follows
 * (`w:basedOn` / `style:parent-style-name`). For a `.docx` these mirror the
 * `w:style` elements of styles.xml; for a `.odt`, the named styles of its
 * styles.xml. The model keeps paragraph/run formatting *baked* (effective)
 * on each paragraph/run, so the catalogue is not needed to render — it
 * drives the style picker/pane, the diff that separates direct formatting
 * from style formatting, and the writers' style definitions. */
export interface NamedStyle {
  /** Style identity as referenced by `Paragraph.styleId`/`Run.styleId`. */
  id: string;
  /** Display name shown in the picker/pane. */
  name?: string;
  /** The style's family; absent means `"paragraph"`. */
  type?: StyleType;
  /** Parent style id the style inherits from. */
  basedOn?: string;
  /** The style id applied to a new paragraph typed after one in this style
   * (SPEC-042 AC-3/AC-4's "style for following paragraph") — `w:next` /
   * `style:next-style-name`. Meaningful for paragraph/list styles only. */
  next?: string;
  /** The style's own paragraph formatting; absent when it sets none. */
  paragraphFormatting?: ParagraphFormatting;
  /** The style's own run formatting; absent when it sets none. */
  runFormatting?: RunFormatting;
  /** ODF-only: a proportional font size ("130%" in the source), resolved
   * against the inherited size during chain resolution. */
  fontSizePercent?: number;
  /** 1-based outline level; presence makes this a heading style. */
  headingLevel?: number;
  /** List-type styles only (SPEC-042): the `DocumentModel.listDefinitions`
   * entry this style's numbering comes from — applying the style links a
   * paragraph to it the same way a heading-numbering link does (see
   * commands.ts's styleLinkFor). */
  listDefinitionId?: string;
  /** Table-type styles only (SPEC-042): default cell borders a table using
   * this style bakes onto cells that carry no border of their own. Table
   * banding (per-region overrides) is out of scope — a single border/shading
   * set applies uniformly. */
  cellBorders?: CellBorders;
  /** Table-type styles only (SPEC-042): default cell background a table
   * using this style bakes onto cells that carry no shading of their own. */
  cellShadingFill?: string;
  /** True for a style read from a source document the editor cannot fully
   * model (SPEC-042 AC-15..18) — a real table style with banding regions, a
   * `w:type="numbering"` entry, or any style carrying structure the reader
   * does not parse. Still listed in the pane and can be applied by
   * reference (`rawXml` is what the writer re-emits verbatim); `Modify` is
   * unavailable for it since the editor cannot express its own properties. */
  unmodeled?: boolean;
  /** The style's original source XML (a `<w:style>`/`<style:style>` element,
   * serialised), preserved verbatim for `unmodeled` styles so export writes
   * it back unchanged (AC-18). Absent for styles the editor fully models. */
  rawXml?: string;
}

/** Typed empty seed for `NamedStyle[]` refs/props in `.vue` scripts (espree
 * bars the type annotation — see asCharacterFormatKey). */
export function noNamedStyles(): NamedStyle[] {
  return [];
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

/** Which edge of a comment thread's anchored range a `CommentAnchor` marks
 * (SPEC-038) — mirrors OOXML's split `w:commentRangeStart`/
 * `w:commentRangeEnd` tags and ODF's `office:annotation`/
 * `office:annotation-end`. */
export type CommentAnchorBoundary = "start" | "end";

/** One edge of a comment thread's anchored range (SPEC-038) — a
 * character-offset point stored as a paragraph-level side-list, the same
 * shape as `BookmarkAnchor`. A thread's `"start"` and `"end"` points are
 * matched by `threadId`, not by array position or paragraph: the common
 * case has both in the same paragraph, but a range spanning a paragraph
 * break puts them in different paragraphs. Offset maintenance across edits
 * (insert/delete/split/merge) follows the exact same rules as
 * `BookmarkAnchor` — see textEditing.ts's `anchorsAfterInsert`/
 * `anchorsAfterDelete`/`mergedAnchors`, shared by both. Deliberately not
 * named `id`: lib/equality.ts's IGNORED_KEYS drops bare "id" keys from
 * structural comparison (see BookmarkAnchor for the same reasoning). */
export interface CommentAnchor {
  threadId: string;
  boundary: CommentAnchorBoundary;
  offset: number;
}

/** One message in a comment thread (SPEC-038): the thread's opening comment
 * (`CommentThread.messages[0]`) or a reply. Plain text only — rich text
 * inside comments is out of scope — so `text` is rendered as plain text
 * everywhere (never interpreted as HTML/markup, AC "no markup injection").
 * Deliberately not named `id` — same equality.ts reason as BookmarkAnchor. */
export interface CommentMessage {
  messageId: string;
  /** The commenter's display name, from the settable identity
   * (lib/commentIdentity.ts) at the time the message was written. */
  author: string;
  text: string;
  /** ISO 8601 timestamp. */
  createdAt: string;
}

/** A comment thread anchored to a range of document text (SPEC-038): the
 * in-memory, format-agnostic shape of OOXML's
 * `w:commentRangeStart/End`+`w:commentReference` and ODF's
 * `office:annotation`+`office:annotation-end`. `messages[0]` is the
 * thread's opening comment; `messages[1..]` are replies. A thread stays in
 * `DocumentModel.comments` even after its anchoring range is deleted from
 * the body — visibility (which highlight to paint, whether the thread
 * shows up anywhere) is always DERIVED from live `CommentAnchor`s
 * (lib/comments.ts's `visibleCommentThreads`), never cascade-deleted — the
 * same contract `NoteContent` uses (SPEC-031), which is what lets undo (and
 * a cut-and-pasted-back range) find the thread's content again. Not named
 * `id` — same equality.ts reason as BookmarkAnchor. */
export interface CommentThread {
  threadId: string;
  messages: CommentMessage[];
  /** Whether the thread is resolved (SPEC-038): resolving hides its
   * highlight and moves it to the "Resolved" section of the comments UI; it
   * can be reopened. Never affects the anchor itself. */
  resolved: boolean;
}

// --- track changes (SPEC-039) -----------------------------------------------

/** One tracked edit's author and timestamp — the common shape every kind of
 * tracked change (insertion, deletion, formatting change) carries. Not
 * named `id` (same equality.ts IGNORED_KEYS reason as BookmarkAnchor);
 * `changeId` mirrors the threadId/messageId/noteId convention. */
export interface TrackedChangeInfo {
  changeId: string;
  /** The settable identity (lib/commentIdentity.ts) at the time of the edit. */
  author: string;
  /** ISO 8601 timestamp. */
  date: string;
}

/** A run's pending formatting change (SPEC-039, `w:rPrChange` / ODF
 * `text:format-change`): the formatting the run had before the change,
 * restored verbatim if the change is rejected (AC "rejecting... restores
 * the exact original text and formatting"). */
export interface RunFormatChange extends TrackedChangeInfo {
  previous: RunFormatting;
}

/** Tracked-change state carried on a single run (SPEC-039). A run can be
 * both inserted and deleted at once: editing over another author's
 * still-pending insertion nests the deletion on top of it instead of
 * losing the insertion outright (AC "produces a nested change, not a lost
 * one") — accepting the deletion then removes the run, rejecting it drops
 * back to the plain pending insertion. */
export interface RunTrackedChanges {
  insert?: TrackedChangeInfo;
  delete?: TrackedChangeInfo;
  formatChange?: RunFormatChange;
}

/** A paragraph's pending formatting change (SPEC-039, `w:pPrChange`). */
export interface ParagraphFormatChange extends TrackedChangeInfo {
  previous: ParagraphFormatting | undefined;
}

/** Tracked-change state carried on a paragraph (SPEC-039): its own trailing
 * mark (the paragraph break) and/or its block-level formatting. A split
 * (Enter) sets `markInsert` on the paragraph before the new break; a merge
 * (Backspace at a paragraph start) sets `markDelete` on the paragraph being
 * folded into its successor rather than physically merging it, so rejecting
 * the deletion simply clears the marker and the paragraphs stay apart. */
export interface ParagraphTrackedChanges {
  markInsert?: TrackedChangeInfo;
  markDelete?: TrackedChangeInfo;
  formatChange?: ParagraphFormatChange;
}

/** Tracked insertion/deletion of a whole block (SPEC-039): a table row or an
 * image added or removed while tracking is on (AC "changes made by list,
 * table and image operations are tracked"). Unlike a run/paragraph, a
 * deleted block has nothing to render struck through in place, so the
 * review pane and margin change bar are its only visible trace until
 * accepted/rejected. */
export interface BlockTrackedChanges {
  insert?: TrackedChangeInfo;
  delete?: TrackedChangeInfo;
}

/** A table cell's pending property edit (SPEC-039 AC-15): a merge/split
 * (colSpan/rowSpan changed) or a border/shading/padding/vertical-alignment
 * edit — `w:tcPrChange`/`w:cellMerge` on `.docx`. Restored verbatim if
 * rejected. `.odt` has no standard slot for "previous" cell-property state
 * (the same documented gap as RunFormatChange on a `.odt` source), so an
 * `.odt`-sourced cell property change is accepted (applied) and reported as
 * a SPEC-008 DropRecord on a cross-format save to `.odt` (AC-19) rather
 * than tracked there. */
export interface CellPropertyChange extends TrackedChangeInfo {
  previous: {
    colSpan?: number;
    rowSpan?: number;
    borders?: CellBorders;
    shadingFill?: string;
    paddingCm?: number;
    verticalAlign?: CellVerticalAlign;
  };
}

/** Tracked-change state carried on a table cell (SPEC-039 AC-15): insertion/
 * deletion (a column added/removed while tracking is on) alongside
 * `BlockTrackedChanges`, plus a pending property change/merge. */
export interface CellTrackedChanges extends BlockTrackedChanges {
  propertyChange?: CellPropertyChange;
}

/** A table row's pending property edit (SPEC-039 AC-15): a height change —
 * `w:trPrChange` on `.docx`. See CellPropertyChange for the `.odt`
 * documented gap this shares. */
export interface RowPropertyChange extends TrackedChangeInfo {
  previous: { minHeightCm?: number };
}

/** Tracked-change state carried on a table row (SPEC-039): insertion/
 * deletion alongside `BlockTrackedChanges`, plus a pending row-property
 * change. */
export interface RowTrackedChanges extends BlockTrackedChanges {
  propertyChange?: RowPropertyChange;
}

/** A table's pending property edit (SPEC-039 AC-15): column widths,
 * alignment or width-mode changed — `w:tblPrChange`/`w:tblGridChange` on
 * `.docx`. See CellPropertyChange for the `.odt` documented gap this
 * shares. */
export interface TablePropertyChange extends TrackedChangeInfo {
  previous: {
    columnWidthsCm?: number[];
    alignment?: TableAlignment;
    widthMode?: TableWidthMode;
    widthPercent?: number;
  };
}

/** Tracked-change state carried on a table (SPEC-039): just its own pending
 * property change — a table itself is never inserted/deleted as a unit
 * (only its rows/cells are). */
export interface TableTrackedChanges {
  propertyChange?: TablePropertyChange;
}

/** A fresh, session-unique tracked-change id. */
export function createChangeId(): string {
  return `change-${crypto.randomUUID()}`;
}

/** The review pane's display-mode selector (SPEC-039 AC): "All markup"
 * shows every insertion/deletion/formatting mark; "Simple markup" shows
 * only the margin change bars; "No markup" renders as if every change were
 * accepted; "Original" renders as if every change were rejected. Documents
 * always open in `"all"`. */
export type TrackChangesDisplayMode = "all" | "simple" | "noMarkup" | "original";

/** Typed pass-through for a display mode coming from a `.vue` template's
 * plain string (same espree constraint as asCharacterFormatKey). */
export function asTrackChangesDisplayMode(value: string): TrackChangesDisplayMode {
  return value as TrackChangesDisplayMode;
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
  /** Comment anchor points in this paragraph (SPEC-038), ordered by
   * `offset`. See `CommentAnchor`; offsets are maintained by the
   * textEditing primitives the same way `bookmarks` are. */
  commentAnchors?: CommentAnchor[];
  /** Pending tracked-change state for this paragraph's own mark/formatting
   * (SPEC-039); see ParagraphTrackedChanges. */
  changes?: ParagraphTrackedChanges;
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
  /** Pending tracked insertion/deletion of this image block (SPEC-039). */
  changes?: BlockTrackedChanges;
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
  /** Pending tracked insertion/deletion/property-change of this cell
   * (SPEC-039 AC-15: a column insert/delete, a merge/split, or a border/
   * shading/padding/vertical-alignment edit made while tracking is on). */
  changes?: CellTrackedChanges;
}

/** One table row. */
export interface TableRow {
  cells: TableCell[];
  /** The row's minimum height in cm (SPEC-023 AC-4, AC-5): dragging a row
   * boundary sets this floor, but the row still grows taller to fit its
   * content — it is never a fixed height. */
  minHeightCm?: number;
  /** Whether this row repeats as the table's header row (SPEC-043's "table
   * without a header row" check): `w:tblHeader` (.docx) — a per-row flag —
   * or membership in a leading `table:table-header-rows` wrapper (.odt),
   * which only supports one contiguous leading run of header rows. Absent
   * means an ordinary body row. */
  isHeader?: boolean;
  /** Verbatim `w:trPr` from a .docx source (row height, header flag, …). */
  docxTrPr?: string;
  /** `table:style-name` of the source .odt row. */
  odtStyleName?: string;
  /** Pending tracked insertion/deletion/property-change of this row
   * (SPEC-039 AC "table row/cell insert/delete", AC-15's row property
   * changes). */
  changes?: RowTrackedChanges;
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
  /** A custom table style applied from the Styles pane (SPEC-042), by
   * reference into `DocumentModel.styles`. Distinct from `odtStyleName`/
   * `docxTblPr`, which hold a source document's own (possibly unmodelled)
   * table styling — applying a catalogue style bakes its borders/shading
   * onto the table's cells and records the reference here for in-use
   * tracking and re-bake on Modify. */
  styleId?: string;
  /** Per-column `table:style-name`s of the source .odt table's
   * `table:table-column` elements; entries may be undefined for columns
   * inserted in the editor (the writer generates a style for those). */
  odtColumnStyleNames?: Array<string | undefined>;
  /** Pending tracked property change of this table (SPEC-039 AC-15: column
   * widths, alignment or width-mode edited while tracking is on). */
  changes?: TableTrackedChanges;
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

/** Which header/footer variant a section stores content for (SPEC-029
 * AC-11..13): `"default"` applies to every page unless overridden;
 * `"first"` applies to the section's first page (only consulted when
 * `Section.titlePage` is set) and `"even"` to even-numbered pages (only
 * consulted when `Section.evenAndOddHeaders` is set). Mirrors OOXML's
 * `w:headerReference`/`w:footerReference` `w:type` values; ODF represents
 * the same split as separate master-page header/footer elements. */
export type HeaderFooterKind = "default" | "first" | "even";

/** Typed pass-through for a header/footer kind coming from a `.vue`
 * template's plain string (same espree constraint as asCharacterFormatKey). */
export function asHeaderFooterKind(value: string): HeaderFooterKind {
  return value as HeaderFooterKind;
}

/** Which band a header/footer command addresses (SPEC-029). */
export type HeaderFooterPart = "header" | "footer";

/** Typed pass-through for a header/footer part coming from a `.vue`
 * template's plain string (same espree constraint as asCharacterFormatKey). */
export function asHeaderFooterPart(value: string): HeaderFooterPart {
  return value as HeaderFooterPart;
}

/** One header or footer's content: a block flow like a section's own body
 * (SPEC-029 AC-3..6 — text, formatting, tables and images all reuse the
 * same `Block` union) — minus manual breaks, since header/footer content
 * never paginates. Absent for a kind that has never been edited: distinct
 * from an explicit empty content the user cleared, and from "this kind
 * does not apply" (see resolveHeaderFooterKind in lib/pageNumbering.ts) —
 * a section with `titlePage` set but no `headers.first` entry shows a
 * genuinely blank first-page header, matching Word's own behaviour. */
export interface HeaderFooterContent {
  blocks: Block[];
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

// --- table of contents (SPEC-033) -------------------------------------------

/** One heading included in a table of contents: the entry text, outline
 * level (1-6) and the heading's derived anchor (lib/links.ts's
 * headingAnchor) an entry click jumps to — the same anchor scheme internal
 * hyperlinks use (SPEC-025), so a TOC entry is a pre-resolved internal link.
 * `pageNumber` is the page the heading fell on when the TOC was last
 * generated or updated; null until pagination has resolved it once. Staying
 * a plain snapshot (rather than a live lookup) is what makes "headings
 * changed" and "pagination changed" two independently detectable staleness
 * conditions (SPEC-033 AC-10/11, see lib/toc.ts). */
export interface TocEntry {
  text: string;
  level: number;
  anchor: string;
  pageNumber: number | null;
}

/** A table of contents's configuration (SPEC-033 AC-5..7): which heading
 * levels (1-6) it includes, whether it shows page numbers, and whether
 * entries render as plain dot-leader-and-page-number rows or as hyperlinks
 * only — Word's "Use hyperlinks instead of page numbers": no dot leader, no
 * page number, the heading text itself is the clickable link. Every entry
 * is clickable either way (AC-8); this only decides what else is shown. */
export interface TocOptions {
  /** Heading levels (1-6) to include, e.g. [1, 2, 3]. */
  levels: number[];
  showPageNumbers: boolean;
  hyperlinksOnly: boolean;
}

/** Typed pass-through for a TocOptions arriving from a `.vue` template as a
 * runtime-declared Object (same espree constraint as asParagraphBlock). */
export function asTocOptions(value: unknown): TocOptions {
  return value as TocOptions;
}

/** A table of contents block (SPEC-033): a generated, block-level atom —
 * blockLength 1, like Image/BreakBlock (see textEditing.blockLength), a
 * caret sits before or after it, never inside. Built from the document's
 * Heading 1-6 paragraphs (lib/links.ts's listHeadings) and not directly
 * editable; its `entries` are only ever replaced wholesale, by the insert
 * command or the "Update table of contents" commands (lib/toc.ts), never
 * typed into. Round-trips as a `TOC` field with a cached result (.docx) and
 * `text:table-of-content` (.odt). */
export interface TableOfContents {
  type: "toc";
  options: TocOptions;
  entries: TocEntry[];
}

/** Whether a block is a table of contents — the isTable-style narrowing
 * helper (also usable from espree-parsed `.vue` scripts). */
export function isToc(block: Block | undefined): block is TableOfContents {
  return block?.type === "toc";
}

/** Typed pass-through for a TableOfContents arriving as a runtime-declared
 * Object prop (espree bars the type annotation — see asImage for the same
 * pattern). */
export function asToc(value: unknown): TableOfContents {
  return value as TableOfContents;
}

/** Block-level children of a section: paragraphs, tables (SPEC-022),
 * images (SPEC-024), manual breaks (SPEC-028) and tables of contents
 * (SPEC-033). */
export type Block = Paragraph | Table | Image | BreakBlock | TableOfContents;

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
  /** This section's header content, by kind (SPEC-029). */
  headers?: Partial<Record<HeaderFooterKind, HeaderFooterContent>>;
  /** This section's footer content, by kind (SPEC-029). */
  footers?: Partial<Record<HeaderFooterKind, HeaderFooterContent>>;
  /** "Different first page" (SPEC-029 AC-11, AC-13): when set, the
   * section's first page uses `headers.first`/`footers.first` instead of
   * `headers.default`/`footers.default`. Round-trips as `w:titlePg`
   * (.docx) — a separate master page referenced by the section's first
   * paragraph (.odt). */
  titlePage?: boolean;
  /** "Different odd and even pages" (SPEC-029 AC-12, AC-13): when set,
   * even-numbered pages (other than a title page) use
   * `headers.even`/`footers.even` instead of the default. OOXML only
   * models this as the document-wide `w:evenAndOddHeaders` setting; the
   * writer ORs every section's flag together into that one setting and
   * every reader section reads back the same document-wide value (see
   * SPEC-029's docx writer/reader). */
  evenAndOddHeaders?: boolean;
  /** "Start numbering at" (SPEC-029 AC-22, AC-23): the displayed page
   * number of this section's first page. Presence restarts numbering at
   * this value; absence continues counting from the previous section (the
   * document's first section always continues from 1 when absent).
   * Mirrors `w:pgNumType/@w:start` (.docx) and an explicit
   * `style:page-number` on the section's first paragraph (.odt). */
  pageNumberStart?: number;
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
  /** The document's default language (SPEC-035), e.g. "en-US" — see
   * lib/language.ts. Absent means the product default (DEFAULT_LANGUAGE);
   * a run's own `RunFormatting.lang` overrides this for its own text. */
  language?: string;
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

// --- document properties and metadata (SPEC-041) ----------------------------

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

/** The types the Custom Properties tab's add-property form offers, in menu
 * order. */
export const CUSTOM_PROPERTY_TYPES: readonly CustomPropertyType[] = ["text", "number", "date", "boolean"];

/** Typed pass-through for a custom-property type coming from a `.vue`
 * template's plain string (same espree constraint as asCharacterFormatKey). */
export function asCustomPropertyType(value: string): CustomPropertyType {
  return value as CustomPropertyType;
}

/** `DocumentProperties` with `customProperties` guaranteed present — what
 * PropertiesDialog.vue's local draft actually needs (it always pushes/
 * splices that array), so its script never has to narrow an optional
 * property itself (the espree constraint rules out both a `!` assertion and
 * an inline non-optional type annotation there). */
export interface DocumentPropertiesDraft extends DocumentProperties {
  customProperties: CustomProperty[];
}

/** A plain-JSON deep clone of `properties` (or a fresh empty draft when
 * absent) — PropertiesDialog.vue's local editable draft seed. Declared here,
 * not the `.vue` file, so vue-tsc infers its return type without a `.vue`
 * script needing the TS-only generic/cast syntax espree can't parse (same
 * convention as noActiveNote/noBlocks above). */
export function cloneDocumentProperties(properties?: DocumentProperties): DocumentPropertiesDraft {
  const cloned = properties ? (JSON.parse(JSON.stringify(properties)) as DocumentProperties) : {};
  return { ...cloned, customProperties: cloned.customProperties ?? [] };
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
  /** Multilevel list definitions (SPEC-032) referenced by
   * `ListFormatting.defId`. Absent when the document has none. */
  listDefinitions?: ListDefinition[];
  /** Footnote/endnote content (SPEC-031) referenced by `NoteRef.noteId`.
   * Absent when the document has none. */
  notes?: NoteContent[];
  /** Footnote numbering settings (SPEC-031 AC-17..20); absent means the
   * defaults (decimal, start 1, continuous). */
  footnoteSettings?: NoteSettings;
  /** Endnote numbering settings (SPEC-031); absent means the defaults
   * (lowerRoman, start 1, continuous). */
  endnoteSettings?: NoteSettings;
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
  /** Comment threads (SPEC-038), addressed into `Section.blocks` by the
   * `CommentAnchor` side-lists on their anchoring paragraphs. Absent when
   * the document has none. Deliberately outside `Section.blocks` (like
   * `notes`) so block-tree walks (word count, pagination, TOC, outline, …)
   * never need to know comments exist. */
  comments?: CommentThread[];
  /** Whether track-changes recording was ON when the document was last
   * saved (SPEC-039) — round-trips through `.docx` settings.xml
   * (`w:trackChanges`) and `.odt` settings.xml (`RecordChanges`), but never
   * drives the toggle's initial UI state on open (AC "leaves the toggle
   * OFF"): that is always session-local, starting OFF regardless of this
   * flag's value, per the explicit AC covering documents that already
   * contain tracked changes. */
  trackChangesRecording?: boolean;
  /** Document properties and metadata (SPEC-041) — absent means a document
   * with no known metadata (e.g. never opened from or saved to a file). */
  properties?: DocumentProperties;
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

// --- footnotes and endnotes (SPEC-031) --------------------------------------

/** One footnote/endnote's content (SPEC-031): a block flow like a
 * header/footer's, referenced by `NoteRef.noteId` from a body run. Notes are
 * document-scoped (like `media`), not section-scoped: OOXML keeps them all
 * in one footnotes.xml/endnotes.xml part regardless of sections.
 *
 * A note whose reference mark has been deleted stays in this list but is
 * invisible everywhere (numbering, note areas and both writers derive from
 * the reference marks, never from this list alone) — which is what makes
 * deleting a mark delete its note (AC-10) while still letting undo, and a
 * cut mark pasted back, find the content again. */
export interface NoteContent {
  /** Identity referenced by `NoteRef.noteId` — not named `id` for the same
   * equality.ts reason as `BookmarkAnchor.name`. */
  noteId: string;
  kind: NoteKind;
  blocks: Block[];
}

/** The numbering formats the footnote settings dialog offers (SPEC-031
 * AC-17): 1 2 3 / i ii iii / a b c / * † ‡ §. The first three reuse
 * lists.ts's formatOrdinal vocabulary; "symbol" cycles the traditional
 * footnote symbols, doubling each full cycle (* † ‡ § ** †† …). */
export type NoteNumberFormat = "decimal" | "lowerRoman" | "lowerAlpha" | "symbol";

/** When note numbering restarts (SPEC-031 AC-19, AC-20); absent means
 * continuous through the document. Mirrors `w:numRestart` (.docx) and
 * `text:start-numbering-at` (.odt). */
export type NoteRestart = "continuous" | "page" | "section";

/** One note stream's numbering settings (SPEC-031 AC-17..20). Every field
 * optional with a documented default so an untouched document stores
 * nothing: footnotes default to decimal, endnotes to lowerRoman (Word's own
 * defaults), both starting at 1 and numbering continuously. */
export interface NoteSettings {
  numberFormat?: NoteNumberFormat;
  startAt?: number;
  restart?: NoteRestart;
}

/** Typed pass-through for a note kind coming from a `.vue` template's plain
 * string (same espree constraint as asCharacterFormatKey). */
export function asNoteKind(value: string): NoteKind {
  return value as NoteKind;
}

/** Typed pass-through for a note number format coming from a `.vue`
 * template's plain string (same espree constraint as asCharacterFormatKey). */
export function asNoteNumberFormat(value: string): NoteNumberFormat {
  return value as NoteNumberFormat;
}

/** Typed pass-through for a note restart coming from a `.vue` template's
 * plain string (same espree constraint as asCharacterFormatKey). */
export function asNoteRestart(value: string): NoteRestart {
  return value as NoteRestart;
}

/** A note reference mark's placeholder text — a single character so the mark
 * occupies exactly one caret position (see FIELD_PLACEHOLDER_TEXT, the same
 * convention). */
export const NOTE_PLACEHOLDER_TEXT = "1";

/** Builds a note reference run (SPEC-031 AC-1, AC-3): superscript by
 * construction — the mark renders and saves as a superscript reference in
 * both file formats. */
export function createNoteRun(
  noteId: string,
  kind: NoteKind,
  formatting: RunFormatting = {},
): Run {
  return {
    text: NOTE_PLACEHOLDER_TEXT,
    formatting: { ...formatting, vertAlign: "superscript" },
    note: { noteId, kind },
  };
}

/** A new, empty note content — the seed a freshly inserted footnote/endnote
 * starts from (one empty paragraph, like a blank document's section). */
export function createEmptyNoteContent(noteId: string, kind: NoteKind): NoteContent {
  return { noteId, kind, blocks: [{ type: "paragraph", runs: [] }] };
}

/** Which note DocumentView is currently editing (SPEC-031) — UI state like
 * ActiveHeaderFooterTarget, declared here purely for espree-parsed `.vue`
 * scripts' typed ref seeds. */
export interface ActiveNoteTarget {
  noteId: string;
  kind: NoteKind;
}

/** Typed `null` seed for activeNote refs in `.vue` scripts (same espree
 * constraint as asCharacterFormatKey). */
export function noActiveNote(): ActiveNoteTarget | null {
  return null;
}

/** Typed `null` seed for note-content parameters in `.vue` scripts (same
 * espree constraint as asCharacterFormatKey). */
export function noNoteContent(): NoteContent | null {
  return null;
}

/** The product's built-in fallback styles (A4, standard margins, Calibri
 * 11pt) — the starting point for a blank document, and the fallback the
 * `.docx` reader falls back to for anything a source file doesn't specify. */
export function createDefaultStyles(): DefaultStyles {
  return {
    pageSize: { ...A4_PAGE_SIZE },
    margins: { ...STANDARD_MARGINS },
    fontFamily: DEFAULT_FONT_FAMILY,
    fontSizePt: DEFAULT_FONT_SIZE_PT,
    language: DEFAULT_LANGUAGE,
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

/** A section's header/footer content for one kind, or undefined when that
 * kind has never been given content (SPEC-029) — see HeaderFooterContent's
 * "absent means genuinely blank" contract. There is deliberately no
 * fallback to `"default"` here: resolveHeaderFooterKind already decides
 * which kind applies to a given page. */
export function sectionHeaderFooter(
  model: DocumentModel,
  sectionIndex: number,
  part: HeaderFooterPart,
  kind: HeaderFooterKind,
): HeaderFooterContent | undefined {
  const section = model.sections[sectionIndex];
  return (part === "header" ? section?.headers : section?.footers)?.[kind];
}

function createEmptyParagraph(): Paragraph {
  return { type: "paragraph", runs: [] };
}

/** A new, empty header/footer content — the seed used the first time a
 * section's header/footer kind is edited (SPEC-029). */
export function createEmptyHeaderFooterContent(): HeaderFooterContent {
  return { blocks: [createEmptyParagraph()] };
}

/** Typed empty seed for a `Block[]` default parameter in `.vue` scripts
 * (same espree constraint as asCharacterFormatKey). */
export function noBlocks(): Block[] {
  return [];
}

/** Which header/footer band DocumentView is currently editing (SPEC-029) —
 * UI state, not part of the document model itself; declared here purely so
 * `.vue` scripts (espree-parsed, no inline type annotations) get a typed
 * `ref()` seed the same way noImageLocation/noCaret do elsewhere. */
export interface ActiveHeaderFooterTarget {
  sectionIndex: number;
  part: HeaderFooterPart;
  kind: HeaderFooterKind;
  /** The page (0-based, into DocumentView's pagedViews) whose band is the
   * live editable one — every other page of the section mirrors the same
   * content read-only. */
  pageIndex: number;
}

/** Typed `null` seed for activeHeaderFooter refs in `.vue` scripts (same
 * espree constraint as asCharacterFormatKey). */
export function noActiveHeaderFooter(): ActiveHeaderFooterTarget | null {
  return null;
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
