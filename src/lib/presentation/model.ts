// The in-memory presentation model — the second product's counterpart of
// documentModel.ts. SPEC-079 turns it into a real, WYSIWYG-renderable slide
// model: each slide is a z-ordered list of frames (text placeholders, images,
// shapes, groups and foreign placeholder boxes) with effective geometry (cm)
// and effective text formatting resolved at read, plus the whole source
// package retained verbatim in `opaque` so a no-edit download stays
// byte-identical (SPEC-008 applied wholesale). Editing specs (SPEC-080+) grow
// and mutate this model and re-render through SPEC-079's renderer — so THIS
// FILE is the normative source for the slide model, the way
// documentModel.ts is for the text model.
//
// Layering rule (enforced by boundary.spec.ts): this directory imports only
// src/lib/core — never documentModel.ts, the commands, or the .docx/.odt
// format layers. Plain serialisable data only (all lengths cm, font sizes pt,
// colours lowercase `#rrggbb` hex), so the document store persists both
// products' models through the same IndexedDB machinery.
import type { CellBorders, CellVerticalAlign } from "../core/tables";
import type { DocumentTheme, ModelColor, ModelFont } from "../core/theme";
import type { DocumentProperties } from "../core/documentProperties";
import type { MediaPart } from "../core/media";
import type { OpaquePart } from "../core/packageParts";
import type { ShapeGeometry } from "../core/shapes";
import type { ChartVisual } from "../core/charts";
import type { ListMarkerStyle } from "../core/lists";

/** The presentation file formats, by the package family they ride in:
 * `.pptx` is OPC/PresentationML, `.odp` is an ODF package. */
export type PresentationFormat = "pptx" | "odp";

/** A placeholder frame's semantic role. Present on a `TextFrame` marks it as
 * a placeholder (vs. a free text box). */
export type PlaceholderRole = "title" | "subtitle" | "body" | "other";

/** The animated slide-transition effects this editor models (SPEC-085) — the
 * deliberate subset that maps cleanly across pptx and odp. `"none"` (an
 * explicit hard cut) is a separate, fully-set model value, distinct from an
 * absent `Slide.transition`. `lib/presentation/transitions.ts` is the
 * normative source for the wire mappings. */
export type TransitionType = "fade" | "push" | "wipe" | "cover" | "uncover" | "split";
export type TransitionDirection = "left" | "right" | "up" | "down";

export interface SlideTransition {
  /** `"none"` = an explicit hard cut the user chose (still a set value). */
  type: TransitionType | "none";
  /** Movement/reveal direction of the incoming content; defaults per type:
   *  push/wipe/cover/uncover → left; fade/split/none ignore it (split is
   *  always horizontal-out). "left" = the incoming slide travels leftward
   *  (enters from the right edge), matching PowerPoint's `dir="l"`. */
  direction?: TransitionDirection;
  /** 100..5000; absent = 600 (`DEFAULT_TRANSITION_DURATION_MS`). */
  durationMs?: number;
  /** Absent = true. Only `false` is stored (the presenter escape hatch keeps
   *  in-editor playback advancing regardless — see SPEC-085 AC-18). */
  advanceOnClick?: false;
  /** Absent = no auto-advance; else advance this slide after N ms. */
  advanceAfterMs?: number;
}

/** The frame animation effects this editor models (SPEC-086) — a deliberate
 * subset of entrance/exit builds with clean equivalents in PowerPoint's
 * `p:timing` and Impress's `anim:*` trees. `lib/presentation/animations.ts`
 * (the grouping/retargeting rules) and the per-format `timing.ts`/`anim.ts`
 * modules (the wire mappings) are the normative source. */
export type AnimationEffect = "appear" | "fadeIn" | "flyIn" | "wipeIn" | "zoomIn" | "fadeOut" | "flyOut";

/** How an animation begins relative to its predecessor in play order. */
export type AnimationTrigger = "click" | "withPrevious" | "afterPrevious";

/** One frame animation, targeting a frame by its minted `SlideFrame.id`
 * (SPEC-082, never serialised — the wire mappings resolve it to each format's
 * shape id). Order in `Slide.animations` is play order. */
export interface SlideAnimation {
  frameId: string;
  effect: AnimationEffect;
  /** flyIn/flyOut/wipeIn only; default "left". "left" means the motion
   *  originates at the LEFT edge: fly-in enters from the left, fly-out exits
   *  toward the left, wipe reveals starting at the left. */
  direction?: TransitionDirection;
  trigger: AnimationTrigger;
  durationMs?: number; // default 500, range 100–5000
  delayMs?: number; // default 0, range 0–10000
}

/** The slide-layout kinds this editor recognises (SPEC-089), a closed union
 * mapped from pptx `p:sldLayout/@type` and classified from odp page-layout
 * slots. Anything outside the set — a custom pptx layout, an Impress page
 * layout with an unrecognised placeholder object — classifies `"other"` and is
 * still offered and applied unchanged. `lib/presentation/layouts.ts` is the
 * normative source for classification and apply/reset. */
export type SlideLayoutKind =
  | "title"
  | "titleAndBody"
  | "sectionHeader"
  | "twoContent"
  | "titleOnly"
  | "blank"
  | "other";

/** One placeholder slot of a `SlideLayout`: a role plus the geometry (cm) a
 * slide's matching placeholder snaps to when the layout is applied. `phType`/
 * `phIdx` carry the pptx `p:ph` `@type`/`@idx` written back verbatim; ODF has
 * no `p:ph idx`, so both are left undefined on odp-derived slots. */
export interface LayoutSlot {
  role: PlaceholderRole;
  phType?: string;
  phIdx?: number;
  xCm: number;
  yCm: number;
  widthCm: number;
  heightCm: number;
}

/** One entry in a deck's read-derived layout catalogue (SPEC-089). Plain
 * serialisable data per this file's rule, so `JSON.parse(JSON.stringify(model))`
 * still `toEqual`s the model. */
export interface SlideLayout {
  /** Deterministic, deck-stable, survives save → reopen. pptx: the absolute
   *  layout part path; odp: `${master}::${pageLayout}`; synthesised:
   *  `standard:{kind}`. Never a UUID and never named `id` (equalStructural's
   *  IGNORED_KEYS drops `"id"`). */
  key: string;
  kind: SlideLayoutKind;
  /** The deck's own name for the layout, shown on the gallery card. */
  name: string;
  /** The layout's placeholder slots, in document order; `[]` for `"blank"`. */
  slots: LayoutSlot[];
  /** odp `draw:master-page-name` / pptx master part path the layout hangs off. */
  masterName?: string;
}

/** The deck's theme (SPEC-094): the shared `DocumentTheme` plus the wiring a
 * presentation package needs — the first master's `p:clrMap`, the theme part's
 * path, and the edited flag the writers read. Plain serialisable data. */
export interface DeckTheme extends DocumentTheme {
  /** pptx `p:clrMap` from the first master: "bg1"|"tx1"|"bg2"|"tx2"|
   *  "accent1".."folHlink" → scheme slot. Absent = identity. */
  colorMap?: Record<string, string>;
  /** "ppt/theme/theme1.xml" | "styles.xml". */
  partName?: string;
  /** Edited since open → the writers rewrite that part in place. */
  dirty?: true;
}

/** A resolved frame/background fill. Only the two forms this spec renders —
 * a solid colour and "no fill"; gradients resolve to their first stop's
 * colour and image/pattern fills to white before they reach here. SPEC-094
 * widens `color` to `ModelColor` — a scheme reference resolves through
 * `resolveColor` at every render/serialise point; the string form stays valid. */
export type SlideFill =
  | { type: "solid"; color: ModelColor }
  | { type: "none" };

/** A frame's outline (border): a solid colour and a point width. Shapes have
 * carried one since SPEC-079; SPEC-082 lets text boxes and images take one
 * too (their style popovers edit stroke), so it is a shared type. */
export interface FrameStroke {
  color: ModelColor;
  widthPt: number;
}

/** One run of text with its EFFECTIVE formatting already resolved (the run's
 * own properties, else the shape/placeholder/theme chain). Absent optional
 * fields render with the renderer's defaults noted below. SPEC-094 widens
 * `color`/`fontFamily` to scheme references; the string forms stay valid. */
export interface SlideRun {
  text: string;
  bold?: true;
  italic?: true;
  underline?: true;
  fontFamily?: ModelFont; // effective; absent = renderer default stack
  sizePt?: number; // effective; absent = 18
  color?: ModelColor; // effective; absent = #000000
  /** Letter spacing in points, −10…10 in 0.1 steps; absent = 0 (SPEC-097). */
  letterSpacingPt?: number;
  /** SPEC-097 hyperlink: an external URL (always `sanitizeHref`-clean) OR a
   *  1-based in-deck slide number — at most one of the two is set. Two
   *  scalars rather than one object, exactly documentModel.ts's
   *  `linkHref`/`linkAnchor` reason: run coalescing and `marksEqual` stay
   *  `===` per key. A run link inside a linked frame wins for that run. */
  linkHref?: string;
  linkSlide?: number;
  /** The link's hover tooltip; meaningless without a link. */
  linkTooltip?: string;
}

/** One paragraph (visual line) of a text body. An explicit line break
 * (`a:br`) starts a new paragraph, so each entry is one rendered row.
 *
 * SPEC-097 adds the list, spacing and preservation fields. Displayed numbers
 * are DERIVED by adjacency (`lib/presentation/slideLists.ts`), never stored. */
export interface SlideParagraph {
  runs: SlideRun[];
  align?: "left" | "center" | "right" | "justify"; // absent = left
  /** Bullet marker: "disc" = a bulleted item, "number" = a numbered item
   *  (SPEC-079's meaning); the exact glyph/format is `bulletStyle`. */
  bullet?: "disc" | "number";
  level?: number; // 0..8 indent level; absent = 0
  /** The marker style (core/lists.ts); absent = "disc" for `bullet: "disc"`,
   *  "decimal" for `bullet: "number"`. `markerType(bulletStyle)` always agrees
   *  with `bullet`. */
  bulletStyle?: ListMarkerStyle;
  /** A custom bullet glyph (one grapheme), `bullet: "disc"` only; wins over
   *  `bulletStyle`. */
  bulletChar?: string;
  /** A picture bullet: a `model.media` part name; wins over `bulletStyle` and
   *  `bulletChar`. Deliberately not named `id` (equalStructural's
   *  IGNORED_KEYS). */
  bulletImageId?: string;
  /** The number a numbered run starts at, 1–32767; absent = 1 / continue. */
  startAt?: number;
  /** Text edge (cm from the box's inner left); absent =
   *  `(level + 1) × LIST_INDENT_STEP_CM` for a list item, 0 otherwise. */
  indentCm?: number;
  /** The marker box hanging before the text edge (cm); absent =
   *  `LIST_MARKER_WIDTH_CM` (0.635). */
  hangingCm?: number;
  /** Line spacing as a multiple, 0.25–5; absent = 1. */
  lineSpacing?: number;
  /** Line spacing in points, 4–1000; wins over `lineSpacing`. */
  lineSpacingPt?: number;
  /** `"exact"` = `lineSpacingPt` is exact; absent = "at least" — the same
   *  three fields and absent meaning as documentModel's ParagraphFormatting. */
  lineSpacingRule?: "exact";
  spaceBeforePt?: number; // 0–500; absent = 0
  spaceAfterPt?: number; // 0–500; absent = 0
  /** pptx: the source `a:pPr`'s UNMODELLED attributes and children, verbatim
   *  — a `<a:pPr …attrs…>…children…</a:pPr>` element string (see
   *  pptx/textDepthWire.ts), re-emitted in schema order and cleared by the
   *  command that edits a modelled property it shadows (Table.docxTblPr's
   *  rule). Not in equality.ts's XML_STRING_KEYS, so it compares raw. */
  pptxPPr?: string;
  /** odp: the source paragraph style's unmodelled property string. Same
   *  clearing rule. */
  odpParaProps?: string;
}

/** The text-body properties of a text-carrying frame (SPEC-097). Declared
 * here, beside SlideTransition/SlideAnimation, never downstream in
 * textBody.ts — textBody.ts's derivations take `TextFrame`/`ShapeFrame`, so a
 * downstream declaration would invert the layering into a cycle. */
export type TextAutofit = "none" | "shrink" | "resizeShape";
export type TextAnchor = "top" | "middle" | "bottom";

export interface TextBodyProps {
  /** Absent = "none" (overflow exactly as SPEC-079 renders it). */
  autofit?: TextAutofit;
  /** "shrink" only: the stored font scale, 10–100. */
  fontScalePct?: number;
  /** "shrink" only: the stored line-spacing scale, 80–100. */
  lineScalePct?: number;
  /** Absent = "top" on a TextFrame, "middle" on a ShapeFrame. */
  anchor?: TextAnchor;
  /** Absent = 0.254/0.127/0.254/0.127 (the pptx EMU defaults). */
  insetsCm?: { left: number; top: number; right: number; bottom: number };
  /** count 1–16; absent = one column. */
  columns?: { count: number; gapCm: number };
  /** `false` = no wrapping; absent = wrap. */
  wrap?: false;
}

interface SlideFrameBase {
  /** Minted at read (crypto.randomUUID()); stable for the model's life —
   *  selection, undo and (SPEC-086) animation targets hang off it. Never
   *  serialised to either file format. */
  id: string;
  xCm: number;
  yCm: number;
  widthCm: number;
  heightCm: number;
  rotationDeg?: number; // clockwise, like DrawingPlacement's convention
  /** SPEC-097 frame link (text/image/shape frames only — `isLinkableKind`):
   *  an external URL or a 1-based slide number, at most one set. */
  linkHref?: string;
  linkSlide?: number;
  linkTooltip?: string;
}

export interface TextFrame extends SlideFrameBase {
  kind: "text";
  role?: PlaceholderRole; // present = it is a placeholder
  /** pptx only: the source `p:ph`'s type/idx attributes, written back
   *  verbatim on rebuild so Office keeps the inheritance chain. */
  phType?: string;
  phIdx?: number;
  paragraphs: SlideParagraph[];
  fill?: SlideFill; // frame background; absent = none
  stroke?: FrameStroke; // frame border; absent = none (SPEC-082)
  /** SPEC-097 text-body properties; absent = every default. */
  body?: TextBodyProps;
  /** pptx: the source `a:bodyPr`'s unmodelled attributes/children, verbatim
   *  — a `<a:bodyPr …attrs…>…children…</a:bodyPr>` element string (see
   *  pptx/textDepthWire.ts); cleared by the command that edits a modelled body
   *  property it shadows. Compares raw under equalStructural. */
  pptxBodyPr?: string;
  /** odp: the source graphic style's unmodelled text-body property string. */
  odpBodyProps?: string;
}

export interface ImageFrame extends SlideFrameBase {
  kind: "image";
  mediaId: string; // key into model.media
  stroke?: FrameStroke; // frame border; absent = none (SPEC-082)
}

export interface ShapeFrame extends SlideFrameBase {
  kind: "shape";
  geometry: ShapeGeometry; // core/shapes id, e.g. "rect", "ellipse"
  adjustments?: number[]; // core/shapeCatalogue fractions
  fill?: SlideFill;
  stroke?: FrameStroke;
  paragraphs?: SlideParagraph[]; // shape text, centred per bodyPr default
  /** SPEC-097 text-body properties (anchor absent = "middle" here). */
  body?: TextBodyProps;
  pptxBodyPr?: string;
  odpBodyProps?: string;
}

/** A slide media clip is one of two element kinds (SPEC-095). */
export type SlideMediaKind = "video" | "audio";

/** An audio or video clip on a slide (SPEC-095). Read from a pptx `p:pic`
 * carrying `a:videoFile`/`a:audioFile`/`p14:media` or an odp `draw:frame`
 * carrying `draw:plugin`; rendered on the stage with inline playback and in
 * SPEC-084 playback under the autoplay policy. `mediaId` absent is never "no
 * media": it pairs with `linkTarget` (a clip we did not embed) or `oversize`
 * (a payload over MAX_MEDIA_SIZE_BYTES we left in `opaque`). */
export interface MediaFrame extends SlideFrameBase {
  kind: "media";
  media: SlideMediaKind;
  /** Key into `model.media`. Absent ⇒ see `linkTarget`/`oversize`. */
  mediaId?: string;
  /** Source href/target of a clip we did not embed (a linked file). */
  linkTarget?: string;
  /** Poster still (`p:blipFill` / sibling `draw:image`); absent = neutral plate. */
  posterMediaId?: string;
  /** From the package part or `draw:mime-type`. */
  contentType?: string;
  autoStart?: true; // absent = start on click
  loop?: true;
  muted?: true;
  volume?: number; // 0..100; absent = 100
  fullScreen?: true; // odp presentation:play-full-screen
  oversize?: true; // in the package, over MAX_MEDIA_SIZE_BYTES
  stroke?: FrameStroke;
  /** The verbatim source `p:pic` / `draw:frame`. Present ⇒ the serialisers
   *  re-emit it unchanged; deleted by any command changing a modelled field
   *  (`markFrameEdited`, the frame-level analogue of `markSlideEdited`'s odp
   *  rule). Absent ⇒ rebuild from the fields above. Not in
   *  equality.ts's `XML_STRING_KEYS`, so it compares as a raw string —
   *  byte-exactness is the point. */
  sourceXml?: string;
}

export interface GroupFrame extends SlideFrameBase {
  kind: "group";
  children: SlideFrame[]; // coordinates already in slide space
  /** The whole source `p:grpSp` / `draw:g`, verbatim. ABSENT means an authored
   *  group the serialisers rebuild — exactly `Slide.sourceXml`'s discipline
   *  (SPEC-096). Any command changing membership, child geometry or rotation
   *  deletes it first — the only way this field ever changes. */
  xml?: string;
}

export interface ForeignFrame extends SlideFrameBase {
  kind: "foreign";
  xml: string; // verbatim source element
  /** Placeholder-box label: "Table" | "Chart" | "Media" | "Object". */
  label: string;
  /** True when x/y/width/height could not be parsed — render nothing. */
  noBounds?: true;
}

/** A chart on a slide (SPEC-092): the product-neutral `ChartVisual` (chart
 * type, categories, series, options + optional out-of-catalogue
 * `unsupportedChartXml`) placed on the slide stage. Its part bytes live in
 * `PresentationModel.opaque` under `sourcePart` (the chart part name, never
 * bytes on the frame — same rule as the text `Chart` block); `edited` and
 * `xml` decide regenerate-vs-verbatim, exactly SPEC-052's switch. */
export interface ChartFrame extends SlideFrameBase, ChartVisual {
  kind: "chart";
  /** "ppt/charts/chart1.xml" | "Object 1/content.xml" — a NAME into
   *  `model.opaque`, never bytes; absent for a chart created here. Not
   *  called `id`, which equalStructural's IGNORED_KEYS would drop. */
  sourcePart?: string;
  /** Set only by the data sheet / options panel — the switch the writer reads
   *  to regenerate the part from the model vs. re-emit `sourcePart`'s bytes.
   *  A move/resize/reorder never sets it (they don't touch the chart part). */
  edited?: true;
  /** Verbatim source `p:graphicFrame` / `draw:frame`, re-emitted while the
   *  frame is untouched; deleted by any geometry or data change, the
   *  serialiser then writing the canonical container. In XML_STRING_KEYS so it
   *  compares canonically under equalStructural. */
  xml?: string;
}

/** The default cell padding (cm) a cell renders with when it carries no
 * explicit `paddingCm` (SPEC-091 AC-1). ~1/20 inch, matching PowerPoint's
 * default `a:tcPr` internal margins. */
export const DEFAULT_TABLE_CELL_PADDING_CM = 0.13;

/** The default row height a freshly-inserted table's rows take (SPEC-091
 * AC-5c). */
export const DEFAULT_TABLE_ROW_HEIGHT_CM = 0.9;

/** One cell of a `TableFrame` (SPEC-091). A covered slot — the space consumed
 * by a neighbour's `colSpan`/`rowSpan` — carries `covered: true` with
 * `paragraphs: []` and is never an editing or navigation target; the owner
 * cell holds the content and the span. Cell content is edited by SPEC-081's
 * text machinery, exactly as a `TextFrame`'s `paragraphs` are. */
export interface SlideTableCell {
  paragraphs: SlideParagraph[];
  /** Columns this cell spans (owner cell only); absent = 1. */
  colSpan?: number;
  /** Rows this cell spans (owner cell only); absent = 1. */
  rowSpan?: number;
  /** A slot covered by a neighbour's span: `paragraphs` is always `[]`. */
  covered?: true;
  fill?: SlideFill; // cell background; absent = inherit look / none
  borders?: CellBorders; // per-edge borders (core/tables shape); absent = look
  paddingCm?: number; // internal margin; absent = DEFAULT_TABLE_CELL_PADDING_CM
  verticalAlign?: CellVerticalAlign; // absent = "top"
  /** pptx: unmodelled `a:tcPr` attributes/children preserved verbatim, cleared
   *  by the command that edits the modelled property it shadows (AC-12). */
  pptxTcPr?: string;
  /** odp: the cell's automatic style name, re-emitted verbatim (AC-3c). */
  odpCellStyleName?: string;
}

/** One row of a `TableFrame` (SPEC-091): exactly `columnWidthsCm.length`
 * cells (covered slots included). */
export interface SlideTableRow {
  cells: SlideTableCell[];
  /** Row height (pptx `a:tr@h`, odp `style:row-height`); absent = auto. */
  heightCm?: number;
  /** A header row (odp `table:table-header-rows`), for the header look. */
  isHeader?: true;
  /** odp: the row's automatic style name, re-emitted verbatim (AC-3c). */
  odpRowStyleName?: string;
}

/** A table on a slide (SPEC-091): a real editable grid that round-trips
 * through pptx `a:tbl` and odp `table:table`. Every row holds exactly
 * `columnWidthsCm.length` cells. Unmodelled markup is preserved verbatim in
 * the `pptx*`/`odp*` fields, each cleared by the command that edits the
 * modelled property it shadows (AC-12), exactly Table.docxTblPr's rule. */
export interface TableFrame extends SlideFrameBase {
  kind: "table";
  rows: SlideTableRow[];
  columnWidthsCm: number[];
  /** Give the first (header) row its distinct look (AC-4b, AC-10d). */
  firstRow?: true;
  /** Band alternate body rows (AC-4b, AC-10d). */
  bandRow?: true;
  /** The deck table style this table resolves its look from (AC-4b). */
  styleId?: string;
  /** pptx: unmodelled `a:tblPr` markup preserved verbatim (AC-2c). */
  pptxTblPr?: string;
  /** pptx: unmodelled `p:graphicFrame`-level markup (`p:nvGraphicFramePr`,
   *  transforms) preserved verbatim (AC-2c). */
  pptxGraphicFramePr?: string;
  /** odp: the table's automatic style name, re-emitted verbatim (AC-3c). */
  odpTableStyleName?: string;
  /** odp: per-column automatic style names, aligned with `columnWidthsCm`
   *  (AC-3c). A resize/insert that no longer matches drops them. */
  odpColumnStyleNames?: (string | undefined)[];
}

export type SlideFrame =
  | TextFrame
  | ImageFrame
  | ShapeFrame
  | ChartFrame
  | MediaFrame
  | GroupFrame
  | ForeignFrame
  | TableFrame;

/** The three deck-level field kinds SPEC-093 models along a slide's bottom
 * edge: a footer string, the slide number and the date. */
export type SlideFieldKind = "footer" | "slideNumber" | "date";

/** The deck-level header/footer values (SPEC-093), shared by every slide that
 * shows a field. Per-slide *visibility* lives on `Slide.fields`; only the
 * values and the title-slide policy live here. All-absent is a deck with no
 * configured footer/date (numbers still count from `firstSlideNumber`). Plain
 * JSON — `roundTrip.spec.ts`'s `JSON.parse(JSON.stringify(model))` must hold. */
export interface DeckHeaderFooter {
  /** The footer text every showing slide displays; absent = none. A per-slide
   *  `SlideFieldPlaceholder.text` overrides it for that slide. */
  footerText?: string;
  /** `"fixed"` = a frozen date string (`dateFixedText`); absent = an
   *  auto-updating date rendered from `dateFormat` at every render and save. */
  dateSource?: "fixed";
  /** The frozen string shown when `dateSource === "fixed"`. */
  dateFixedText?: string;
  /** A core `dateTimeFormats` catalogue pattern for the auto date; absent =
   *  `d MMMM yyyy` (`DEFAULT_DATE_FORMAT`). */
  dateFormat?: string;
  /** `false` = suppress every field on the title slide; absent = true. Only
   *  `false` is ever stored (the absent-is-true rule). */
  showOnTitleSlide?: false;
}

/** One field placeholder on a slide (SPEC-093): its presence in `Slide.fields`
 * IS the slide's visibility for that kind. Geometry and look are resolved once
 * at read (slide's own `a:xfrm`/`svg:*` → the reader's placeholder maps →
 * `DEFAULT_FIELD_BAND`), so rendering never re-walks the layout chain. Plain
 * JSON. `slideNumber` entries never carry `text`. */
export interface SlideFieldPlaceholder {
  kind: SlideFieldKind;
  xCm: number;
  yCm: number;
  widthCm: number;
  heightCm: number;
  /** Text alignment inside the box; absent = left. */
  align?: "left" | "center" | "right";
  sizePt?: number; // effective; absent = 12
  color?: string; // effective; absent = #595959
  fontFamily?: string; // effective; absent = renderer default
  /** A per-slide override of the deck value (footer/date only); absent = the
   *  deck value. */
  text?: string;
  /** pptx `a:fld @type` / odp declaration source, kept verbatim so an
   *  out-of-subset field (SPEC-093 "Date formats") re-emits unchanged. */
  wireType?: string;
  /** odp `style:data-style-name` of an auto date, kept verbatim for the same
   *  preserve-don't-parse reason. */
  wireStyleName?: string;
}

export interface Slide {
  id: string; // minted at read, never serialised
  frames: SlideFrame[]; // z-order, back to front
  background?: SlideFill; // effective (inheritance resolved)
  /** Speaker notes as plain paragraphs (SPEC-083): pptx from the notes-slide
   *  part's body placeholder, odp from the page's `presentation:notes` frame,
   *  both with formatting flattened and trailing empty paragraphs stripped.
   *  Absent when the slide has none — `[]` is never stored or produced. */
  notes?: string[];
  /** Set by any command that changes this slide's own content or properties
   *  (SPEC-081+ set it; this spec's new/duplicated slides are born without
   *  partName/sourceXml instead). Dirty slides are rebuilt by the writers;
   *  clean ones re-emit verbatim (SPEC-080). */
  dirty?: true;
  /** Set by a notes edit (SPEC-083) — distinct from `dirty` so a notes-only
   *  change rebuilds just the notes-slide part while the slide part stays
   *  byte-verbatim (pptx AC). Odp rebuilds the whole page either way. */
  notesDirty?: true;
  /** The slide's animated entrance transition (SPEC-085): a modelled subset
   *  value the readers populate from the wire mappings and the writers
   *  serialise. Absent means either a plain slide OR a preserved out-of-subset
   *  transition (distinguished by `preserved`); an explicit hard cut the user
   *  chose is `{ type: "none" }`. See `lib/presentation/transitions.ts`. */
  transition?: SlideTransition;
  /** The slide's frame animations in play order (SPEC-086): a modelled subset
   *  the readers populate from the wire mappings and the writers serialise.
   *  Absent means either a plain slide OR preserved out-of-subset timing
   *  (distinguished by `preserved["p:timing"]`/`preserved["anim"]`) — a slide
   *  whose timing this editor can't model keeps `animations` unset and its
   *  source verbatim. A present (even empty) array is the modelled state: the
   *  writers emit it and drop the preserved timing. See
   *  `lib/presentation/animations.ts`. */
  animations?: SlideAnimation[];
  /** The catalogue key (`PresentationModel.layouts[].key`) of the layout this
   *  slide uses (SPEC-089). Set at read from the source layout relationship /
   *  page-layout reference, and by `applySlideLayout`. A `layoutKey` naming no
   *  catalogue entry resolves to `undefined` and is left alone on save — never
   *  repaired, never dropped. */
  layoutKey?: string;
  /** pptx: this slide's part path ("ppt/slides/slide2.xml"); ties the
   *  model slide to its verbatim part in `opaque`. Absent for .odp and
   *  (later) from-scratch slides. */
  partName?: string;
  /** odp: the source `draw:page` element markup, verbatim. SPEC-081's
   *  writer splices it back for untouched pages. */
  sourceXml?: string;
  /** Verbatim XML captured for rebuild-time retention (SPEC-081 uses it;
   *  this spec only fills it): pptx `p:bg`, `p:transition`, `p:timing`,
   *  `p:clrMapOvr`, `p:extLst` and the root `p:sld` attribute string; odp
   *  `presentation:notes`, `anim:*` children and the `draw:page` attribute
   *  string. Keys are those element names. */
  preserved?: Record<string, string>;
  /** The slide's visible deck fields (SPEC-093): at most one entry per kind
   *  `footer`/`slideNumber`/`date`. Presence is visibility — absent or `[]`
   *  means the slide shows none. The pptx reader also keeps each source `p:sp`
   *  in `preserved["p:ph:ftr"|"p:ph:sldNum"|"p:ph:dt"]` so save re-emits it
   *  with only its cached `a:t` refreshed. */
  fields?: SlideFieldPlaceholder[];
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
  slides: Slide[];
  /** The deck's theme (SPEC-094): scheme + fonts read from the theme part
   * (pptx) / the `de-theme-*` block (odp). Absent = an unthemed deck —
   * everything resolves against `DEFAULT_THEME` and every colour stays a
   * literal; nothing is retroactively converted to a reference. */
  theme?: DeckTheme;
  /** The deck's read-derived layout catalogue (SPEC-089), in package order
   * with every master's layouts flattened and de-duplicated by `key`. Absent
   * means a deck stored before SPEC-089 (re-derived on the next open); `[]`
   * means a package with no layout parts (the gallery falls back to the
   * synthesised standard set). */
  layouts?: SlideLayout[];
  /** The slide surface size, in centimetres (pptx `p:sldSz`, odp master page
   * `fo:page-width`/`fo:page-height`); the stage's aspect ratio. */
  slideWidthCm: number;
  slideHeightCm: number;
  /** Slide-number origin (pptx `p:presentation @firstSlideNum`), 1..9999;
   * absent = 1, and only a value ≠ 1 is stored (SPEC-090). Model-only on odp
   * (that format does not serialise it). */
  firstSlideNumber?: number;
  /** The deck-level footer/date values and title-slide policy (SPEC-093);
   * absent = a deck with no configured footer/date. Per-slide visibility lives
   * on each `Slide.fields`. */
  headerFooter?: DeckHeaderFooter;
  /** The header/footer values or any slide's field visibility changed since
   * open → the writers regenerate the field placeholders and (pptx)
   * `showSpecialPlsOnTitleSld` / (odp) the declaration block. Absent =
   * unchanged since open; only `true` is stored, like `deckDirty` (SPEC-093).
   * `stripRetention` clears it. */
  headerFooterDirty?: true;
  /** Deck-level setup (size, first slide number, properties) changed since
   * open → the writers rebuild the deck-level parts. Absent = unchanged since
   * open; only `true` is stored, exactly like `structureDirty` (SPEC-090). */
  deckDirty?: true;
  /** Every part of the source package, verbatim and in package order —
   * SPEC-008's retention rule applied wholesale. The writers re-emit exactly
   * these, so a no-edit download stays byte-identical; media referenced by
   * `ImageFrame`s is ALSO copied into `media`, but the bytes stay here too
   * (bounded duplication until SPEC-081's rebuild writers take over media).
   * (`.odp`'s `mimetype` entry is included and written first + STOREd.)
   *
   * Absent on from-scratch decks (`createBlankPresentation`) and on the copy
   * `stripRetention` hands the cross-format writer: with no retained package
   * the writers build a fresh minimal package instead (SPEC-080). */
  opaque?: OpaquePart[];
  /** Image bytes split out the way `DocumentModel.media` is, so the store's
   * media split (MEDIA_KEY_PREFIX) persists them once per record and an
   * `ImageFrame` references one by `mediaId` (AC-32). */
  media?: MediaPart[];
  /** Document metadata read from `docProps/core.xml` (.pptx) / `meta.xml`
   * (.odp); absent when the package declares none. */
  properties?: DocumentProperties;
  /** Set when slides were added/removed/reordered since open (SPEC-080).
   * Absent on a freshly opened deck; from-scratch decks are always fully
   * built. The writers rebuild the structural parts only when this (or a
   * slide's `dirty`) is set. */
  structureDirty?: true;
  /** odp only: content.xml split at read — everything before the first
   * `draw:page` and everything after the last (the `office:presentation`
   * close onwards). `prefix + pages.join("") + suffix` equals the source
   * content.xml byte-for-byte for an untouched deck. Absent for pptx and
   * from-scratch decks (SPEC-080). */
  odpContentEnvelope?: { prefix: string; suffix: string };
}

export const DEFAULT_PRESENTATION_TITLE = "Untitled presentation";

/** The default slide size when a package declares none: PowerPoint's 16:9
 * widescreen (12192000×6858000 EMU). */
export const DEFAULT_SLIDE_WIDTH_CM = 33.867;
export const DEFAULT_SLIDE_HEIGHT_CM = 19.05;

/** What a presentation reader hands back — the shared assembly input of
 * `assemblePresentation` (presentation/open.ts), the same split as the text
 * side's `ParsedDocxDocument`/`ParsedOdtDocument` vs `assembleDocumentModel`. */
export interface ParsedPresentation {
  slides: Slide[];
  /** The deck's theme (SPEC-094); absent when the package declares no scheme. */
  theme?: DeckTheme;
  /** The deck's read-derived layout catalogue (SPEC-089); `[]` when the
   * package carries no layout parts. */
  layouts: SlideLayout[];
  slideWidthCm: number;
  slideHeightCm: number;
  /** Slide-number origin from pptx `p:presentation @firstSlideNum`; absent for
   * odp and for a deck that counts from 1 (SPEC-090). */
  firstSlideNumber?: number;
  /** The deck-level footer/date values the reader resolved (SPEC-093); absent
   * when the deck configures no footer/date. */
  headerFooter?: DeckHeaderFooter;
  /** Image bytes referenced by the slides' `ImageFrame`s. */
  media: MediaPart[];
  /** Every source package part, verbatim, in package order. */
  opaque: OpaquePart[];
  /** Metadata from `docProps/core.xml`/`meta.xml`; `{}` when none. */
  properties: DocumentProperties;
  /** odp only: content.xml split around its `draw:page`s (SPEC-080) — the
   * writer's splice-equality envelope. Absent for pptx. */
  odpContentEnvelope?: { prefix: string; suffix: string };
}

/** How many slides the library/preview surfaces report for a model. */
export function slideCount(model: PresentationModel): number {
  return model.slides.length;
}

/** Whether a stored record's model predates SPEC-079 (its slides are the
 * preparation milestone's outline shape, not `Slide[]`): it has no
 * `slideWidthCm`. The view upgrades such a record on open (AC-54). */
export function isLegacyPresentationModel(model: PresentationModel): boolean {
  return typeof (model as { slideWidthCm?: unknown }).slideWidthCm !== "number";
}
