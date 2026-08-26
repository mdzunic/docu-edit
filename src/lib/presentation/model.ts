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
import type { DocumentProperties } from "../core/documentProperties";
import type { MediaPart } from "../core/media";
import type { OpaquePart } from "../core/packageParts";
import type { ShapeGeometry } from "../core/shapes";

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

/** A resolved frame/background fill. Only the two forms this spec renders —
 * a solid colour and "no fill"; gradients resolve to their first stop's
 * colour and image/pattern fills to white before they reach here. */
export type SlideFill =
  | { type: "solid"; color: string }
  | { type: "none" };

/** A frame's outline (border): a solid colour and a point width. Shapes have
 * carried one since SPEC-079; SPEC-082 lets text boxes and images take one
 * too (their style popovers edit stroke), so it is a shared type. */
export interface FrameStroke {
  color: string;
  widthPt: number;
}

/** One run of text with its EFFECTIVE formatting already resolved (the run's
 * own properties, else the shape/placeholder/theme chain). Absent optional
 * fields render with the renderer's defaults noted below. */
export interface SlideRun {
  text: string;
  bold?: true;
  italic?: true;
  underline?: true;
  fontFamily?: string; // effective; absent = renderer default stack
  sizePt?: number; // effective; absent = 18
  color?: string; // effective; absent = #000000
}

/** One paragraph (visual line) of a text body. An explicit line break
 * (`a:br`) starts a new paragraph, so each entry is one rendered row. */
export interface SlideParagraph {
  runs: SlideRun[];
  align?: "left" | "center" | "right" | "justify"; // absent = left
  /** Bullet marker: "disc" renders •, "number" renders 1./2./… restarting
   *  per contiguous same-level run of numbered paragraphs. */
  bullet?: "disc" | "number";
  level?: number; // 0..8 indent level; absent = 0
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
}

export interface GroupFrame extends SlideFrameBase {
  kind: "group";
  children: SlideFrame[]; // coordinates already in slide space
  /** The whole source element (`p:grpSp` / `draw:g`), verbatim — what the
   *  writers re-emit; children exist for rendering only until SPEC-082's
   *  ungroup command drops the xml and promotes them. */
  xml: string;
}

export interface ForeignFrame extends SlideFrameBase {
  kind: "foreign";
  xml: string; // verbatim source element
  /** Placeholder-box label: "Table" | "Chart" | "Media" | "Object". */
  label: string;
  /** True when x/y/width/height could not be parsed — render nothing. */
  noBounds?: true;
}

export type SlideFrame = TextFrame | ImageFrame | ShapeFrame | GroupFrame | ForeignFrame;

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
  /** The slide surface size, in centimetres (pptx `p:sldSz`, odp master page
   * `fo:page-width`/`fo:page-height`); the stage's aspect ratio. */
  slideWidthCm: number;
  slideHeightCm: number;
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
  slideWidthCm: number;
  slideHeightCm: number;
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
