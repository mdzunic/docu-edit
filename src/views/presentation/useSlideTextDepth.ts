// The slide text-depth controller (SPEC-097) — the presentation view's wiring
// for list authoring, paragraph/letter spacing, text-body properties
// (autofit, anchor, insets, columns) and hyperlinks. It owns:
//   · the Text panel's open state (the `presentationTextOpen` preference) and
//     the values its controls show, folded from the selection ("mixed" → "");
//   · the list commands the toolbar's numbering toggle and list popover emit
//     (style, custom glyph, picture bullet, start-at, indent, hanging);
//   · the body commands (applyTextBody on the edited/selected text or shape
//     frame) and the AUTOFIT WRITE-BACK RULE: a scale/height is stored only
//     inside a mutation that has already called markSlideEdited — picking a
//     mode in text-autofit, or any later command of this composable on that
//     frame (re-measured after the stage re-renders). Rendering never writes;
//   · the link dialog (Cmd/Ctrl+K in a session links the selected characters,
//     the insert bar links a whole selected text/image/shape frame) and the
//     stage's Cmd/Ctrl+click `slide-link-follow` navigation.
// The model rules live in lib/presentation (text.ts, textBody.ts,
// slideLinks.ts); paragraph/run commands run through useSlideTextEditing's
// runTextCommand so a session's host re-renders exactly as for the toolbar.
// Every command is ONE undo entry: record(label) before mutating, commit after.
import {
  computed,
  getCurrentScope,
  nextTick,
  onScopeDispose,
  ref,
  type ComputedRef,
  type Ref,
} from "vue";
import type { PresentationModel, Slide, SlideFrame, SlideParagraph } from "../../lib/presentation/model";
import type { ListMarkerStyle } from "../../lib/core/lists";
import { nextGraphemeBoundary } from "../../lib/core/grapheme";
import { bytesToBase64 } from "../../lib/core/packageParts";
import { mintMediaId } from "../../lib/presentation/frames";
import {
  applyParagraphFormat,
  clampLetterSpacing,
  markSlideEdited,
  parseLineSpacingKey,
  setLetterSpacing,
  setLineSpacing,
  setListStyle,
  setParagraphSpacing,
  toggleBullet,
  type FieldState,
  type LineSpacingChoice,
  type ParagraphOwner,
  type ParagraphRange,
  type SelectionFormatState,
  type TextRange,
} from "../../lib/presentation/text";
import {
  applyTextBody,
  effectiveAnchor,
  effectiveAutofit,
  effectiveColumns,
  effectiveInsetsCm,
  isTextBodyFrame,
  resizedHeightCm,
  type TextAnchor,
  type TextAutofit,
  type TextBodyFrame,
  type TextBodyPatch,
} from "../../lib/presentation/textBody";
import {
  cleanLinkTarget,
  isLinkableKind,
  linkOf,
  setFrameLink,
  setRunLink,
  type SlideLinkTarget,
} from "../../lib/presentation/slideLinks";
import { effectiveHangingCm, effectiveIndentCm } from "../../lib/presentation/slideLists";
import { slideOutline } from "../../lib/presentation/slideModel";
import { frameTextHost, measureFrameFit, type FitScale } from "../../components/presentation/slideTextFit";
import { usePreferences } from "../../composables/usePreferences";
import { PRESENTATION_IMAGE_ACCEPT, PRESENTATION_IMAGE_MIME_TYPES } from "./useFrameInsertion";
import type { FrameSelection } from "./useFrameSelection";
import { paragraphSpan, type SlideTextEditing } from "./useSlideTextEditing";

/** What a rendered text body measures to (slideTextFit.ts#measureFrameFit). */
export interface FrameFit {
  scale: FitScale;
  contentHeightCm: number;
}

export interface SlideTextDepthDeps {
  model: Ref<PresentationModel | null>;
  currentSlide: ComputedRef<Slide | null>;
  currentIndex: Ref<number>;
  /** Snapshot pre-command state for undo (usePresentationUndo.record). */
  record: (label?: string) => void;
  /** Re-render + persist write-through (useSlideRender.commit). */
  commit: () => void;
  /** The stage selection (stageEdit.selection). */
  selection: FrameSelection;
  /** The text-editing session (stageEdit.editing). */
  editing: SlideTextEditing;
  /** Measures a rendered text body. Defaults to the stage DOM measurer;
   * injected by unit tests (jsdom cannot lay text out). */
  measureFit?: (frame: TextBodyFrame) => FrameFit | null;
}

/** The Text panel's line-spacing modes (CustomSpacingDialog's vocabulary). */
export type LineSpacingMode = LineSpacingChoice["mode"];

/** The open link dialog: what it links (a character range of the edited
 * owner, or a whole frame) and the link it starts from. */
export interface SlideLinkDialogState {
  kind: "run" | "frame";
  frameId: string;
  /** The character range a run link covers (captured before the dialog took
   * focus); null for a frame link. */
  range: TextRange | null;
  /** The existing link under the range / on the frame, for prefill. */
  initial: SlideLinkTarget | null;
}

/** One `link-slide` option: a 1-based slide number labelled by its title. */
export interface SlideLinkChoice {
  number: number;
  label: string;
}

/** The custom-bullet message when a glyph is not exactly one grapheme. */
export const BULLET_CHAR_ERROR = "Choose a single character for the bullet.";
/** The picture-bullet message for a non-image file. */
export const BULLET_IMAGE_TYPE_ERROR = "That file isn't a supported image (PNG, JPEG, GIF or WebP).";
/** The link dialog's message for a dangerous or empty address. */
export const LINK_URL_ERROR = "This address type isn't allowed in links";
/** The link dialog's message for a slide outside the deck. */
export const LINK_SLIDE_ERROR = "Choose a slide in this deck";

/** The bubbling event the stage dispatches on Cmd/Ctrl+click of an in-deck
 * slide link (`detail.slide` is 1-based). */
export const SLIDE_LINK_FOLLOW_EVENT = "slide-link-follow";

/** The DOM measurer: the frame's rendered body on the stage (not a filmstrip
 * thumbnail), or null when it is not rendered. */
function stageFit(frame: TextBodyFrame): FrameFit | null {
  if (typeof document === "undefined") return null;
  const stage = document.querySelector('[data-testid="slide-stage"]') ?? document;
  return measureFrameFit(frame, frameTextHost(frame.id, stage));
}

function numberOrBlank(value: FieldState<number>): number | "" {
  return typeof value === "number" ? value : "";
}

/** One fold over a paragraph range: the shared value, "mixed", or undefined. */
function foldParagraphs(owner: ParagraphOwner, range: ParagraphRange, read: (para: SlideParagraph) => number): number | "" {
  let value: number | undefined;
  for (let p = range.startPara; p <= range.endPara; p++) {
    const para = owner.paragraphs[p];
    if (!para) continue;
    const next = read(para);
    if (value === undefined) value = next;
    else if (Math.abs(value - next) > 1e-9) return "";
  }
  return value === undefined ? "" : Math.round(value * 1000) / 1000;
}

/** The run link at a range's start, for the dialog's prefill. */
function runLinkAt(owner: ParagraphOwner, range: TextRange): SlideLinkTarget | null {
  const para = owner.paragraphs[range.start.para];
  if (!para) return null;
  let offset = 0;
  for (const run of para.runs) {
    const end = offset + run.text.length;
    if (range.start.offset < end || (range.start.offset === end && run === para.runs[para.runs.length - 1])) {
      return linkOf(run);
    }
    offset = end;
  }
  return null;
}

function isWord(ch: string | undefined): boolean {
  return ch !== undefined && /\S/.test(ch);
}

/** A collapsed caret widens to the word it sits in (so Cmd/Ctrl+K on a caret
 * links that word); a real selection is kept as is. */
function wordRange(owner: ParagraphOwner, range: TextRange): TextRange {
  const collapsed = range.start.para === range.end.para && range.start.offset === range.end.offset;
  if (!collapsed) return range;
  const para = owner.paragraphs[range.start.para];
  const text = para ? para.runs.map((run) => run.text).join("") : "";
  let start = range.start.offset;
  let end = range.start.offset;
  while (start > 0 && isWord(text[start - 1])) start--;
  while (end < text.length && isWord(text[end])) end++;
  return { start: { para: range.start.para, offset: start }, end: { para: range.start.para, offset: end } };
}

/** A collapsed caret widens to its whole paragraph (letter spacing on a caret
 * spaces the paragraph rather than doing nothing). */
function paragraphsOf(owner: ParagraphOwner, range: TextRange): TextRange {
  const collapsed = range.start.para === range.end.para && range.start.offset === range.end.offset;
  if (!collapsed) return range;
  const para = owner.paragraphs[range.start.para];
  const length = para ? para.runs.reduce((n, run) => n + run.text.length, 0) : 0;
  return { start: { para: range.start.para, offset: 0 }, end: { para: range.start.para, offset: length } };
}

export function useSlideTextDepth(deps: SlideTextDepthDeps) {
  const measure = deps.measureFit ?? stageFit;

  // --- the Text panel's open state --------------------------------------------
  const { preferences, setPreference } = usePreferences();
  const panelOpen = computed(() => preferences.presentationTextOpen);
  function togglePanel(): void {
    setPreference("presentationTextOpen", !panelOpen.value);
  }

  // --- targets ----------------------------------------------------------------
  function frameById(id: string | null): SlideFrame | null {
    const slide = deps.currentSlide.value;
    if (!slide || id === null) return null;
    return slide.frames.find((frame) => frame.id === id) ?? null;
  }

  /** The one selected frame (a multi-selection targets nothing here). */
  const soleSelected = computed<SlideFrame | null>(() =>
    deps.selection.selectedFrameIds.value.length === 1 ? deps.selection.selectedFrame.value : null,
  );

  /** The frame the body controls edit: the frame being edited, else the one
   * selected text/shape frame. Null for a table-cell session. */
  const bodyFrame = computed<TextBodyFrame | null>(() => {
    const editingId = deps.editing.editingFrameId.value;
    const frame = editingId ? frameById(editingId) : soleSelected.value;
    return isTextBodyFrame(frame) ? frame : null;
  });

  /** The selection's paragraph/run formatting (null when no text is targeted). */
  const formatState = computed<SelectionFormatState | null>(() => deps.editing.toolbarState.value);
  const paragraphsEnabled = computed(() => deps.editing.toolbarEnabled.value);
  const bodyEnabled = computed(() => bodyFrame.value !== null);

  // --- values the panel and the list popover show ------------------------------
  const lineSpacingChoice = computed<LineSpacingChoice | "mixed" | null>(() => {
    const token = formatState.value?.lineSpacing;
    if (token === undefined) return null;
    return token === "mixed" ? "mixed" : parseLineSpacingKey(token);
  });
  const lineSpacingMode = computed<LineSpacingMode | "">(() => {
    const choice = lineSpacingChoice.value;
    return choice && choice !== "mixed" ? choice.mode : "";
  });
  const lineSpacingValue = computed<number | "">(() => {
    const choice = lineSpacingChoice.value;
    if (!choice || choice === "mixed") return "";
    return choice.mode === "multiple" ? choice.value : choice.pt;
  });
  const spaceBeforePt = computed(() => numberOrBlank(formatState.value?.spaceBeforePt));
  const spaceAfterPt = computed(() => numberOrBlank(formatState.value?.spaceAfterPt));
  const letterSpacingPt = computed(() => numberOrBlank(formatState.value?.letterSpacingPt));
  const startAt = computed(() => numberOrBlank(formatState.value?.startAt));

  /** The touched paragraphs' text edge and hanging marker box (cm), "" when
   * they differ. Re-read whenever the selection state does. */
  const listGeometry = computed<{ indentCm: number | ""; hangingCm: number | "" }>(() => {
    if (!formatState.value) return { indentCm: "", hangingCm: "" };
    const target = deps.editing.commandTarget();
    if (!target) return { indentCm: "", hangingCm: "" };
    const span = paragraphSpan(target.range);
    return {
      indentCm: foldParagraphs(target.owner, span, effectiveIndentCm),
      hangingCm: foldParagraphs(target.owner, span, effectiveHangingCm),
    };
  });

  const autofit = computed<TextAutofit | "">(() => (bodyFrame.value ? effectiveAutofit(bodyFrame.value) : ""));
  const anchor = computed<TextAnchor | "">(() => (bodyFrame.value ? effectiveAnchor(bodyFrame.value) : ""));
  const insetsCm = computed(() => (bodyFrame.value ? effectiveInsetsCm(bodyFrame.value) : null));
  const columns = computed(() => (bodyFrame.value ? effectiveColumns(bodyFrame.value) : null));

  // --- autofit write-back ------------------------------------------------------
  /** Stores what the measurer reports for a frame's autofit mode — the shrink
   * scale, or the resized height — and returns whether anything changed.
   * Called ONLY inside a mutation that already called markSlideEdited. */
  function writeFit(frame: TextBodyFrame): boolean {
    const mode = effectiveAutofit(frame);
    if (mode === "none") return false;
    const fit = measure(frame);
    if (!fit) return false; // not rendered / cannot lay out: keep what is stored
    if (mode === "shrink") {
      const before = `${frame.body?.fontScalePct}/${frame.body?.lineScalePct}`;
      applyTextBody(frame, {
        fontScalePct: fit.scale.fontScalePct < 100 ? fit.scale.fontScalePct : null,
        lineScalePct: fit.scale.lineScalePct < 100 ? fit.scale.lineScalePct : null,
      });
      return `${frame.body?.fontScalePct}/${frame.body?.lineScalePct}` !== before;
    }
    const height = resizedHeightCm(frame, fit.contentHeightCm);
    if (height === frame.heightCm) return false;
    frame.heightCm = height;
    return true;
  }

  /** After a command on an autofitting frame, re-measure once the stage has
   * re-rendered and fold the result into the SAME edit (no new undo entry —
   * the command's record() already snapshotted the pre-edit state). */
  async function refit(frameId: string | null): Promise<void> {
    if (!frameId) return;
    await nextTick();
    const frame = frameById(frameId);
    const slide = deps.currentSlide.value;
    const model = deps.model.value;
    if (!isTextBodyFrame(frame) || !slide || !model || effectiveAutofit(frame) === "none") return;
    if (writeFit(frame)) {
      markSlideEdited(slide, model.format);
      deps.commit();
    }
  }

  // --- command runners ----------------------------------------------------------
  function editedBodyFrameId(): string | null {
    const target = deps.editing.commandTarget();
    return target && isTextBodyFrame(target.frame) ? target.frame.id : null;
  }

  /** A paragraph command over the selection, then an autofit refit. */
  function paragraphCommand(label: string, apply: (owner: ParagraphOwner, range: ParagraphRange) => void): void {
    const frameId = editedBodyFrameId();
    if (deps.editing.runParagraphCommand(label, apply)) void refit(frameId);
  }

  /** A character-range command over the selection, then an autofit refit. */
  function textCommand(label: string, apply: (owner: ParagraphOwner, range: TextRange) => void, range?: TextRange): void {
    const frameId = editedBodyFrameId();
    if (deps.editing.runTextCommand(label, apply, range)) void refit(frameId);
  }

  /** A frame-level command: record → mutate → markSlideEdited → commit. */
  function frameCommand(label: string, frame: SlideFrame | null, mutate: (frame: SlideFrame) => void): boolean {
    const slide = deps.currentSlide.value;
    const model = deps.model.value;
    if (!frame || !slide || !model) return false;
    deps.record(label);
    mutate(frame);
    markSlideEdited(slide, model.format);
    if (isTextBodyFrame(frame)) writeFit(frame);
    deps.commit();
    void refit(frame.id);
    return true;
  }

  function bodyCommand(label: string, patch: TextBodyPatch): void {
    frameCommand(label, bodyFrame.value, (frame) => {
      if (isTextBodyFrame(frame)) applyTextBody(frame, patch);
    });
  }

  // --- list commands --------------------------------------------------------------
  /** The numbered toggle (slides-numbering): on as "1." for every touched
   * paragraph, or off when they all already are numbered. */
  function toggleNumbering(): void {
    paragraphCommand("Numbering", (owner, range) => toggleBullet(owner, range, "decimal"));
  }

  /** One of core/lists.ts's eight marker styles (the list type follows it). */
  function applyListStyle(style: ListMarkerStyle): void {
    paragraphCommand("List style", (owner, range) => setListStyle(owner, range, style));
  }

  const listError = ref("");

  /** A custom bullet glyph — exactly one grapheme. */
  function setBulletChar(char: string): void {
    if (char.length === 0 || nextGraphemeBoundary(char, 0) !== char.length) {
      listError.value = BULLET_CHAR_ERROR;
      return;
    }
    listError.value = "";
    paragraphCommand("List style", (owner, range) => setListStyle(owner, range, { bulletChar: char }));
  }

  /** A picture bullet: the file is minted into model.media by SPEC-082's
   * naming rule (`ppt/media/image{n}.{ext}` / `Pictures/image{n}.{ext}`) inside
   * the same undo entry that points the paragraphs at it. */
  async function setBulletImage(file: File): Promise<void> {
    const contentType = file.type;
    if (!(PRESENTATION_IMAGE_MIME_TYPES as readonly string[]).includes(contentType)) {
      listError.value = BULLET_IMAGE_TYPE_ERROR;
      return;
    }
    // Capture the target before the read: the selection must not drift.
    const target = deps.editing.commandTarget();
    if (!target) return;
    let data: string;
    try {
      data = bytesToBase64(new Uint8Array(await file.arrayBuffer()));
    } catch {
      listError.value = "That image couldn't be read.";
      return;
    }
    if (deps.editing.commandTarget()?.frame.id !== target.frame.id) return;
    listError.value = "";
    textCommand(
      "List style",
      (owner, range) => {
        const model = deps.model.value;
        if (!model) return;
        const mediaId = mintMediaId(model, contentType);
        // A fresh array reference so the stage's media resolver (and the
        // store's media record) pick the new part up on this commit.
        model.media = [...(model.media ?? []), { mediaId, contentType, data }];
        setListStyle(owner, paragraphSpan(range), { bulletImageId: mediaId });
      },
      target.range,
    );
  }

  function setStartAt(value: number): void {
    if (!Number.isFinite(value)) return;
    paragraphCommand("Numbering", (owner, range) => applyParagraphFormat(owner, range, { startAt: value }));
  }
  function setListIndent(cm: number): void {
    if (!Number.isFinite(cm)) return;
    paragraphCommand("Indent", (owner, range) => applyParagraphFormat(owner, range, { indentCm: cm }));
  }
  function setListHanging(cm: number): void {
    if (!Number.isFinite(cm)) return;
    paragraphCommand("Indent", (owner, range) => applyParagraphFormat(owner, range, { hangingCm: cm }));
  }

  // --- spacing commands -------------------------------------------------------------
  /** The point value a switch from "multiple" starts at: 1.2 × the run size. */
  function defaultLinePt(): number {
    const size = formatState.value?.sizePt;
    return Math.round((typeof size === "number" ? size : 18) * 1.2);
  }

  function setLineSpacingMode(mode: LineSpacingMode): void {
    const current = lineSpacingChoice.value;
    const known = current && current !== "mixed" ? current : null;
    let choice: LineSpacingChoice;
    if (mode === "multiple") {
      choice = { mode, value: known?.mode === "multiple" ? known.value : 1 };
    } else {
      choice = { mode, pt: known && known.mode !== "multiple" ? known.pt : defaultLinePt() };
    }
    paragraphCommand("Line spacing", (owner, range) => setLineSpacing(owner, range, choice));
  }

  function setLineSpacingValue(value: number): void {
    if (!Number.isFinite(value) || value <= 0) return;
    const mode = lineSpacingMode.value || "multiple";
    const choice: LineSpacingChoice = mode === "multiple" ? { mode, value } : { mode, pt: value };
    paragraphCommand("Line spacing", (owner, range) => setLineSpacing(owner, range, choice));
  }

  function setSpaceBefore(pt: number): void {
    if (!Number.isFinite(pt)) return;
    paragraphCommand("Paragraph spacing", (owner, range) => setParagraphSpacing(owner, range, { spaceBeforePt: pt }));
  }
  function setSpaceAfter(pt: number): void {
    if (!Number.isFinite(pt)) return;
    paragraphCommand("Paragraph spacing", (owner, range) => setParagraphSpacing(owner, range, { spaceAfterPt: pt }));
  }

  function setLetterSpacingPt(pt: number): void {
    if (!Number.isFinite(pt)) return;
    const target = deps.editing.commandTarget();
    if (!target) return;
    const value = clampLetterSpacing(pt);
    textCommand("Letter spacing", (owner, range) => setLetterSpacing(owner, range, value), paragraphsOf(target.owner, target.range));
  }

  // --- body commands -----------------------------------------------------------------
  /** Picking a mode in text-autofit is the one place a scale is first stored:
   * "shrink" measures and stores fontScalePct/lineScalePct, "resizeShape" sets
   * heightCm to the measured content height, "none" clears both. */
  function setAutofit(mode: TextAutofit): void {
    bodyCommand("Text autofit", { autofit: mode });
  }
  function setAnchor(value: TextAnchor): void {
    bodyCommand("Text anchor", { anchor: value });
  }
  function setInset(side: "left" | "top" | "right" | "bottom", cm: number): void {
    const current = insetsCm.value;
    if (!current || !Number.isFinite(cm)) return;
    bodyCommand("Text insets", { insetsCm: { ...current, [side]: cm } });
  }
  function setColumnCount(count: number): void {
    const current = columns.value;
    if (!current || !Number.isFinite(count)) return;
    bodyCommand("Text columns", { columns: { count, gapCm: current.gapCm } });
  }
  function setColumnGap(cm: number): void {
    const current = columns.value;
    if (!current || !Number.isFinite(cm)) return;
    bodyCommand("Text columns", { columns: { count: current.count, gapCm: cm } });
  }

  // --- links ------------------------------------------------------------------------
  const linkDialog = ref<SlideLinkDialogState | null>(null);

  /** The one selected frame a whole-frame link applies to (text/image/shape),
   * only outside a text session (inside one, the link covers characters). */
  const linkableFrame = computed<SlideFrame | null>(() => {
    if (deps.editing.editingFrameId.value) return null;
    const frame = soleSelected.value;
    return frame && isLinkableKind(frame.kind) ? frame : null;
  });
  /** Whether the text toolbar's link button acts (a text session is open). */
  const canLinkText = computed(() => deps.editing.editingFrameId.value !== null);

  const slideChoices = computed<SlideLinkChoice[]>(() =>
    (deps.model.value?.slides ?? []).map((slide, index) => ({
      number: index + 1,
      label: slideOutline(slide).title ?? `Slide ${index + 1}`,
    })),
  );

  /** Cmd/Ctrl+K / the toolbar's link button: in a session, link the selected
   * characters (a caret widens to its word); otherwise a whole frame. */
  function openLinkDialog(): void {
    if (deps.editing.editingFrameId.value) {
      const target = deps.editing.commandTarget();
      if (!target) return;
      const range = wordRange(target.owner, target.range);
      linkDialog.value = { kind: "run", frameId: target.frame.id, range, initial: runLinkAt(target.owner, range) };
      return;
    }
    openFrameLinkDialog();
  }

  /** The insert bar's link button: link the whole selected frame. */
  function openFrameLinkDialog(): void {
    const frame = linkableFrame.value;
    if (!frame) return;
    linkDialog.value = { kind: "frame", frameId: frame.id, range: null, initial: linkOf(frame) };
  }

  function closeLinkDialog(): void {
    linkDialog.value = null;
  }

  /** Applies the dialog's target ("Add link"). Returns an error message for
   * an unusable target (a dangerous URL, a slide outside the deck) and keeps
   * the dialog open; "" once applied. */
  function applyLink(target: SlideLinkTarget): string {
    const dialog = linkDialog.value;
    const model = deps.model.value;
    if (!dialog || !model) return "";
    const clean = cleanLinkTarget(target, model.slides.length);
    if (!clean) return target.slide !== undefined ? LINK_SLIDE_ERROR : LINK_URL_ERROR;
    if (dialog.kind === "run") {
      textCommand("Add link", (owner, range) => setRunLink(owner, range, clean), dialog.range ?? undefined);
    } else {
      frameCommand("Add link", frameById(dialog.frameId), (frame) => setFrameLink(frame, clean));
    }
    linkDialog.value = null;
    return "";
  }

  /** Removes the link under the dialog's range / from its frame ("Remove link"). */
  function removeLink(): void {
    const dialog = linkDialog.value;
    if (!dialog) return;
    if (dialog.kind === "run") {
      textCommand("Remove link", (owner, range) => setRunLink(owner, range, null), dialog.range ?? undefined);
    } else {
      frameCommand("Remove link", frameById(dialog.frameId), (frame) => setFrameLink(frame, null));
    }
    linkDialog.value = null;
  }

  // Cmd/Ctrl+K inside a session opens the dialog.
  deps.editing.setLinkShortcut(openLinkDialog);

  /** The stage's Cmd/Ctrl+click on an in-deck link: go to that slide (any
   * session stops first, the selection clears). */
  function onLinkFollow(event: Event): void {
    const model = deps.model.value;
    const slide = event instanceof CustomEvent ? Number(event.detail?.slide) : NaN;
    if (!model || !Number.isInteger(slide) || slide < 1 || slide > model.slides.length) return;
    if (deps.editing.editingFrameId.value) deps.editing.stopEditing();
    deps.selection.clear();
    deps.currentIndex.value = slide - 1;
  }
  if (typeof document !== "undefined") {
    document.addEventListener(SLIDE_LINK_FOLLOW_EVENT, onLinkFollow);
  }
  if (getCurrentScope()) {
    onScopeDispose(() => {
      if (typeof document !== "undefined") document.removeEventListener(SLIDE_LINK_FOLLOW_EVENT, onLinkFollow);
      deps.editing.setLinkShortcut(null);
    });
  }

  return {
    // panel
    panelOpen,
    togglePanel,
    paragraphsEnabled,
    bodyEnabled,
    formatState,
    // values
    lineSpacingMode,
    lineSpacingValue,
    spaceBeforePt,
    spaceAfterPt,
    letterSpacingPt,
    startAt,
    listGeometry,
    autofit,
    anchor,
    insetsCm,
    columns,
    // lists
    imageAccept: PRESENTATION_IMAGE_ACCEPT,
    listError,
    toggleNumbering,
    applyListStyle,
    setBulletChar,
    setBulletImage,
    setStartAt,
    setListIndent,
    setListHanging,
    // spacing
    setLineSpacingMode,
    setLineSpacingValue,
    setSpaceBefore,
    setSpaceAfter,
    setLetterSpacingPt,
    // body
    setAutofit,
    setAnchor,
    setInset,
    setColumnCount,
    setColumnGap,
    // links
    linkDialog,
    linkableFrame,
    canLinkText,
    slideChoices,
    openLinkDialog,
    openFrameLinkDialog,
    closeLinkDialog,
    applyLink,
    removeLink,
    onLinkFollow,
  };
}

export type SlideTextDepth = ReturnType<typeof useSlideTextDepth>;
