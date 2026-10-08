// The slide text-depth controller (SPEC-097) — PresentationView's ONE wiring
// call for list authoring, paragraph/letter spacing, text-body properties and
// hyperlinks, consumed by the Text panel, the toolbar's numbering toggle and
// list popover, the insert bar's link button and the slide link dialog. It
// owns the Text panel's open state (the `presentationTextOpen` preference),
// the values its controls show (folded from the selection, "mixed" → ""), and
// the list and spacing commands; it composes the body/autofit half
// (useSlideTextBody — which holds the autofit write-back rule) and the link
// half (useSlideLinks). Paragraph/run commands run through
// useSlideTextEditing's runTextCommand, so a session's host re-renders exactly
// as for the toolbar, and each is ONE undo entry (record before, commit after)
// in the spec's labels. Every text edit the session commits — these commands,
// the toolbar's, typing, paste, cut, a session's end — re-measures an
// autofitting frame inside that same edit, through the session's after-edit
// hook (AC-38/41). The model rules live in lib/presentation.
import { computed, getCurrentScope, onScopeDispose, ref, type ComputedRef, type Ref } from "vue";
import type { PresentationModel, Slide, SlideFrame } from "../../lib/presentation/model";
import type { ListMarkerStyle } from "../../lib/core/lists";
import { nextGraphemeBoundary } from "../../lib/core/grapheme";
import { bytesToBase64 } from "../../lib/core/packageParts";
import { mintMediaId } from "../../lib/presentation/frames";
import {
  applyParagraphFormat,
  clampLetterSpacing,
  parseLineSpacingKey,
  setLetterSpacing,
  setLineSpacing,
  setListStyle,
  setParagraphSpacing,
  toggleBullet,
  type LineSpacingChoice,
  type ParagraphOwner,
  type ParagraphRange,
  type SelectionFormatState,
  type TextRange,
} from "../../lib/presentation/text";
import { isTextBodyFrame, type TextBodyFrame } from "../../lib/presentation/textBody";
import { effectiveHangingCm, effectiveIndentCm } from "../../lib/presentation/slideLists";
import { usePreferences } from "../../composables/usePreferences";
import { PRESENTATION_IMAGE_ACCEPT, PRESENTATION_IMAGE_MIME_TYPES } from "./useFrameInsertion";
import type { FrameSelection } from "./useFrameSelection";
import type { SlideTextEditing } from "./useSlideTextEditing";
import { paragraphSpan } from "./slideTextSession";
import { useSlideTextBody } from "./useSlideTextBody";
import { useSlideLinks } from "./useSlideLinks";
import {
  BULLET_CHAR_ERROR,
  BULLET_IMAGE_READ_ERROR,
  BULLET_IMAGE_TYPE_ERROR,
  foldParagraphs,
  numberOrBlank,
  paragraphsOf,
  type FrameMeasurer,
} from "./slideTextDepthFields";

export {
  BULLET_CHAR_ERROR,
  BULLET_IMAGE_TYPE_ERROR,
  LINK_SLIDE_ERROR,
  LINK_URL_ERROR,
  SLIDE_LINK_FOLLOW_EVENT,
  type FrameFit,
  type SlideLinkChoice,
  type SlideLinkDialogState,
} from "./slideTextDepthFields";

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
  /** Measures a rendered text body. Defaults to the stage DOM measurer
   * (slideTextFit.ts); injected by unit tests — jsdom cannot lay text out. */
  measureFit?: FrameMeasurer;
}

/** The Text panel's line-spacing modes (CustomSpacingDialog's vocabulary). */
export type LineSpacingMode = LineSpacingChoice["mode"];

export function useSlideTextDepth(deps: SlideTextDepthDeps) {
  // --- the Text panel's open state --------------------------------------------
  const { preferences, setPreference } = usePreferences();
  const panelOpen = computed(() => preferences.presentationTextOpen);
  function togglePanel(): void {
    setPreference("presentationTextOpen", !panelOpen.value);
  }

  // --- targets ----------------------------------------------------------------
  /** The one selected frame (a multi-selection targets nothing here). */
  const soleSelected = computed<SlideFrame | null>(() =>
    deps.selection.selectedFrameIds.value.length === 1 ? deps.selection.selectedFrame.value : null,
  );
  /** The frame the body controls edit: the frame being edited, else the one
   * selected text/shape frame. Null for a table-cell session. */
  const bodyFrame = computed<TextBodyFrame | null>(() => {
    const editingId = deps.editing.editingFrameId.value;
    const frame = editingId ? (deps.currentSlide.value?.frames.find((f) => f.id === editingId) ?? null) : soleSelected.value;
    return isTextBodyFrame(frame) ? frame : null;
  });

  /** The selection's paragraph/run formatting (null when no text is targeted). */
  const formatState = computed<SelectionFormatState | null>(() => deps.editing.toolbarState.value);
  const paragraphsEnabled = computed(() => deps.editing.toolbarEnabled.value);

  const body = useSlideTextBody({
    model: deps.model,
    currentSlide: deps.currentSlide,
    record: deps.record,
    commit: deps.commit,
    bodyFrame,
    measureFit: deps.measureFit,
  });

  // --- command runners ---------------------------------------------------------
  // Any text edit the session commits refits the edited frame, inside that
  // edit — a "resizeShape" frame tracks its text as it is typed (AC-38).
  deps.editing.setAfterEdit((frameId) => void body.refit(frameId));
  if (getCurrentScope()) onScopeDispose(() => deps.editing.setAfterEdit(null));
  function paragraphCommand(label: string, apply: (owner: ParagraphOwner, range: ParagraphRange) => void): void {
    deps.editing.runParagraphCommand(label, apply);
  }
  function textCommand(label: string, apply: (owner: ParagraphOwner, range: TextRange) => void, range?: TextRange): void {
    deps.editing.runTextCommand(label, apply, range);
  }

  const links = useSlideLinks({
    model: deps.model,
    currentSlide: deps.currentSlide,
    currentIndex: deps.currentIndex,
    selection: deps.selection,
    editing: deps.editing,
    soleSelected,
    textCommand,
    frameCommand: body.frameCommand,
  });

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
   * they differ. Re-read whenever the selection state is. */
  const listGeometry = computed<{ indentCm: number | ""; hangingCm: number | "" }>(() => {
    const target = formatState.value ? deps.editing.commandTarget() : null;
    if (!target) return { indentCm: "", hangingCm: "" };
    const span = paragraphSpan(target.range);
    return {
      indentCm: foldParagraphs(target.owner, span, effectiveIndentCm),
      hangingCm: foldParagraphs(target.owner, span, effectiveHangingCm),
    };
  });

  // --- list commands ------------------------------------------------------------
  /** The numbered toggle (slides-numbering): "1." for every touched paragraph,
   * or off when they all already are numbered. */
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
   * naming rule (`ppt/media/image{n}.{ext}` / `Pictures/image{n}.{ext}`, n = 1
   * + the highest suffix) inside the same undo entry that points the touched
   * paragraphs at it. */
  async function setBulletImage(file: File): Promise<void> {
    const contentType = file.type;
    if (!(PRESENTATION_IMAGE_MIME_TYPES as readonly string[]).includes(contentType)) {
      listError.value = BULLET_IMAGE_TYPE_ERROR;
      return;
    }
    const target = deps.editing.commandTarget(); // before the read: no drift
    if (!target) return;
    let data: string;
    try {
      data = bytesToBase64(new Uint8Array(await file.arrayBuffer()));
    } catch {
      listError.value = BULLET_IMAGE_READ_ERROR;
      return;
    }
    if (deps.editing.commandTarget()?.frame.id !== target.frame.id) return;
    listError.value = "";
    const mint = (owner: ParagraphOwner, range: TextRange): void => {
      const model = deps.model.value;
      if (!model) return;
      const mediaId = mintMediaId(model, contentType);
      // A fresh array reference: the stage's media resolver (and the store's
      // media record) pick the new part up on this commit.
      model.media = [...(model.media ?? []), { mediaId, contentType, data }];
      setListStyle(owner, paragraphSpan(range), { bulletImageId: mediaId });
    };
    textCommand("List style", mint, target.range);
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
  /** Switching to a point mode keeps a known point value, else starts at 1.2 ×
   * the run size; switching to "multiple" keeps a known multiple, else 1. */
  function setLineSpacingMode(mode: LineSpacingMode): void {
    const current = lineSpacingChoice.value;
    const known = current && current !== "mixed" ? current : null;
    const size = formatState.value?.sizePt;
    const choice: LineSpacingChoice =
      mode === "multiple"
        ? { mode, value: known?.mode === "multiple" ? known.value : 1 }
        : { mode, pt: known && known.mode !== "multiple" ? known.pt : Math.round((typeof size === "number" ? size : 18) * 1.2) };
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
  /** Letter spacing over the selected characters (a caret: its paragraph). */
  function setLetterSpacingPt(pt: number): void {
    const target = deps.editing.commandTarget();
    if (!Number.isFinite(pt) || !target) return;
    const value = clampLetterSpacing(pt);
    textCommand("Letter spacing", (owner, range) => setLetterSpacing(owner, range, value), paragraphsOf(target.owner, target.range));
  }

  return {
    // panel
    panelOpen,
    togglePanel,
    paragraphsEnabled,
    formatState,
    lineSpacingMode,
    lineSpacingValue,
    spaceBeforePt,
    spaceAfterPt,
    letterSpacingPt,
    startAt,
    listGeometry,
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
    // body (useSlideTextBody) and links (useSlideLinks)
    ...body,
    ...links,
  };
}

export type SlideTextDepth = ReturnType<typeof useSlideTextDepth>;
