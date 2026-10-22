// The hyperlink half of SPEC-097's slide text depth (composed by
// useSlideTextDepth): the slide link dialog's state and commands, and the
// stage's Cmd/Ctrl+click navigation. Two entry points open the dialog —
// Cmd/Ctrl+K (or the text toolbar's `slides-link`) inside a text session links
// the selected characters (a caret widens to its word), and the insert bar's
// `slides-link` links a whole selected text/image/shape frame
// (`isLinkableKind`). Every target passes `cleanLinkTarget` (sanitizeHref
// inside), and applying/removing is ONE undo entry, "Add link"/"Remove link".
// A bubbling `slide-link-follow` CustomEvent (1-based `detail.slide`) from the
// stage moves the editor to that slide, stopping any session first.
import { computed, getCurrentScope, onScopeDispose, ref, type ComputedRef, type Ref } from "vue";
import type { PresentationModel, Slide, SlideFrame } from "../../lib/presentation/model";
import type { ParagraphOwner, TextRange } from "../../lib/presentation/text";
import { cleanLinkTarget, isLinkableKind, linkOf, setFrameLink, setRunLink, type SlideLinkTarget } from "../../lib/presentation/slideLinks";
import { slideOutline } from "../../lib/presentation/slideModel";
import type { FrameSelection } from "./useFrameSelection";
import type { SlideTextEditing } from "./useSlideTextEditing";
import {
  LINK_SLIDE_ERROR,
  LINK_URL_ERROR,
  SLIDE_LINK_FOLLOW_EVENT,
  runLinkAt,
  wordRange,
  type SlideLinkChoice,
  type SlideLinkDialogState,
} from "./slideTextDepthFields";

export interface SlideLinksDeps {
  model: Ref<PresentationModel | null>;
  currentSlide: ComputedRef<Slide | null>;
  currentIndex: Ref<number>;
  selection: FrameSelection;
  editing: SlideTextEditing;
  /** The one selected frame (null for none or a multi-selection). */
  soleSelected: ComputedRef<SlideFrame | null>;
  /** A character-range command through the session (useSlideTextDepth). */
  textCommand: (label: string, apply: (owner: ParagraphOwner, range: TextRange) => void, range?: TextRange) => void;
  /** A frame-level command (useSlideTextBody.frameCommand). */
  frameCommand: (label: string, frame: SlideFrame | null, mutate: (frame: SlideFrame) => void) => boolean;
}

export function useSlideLinks(deps: SlideLinksDeps) {
  const linkDialog = ref<SlideLinkDialogState | null>(null);

  function frameById(id: string): SlideFrame | null {
    return deps.currentSlide.value?.frames.find((frame) => frame.id === id) ?? null;
  }

  /** The one selected frame a whole-frame link applies to (text/image/shape),
   * only outside a text session (inside one, a link covers characters). */
  const linkableFrame = computed<SlideFrame | null>(() => {
    if (deps.editing.editingFrameId.value) return null;
    const frame = deps.soleSelected.value;
    return frame && isLinkableKind(frame.kind) ? frame : null;
  });
  /** Whether the text toolbar's link button shows (a text session is open). */
  const canLinkText = computed(() => deps.editing.editingFrameId.value !== null);

  const slideChoices = computed<SlideLinkChoice[]>(() =>
    (deps.model.value?.slides ?? []).map((slide, index) => ({
      number: index + 1,
      label: slideOutline(slide).title ?? `Slide ${index + 1}`,
    })),
  );

  /** Cmd/Ctrl+K / the text toolbar's button: in a session, link the selected
   * characters; otherwise the whole selected frame. */
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

  /** The insert bar's button: link the whole selected frame. */
  function openFrameLinkDialog(): void {
    const frame = linkableFrame.value;
    if (!frame) return;
    linkDialog.value = { kind: "frame", frameId: frame.id, range: null, initial: linkOf(frame) };
  }

  function closeLinkDialog(): void {
    linkDialog.value = null;
  }

  /** Applies the dialog's target ("Add link"). Returns an error message for an
   * unusable target (a dangerous URL, a slide outside the deck) and keeps the
   * dialog open; "" once applied. */
  function applyLink(target: SlideLinkTarget): string {
    const dialog = linkDialog.value;
    const model = deps.model.value;
    if (!dialog || !model) return "";
    const clean = cleanLinkTarget(target, model.slides.length);
    if (!clean) return target.slide !== undefined ? LINK_SLIDE_ERROR : LINK_URL_ERROR;
    if (dialog.kind === "run") {
      deps.textCommand("Add link", (owner, range) => setRunLink(owner, range, clean), dialog.range ?? undefined);
    } else {
      deps.frameCommand("Add link", frameById(dialog.frameId), (frame) => setFrameLink(frame, clean));
    }
    linkDialog.value = null;
    return "";
  }

  /** Removes the link under the dialog's range / from its frame ("Remove link"). */
  function removeLink(): void {
    const dialog = linkDialog.value;
    if (!dialog) return;
    if (dialog.kind === "run") {
      deps.textCommand("Remove link", (owner, range) => setRunLink(owner, range, null), dialog.range ?? undefined);
    } else {
      deps.frameCommand("Remove link", frameById(dialog.frameId), (frame) => setFrameLink(frame, null));
    }
    linkDialog.value = null;
  }

  // Cmd/Ctrl+K inside a session opens the dialog.
  deps.editing.setLinkShortcut(openLinkDialog);

  /** The stage's Cmd/Ctrl+click on an in-deck link: go to that slide. */
  function onLinkFollow(event: Event): void {
    const model = deps.model.value;
    const slide = event instanceof CustomEvent ? Number(event.detail?.slide) : Number.NaN;
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
