// Character and paragraph formatting (SPEC-012, SPEC-013, SPEC-014, and the
// list/tab/pagination commands layered on them): the toolbar/shortcut
// formatting commands, their dialogs' state, and the ruler's drag handler.
//
// DocumentView used to hold each of the three core commands
// (toggleCharacterFormat / applyRunFormatPatch / applyParagraphCommand) in
// TRIPLICATE — a body version, a header/footer version (SPEC-029 AC-4) and a
// note version (SPEC-031 AC-6) whose bodies differed only in which model they
// formatted, which selection they read, which commit path applied the result,
// which element they refocused and which pending-format store they used. Here
// each command exists once, dispatched over the active EditingSurface below;
// the surface table is the whole difference between the old triad arms.
import { ref } from "vue";
import type { Ref } from "vue";
import {
  adjustIndent,
  applyFormat,
  applyListDefinition,
  applyMultilevelPreset,
  applyParagraphFormat,
  caretSelection,
  clearFormatting,
  documentDefaultTabStopCm,
  formatState,
  INDENT_STEP_CM,
  isCollapsed,
  linkHeadingNumbering,
  noParagraphCommand,
  noPendingFormat,
  paragraphPaginationState,
  selectionListDefinitionLevels,
  selectionTabStops,
  setAlignment,
  setListMarker,
  setPageMargins,
  setTabStops,
  setTabStopsWithDefault,
  toggleFormat,
  toggleList,
  toggleSpaceBefore,
  toggleVerticalAlign,
  verticalAlignState,
  withHeaderFooterModel,
  withNoteModel,
} from "../../lib/commands";
import { setColumnBoundary } from "../../lib/columns";
import { loadRecentColors, noHighlight, pushRecentColor } from "../../lib/core/colors";
import type { ThemeColorRef } from "../../lib/core/theme";
import { asListLevelDefinitions, asListMarkerStyle, asListType } from "../../lib/core/lists";
import {
  asCharacterFormatKey,
  asParagraphAlignment,
  asTabStopAlignment,
  asTabStops,
} from "../../lib/documentModel";
import type { CommandResult, ModelSelection } from "../../lib/commands";
import type {
  ActiveNoteTarget,
  DocumentModel,
  ParagraphFormatting,
  PendingFormat,
} from "../../lib/documentModel";
import type { CaretPosition } from "../../lib/textEditing";
import type { EditorContext, HeaderFooterBridge } from "./context";

export interface FormattingSurfacesDeps extends EditorContext, HeaderFooterBridge {
  /** DocumentView's paragraphSelection always resolves to a selection. */
  paragraphSelection: (root?: HTMLElement | null) => ModelSelection;
  /** Live read/write of the view's pending collapsed-caret format
   * (SPEC-012 AC-9..11) — the body surface's pending-format store, the same
   * lambdas the clipboard/autocorrect/context-menu wirings receive. */
  getPendingFormat: () => PendingFormat | null;
  setPendingFormat: (next: PendingFormat | null) => void;
  updateCaretStatus: (range?: { start: CaretPosition; end: CaretPosition }) => void;
  // --- the header/footer surface's remaining bundle (SPEC-029 AC-4);
  // activeHeaderFooter / currentHfSelection / applyHeaderFooterEdit ride in
  // via HeaderFooterBridge -------------------------------------------------
  hfRootEl: () => HTMLElement | null;
  getHfPendingFormat: () => PendingFormat | null;
  setHfPendingFormat: (next: PendingFormat | null) => void;
  // --- the note surface's bundle (SPEC-031 AC-6) --------------------------
  activeNote: Ref<ActiveNoteTarget | null>;
  currentNoteSelection: () => ModelSelection | null;
  applyNoteEdit: (
    resultModel: DocumentModel,
    resultSelection: ModelSelection,
    label: string,
  ) => void;
  noteRootEl: () => HTMLElement | null;
  getNotePendingFormat: () => PendingFormat | null;
  setNotePendingFormat: (next: PendingFormat | null) => void;
  // --- toolbar readouts the handlers below read (owned by the view, written
  // by its updateCaretStatus) — getters over refs, read live at call time --
  getParagraphState: () => { alignment: string; spaceBefore: boolean; direction: string };
  getParagraphFormatting: () => ParagraphFormatting;
}

/** One editing surface a formatting command can act on: the body, the active
 * header/footer band, or the active note. Encodes exactly what the old
 * per-surface command triads hard-coded. */
interface EditingSurface {
  /** The old guards' per-surface bail, re-read at call time like the
   * originals: the body's `!root` (its canvas ref can be unmounted), the
   * band's `!active`, the note's `!active`. */
  available(): boolean;
  /** The model formatting commands run against (identity for body;
   * the synthetic single-surface model for a header/footer band or note). */
  resolveModel(model: DocumentModel): DocumentModel;
  selection(): ModelSelection | null;
  /** Feeds a command result back to this surface's commit path. */
  apply(result: CommandResult, before: ModelSelection, label: string): void;
  rootEl(): HTMLElement | null;
  getPendingFormat(): PendingFormat | null;
  setPendingFormat(next: PendingFormat | null): void;
  /** Body only: refresh the toolbar readouts after a pending-format change. */
  pendingFormatChanged(selection: ModelSelection): void;
  /** Paragraph-command fallback when selection() is null. */
  fallbackSelection(): ModelSelection;
  /** Whether applyParagraphCommand honors its `refocus` flag (body) or
   * always refocuses (header/footer and note — their current behavior). */
  honorsRefocus: boolean;
}

export function useFormattingSurfaces(deps: FormattingSurfacesDeps) {
  const {
    doc,
    contentRoot,
    editable,
    currentSelection,
    applyEdit,
    paragraphSelection,
    getPendingFormat,
    setPendingFormat,
    updateCaretStatus,
    activeHeaderFooter,
    currentHfSelection,
    applyHeaderFooterEdit,
    hfRootEl,
    getHfPendingFormat,
    setHfPendingFormat,
    activeNote,
    currentNoteSelection,
    applyNoteEdit,
    noteRootEl,
    getNotePendingFormat,
    setNotePendingFormat,
    getParagraphState,
    getParagraphFormatting,
  } = deps;

  // --- dialog and picker state -------------------------------------------------
  // The custom-spacing dialog (SPEC-014 AC-12..14): open flag plus the prefill
  // taken from the first touched paragraph at the moment it opens.
  const customSpacingOpen = ref(false);
  const customSpacingInitial = ref({ mode: "multiple", value: 1, beforePt: 0, afterPt: 0 });
  // --- multilevel lists and tab stops (SPEC-032) ------------------------------
  // The "Define new multilevel list…" dialog: open flag plus the nine level
  // definitions it prefills from (the selection's own definition, or the
  // decimal preset).
  const defineListOpen = ref(false);
  const defineListInitial = ref(asListLevelDefinitions([]));
  // The Tabs dialog (AC-25): open flag plus the selection's current stops and
  // the document's default tab interval (SPEC-054 AC-10).
  const tabsDialogOpen = ref(false);
  const tabsDialogInitial = ref(asTabStops([]));
  const tabsDialogDefaultInterval = ref(1.27);
  // Recently applied text colours, shared across documents (SPEC-013 AC).
  const recentColors = ref(loadRecentColors());

  // --- the editing-surface table ----------------------------------------------
  // Each surface maps the old triad arm it replaces:
  // - body: the plain toggleCharacterFormat / applyRunFormatPatch /
  //   applyParagraphCommand path — identity model, currentSelection, the full
  //   applyEdit pipeline (undo, track changes), the view's pendingFormat store,
  //   and updateCaretStatus after a collapsed-caret pending change.
  // - headerFooter (SPEC-029 AC-4, the old toggleHfCharacterFormat /
  //   applyHfRunFormatPatch / applyHfParagraphCommand): identical logic against
  //   the active band's synthetic model (withHeaderFooterModel) and its own
  //   pending-format/caret state, committed through applyHeaderFooterEdit —
  //   the same chokepoint alignment/indent/list/style toolbar actions run
  //   through for body content, redirected at the band.
  // - note (SPEC-031 AC-6, the old toggleNoteCharacterFormat /
  //   applyNoteRunFormatPatch / applyNoteParagraphCommand): the header/footer
  //   arms' note counterpart over withNoteModel and applyNoteEdit.

  function bodyRootEl(): HTMLElement | null {
    return contentRoot.value as HTMLElement | null;
  }

  const bodySurface: EditingSurface = {
    available: () => !!bodyRootEl(),
    resolveModel: (model) => model,
    selection: () => currentSelection(bodyRootEl()),
    apply: (result, before, label) => applyEdit(result, before, label),
    rootEl: bodyRootEl,
    getPendingFormat,
    setPendingFormat,
    pendingFormatChanged: (selection) =>
      updateCaretStatus({ start: selection.anchor, end: selection.focus }),
    fallbackSelection: () => paragraphSelection(bodyRootEl()),
    honorsRefocus: true,
  };

  const headerFooterSurface: EditingSurface = {
    available: () => !!activeHeaderFooter.value,
    resolveModel: (model) => {
      const active = activeHeaderFooter.value;
      return active
        ? withHeaderFooterModel(model, active.sectionIndex, active.part, active.kind)
        : model;
    },
    selection: () => currentHfSelection(),
    apply: (result, _before, label) => applyHeaderFooterEdit(result.model, result.selection, label),
    rootEl: hfRootEl,
    getPendingFormat: getHfPendingFormat,
    setPendingFormat: setHfPendingFormat,
    pendingFormatChanged: () => {},
    fallbackSelection: () => caretSelection({ section: 0, block: 0, offset: 0 }),
    honorsRefocus: false,
  };

  const noteSurface: EditingSurface = {
    available: () => !!activeNote.value,
    resolveModel: (model) => {
      const active = activeNote.value;
      return active ? withNoteModel(model, active.noteId) : model;
    },
    selection: () => currentNoteSelection(),
    apply: (result, _before, label) => applyNoteEdit(result.model, result.selection, label),
    rootEl: noteRootEl,
    getPendingFormat: getNotePendingFormat,
    setPendingFormat: setNotePendingFormat,
    pendingFormatChanged: () => {},
    fallbackSelection: () => caretSelection({ section: 0, block: 0, offset: 0 }),
    honorsRefocus: false,
  };

  /** The surface the next formatting command routes to — a live header/footer
   * band first, else a live note, else the body: the same precedence order
   * every old per-command dispatcher checked in (activeHeaderFooter, then
   * activeNote). The two can never actually be live at once — entering a note
   * exits the band and vice versa — so the order is a tiebreak in name only. */
  function activeSurface(): EditingSurface {
    if (activeHeaderFooter.value) {
      return headerFooterSurface;
    }
    if (activeNote.value) {
      return noteSurface;
    }
    return bodySurface;
  }

  // --- character formatting toolbar/shortcuts (SPEC-012, SPEC-013) -----------

  /** Bold/italic/underline/strikethrough (SPEC-012 AC-3..6, AC-16, AC-17)
   * and subscript/superscript (SPEC-013): a real selection is formatted
   * through the command API (which keeps the selection live, see applyEdit);
   * a collapsed caret instead sets/flips the pending format the next typed
   * text will carry (AC-9, AC-10). */
  function toggleCharacterFormat(key = "") {
    const surface = activeSurface();
    const model = doc.value;
    if (!surface.available() || !model || !editable.value) {
      return;
    }
    const selection = surface.selection();
    if (!selection) {
      return;
    }
    const target = surface.resolveModel(model);
    if (key === "subscript" || key === "superscript") {
      // One vertAlign field makes the pair mutually exclusive (SPEC-013).
      if (isCollapsed(selection)) {
        const on = verticalAlignState(target, selection, key, surface.getPendingFormat()) === "on";
        surface.setPendingFormat({ ...surface.getPendingFormat(), vertAlign: on ? false : key });
        surface.pendingFormatChanged(selection);
      } else {
        surface.apply(toggleVerticalAlign(target, selection, key), selection, key);
      }
      surface.rootEl()?.focus();
      return;
    }
    const formatKey = asCharacterFormatKey(key);
    if (isCollapsed(selection)) {
      const next = formatState(target, selection, formatKey, surface.getPendingFormat()) !== "on";
      surface.setPendingFormat({ ...surface.getPendingFormat(), [formatKey]: next });
      surface.pendingFormatChanged(selection);
    } else {
      surface.apply(toggleFormat(target, selection, formatKey), selection, formatKey);
    }
    surface.rootEl()?.focus();
  }

  /** Applies a value-format patch (font family/size, colour, highlight —
   * SPEC-013): a real selection goes through the command API; a collapsed
   * caret merges the patch into the pending format the next typed text
   * carries, exactly like the boolean toggles. */
  function applyRunFormatPatch(patch = noPendingFormat(), label = "formatting") {
    const surface = activeSurface();
    const model = doc.value;
    if (!patch || !surface.available() || !model || !editable.value) {
      return;
    }
    const selection = surface.selection();
    if (!selection) {
      return;
    }
    const target = surface.resolveModel(model);
    if (isCollapsed(selection)) {
      surface.setPendingFormat({ ...surface.getPendingFormat(), ...patch });
      surface.pendingFormatChanged(selection);
    } else {
      surface.apply(applyFormat(target, selection, patch), selection, label);
    }
    surface.rootEl()?.focus();
  }

  function handleFontFamilySelect(family = "") {
    if (family !== "") {
      applyRunFormatPatch({ fontFamily: family }, "font family");
    }
  }

  function handleFontSizeChange(sizePt = 0) {
    if (sizePt > 0) {
      applyRunFormatPatch({ fontSizePt: sizePt }, "font size");
    }
  }

  /** The picker emits a "#rrggbb" literal or a ThemeColorRef (SPEC-077
   * AC-23) — the reference is stored as-is, and only literals enter the
   * recent list (AC-24: recents hold literals, never references). */
  function handleTextColorSelect(color: string | ThemeColorRef = "") {
    if (color !== "") {
      applyRunFormatPatch({ color }, "text color");
      if (typeof color === "string") {
        recentColors.value = pushRecentColor(color, recentColors.value);
      }
    }
  }

  /** The highlight picker emits a palette name, false for "No colour" —
   * an explicit off, not an absence (SPEC-013) — or a ThemeColorRef
   * (SPEC-077 AC-23), stored as the reference itself. */
  function handleHighlightSelect(value: string | false | ThemeColorRef = noHighlight()) {
    applyRunFormatPatch({ highlight: value }, "highlight");
  }

  // --- paragraph formatting toolbar/ruler/shortcuts (SPEC-014) ----------------

  /** Runs one paragraph command against the current selection and applies its
   * result; `label` names the edit for the undo history (SPEC-015 AC-7, AC-8,
   * AC-13 — every paragraph command is its own undo step); `refocus` returns
   * focus to the canvas (skipped for ruler drags, which would tear down the
   * pointer capture mid-drag) — honored by the body surface only: an active
   * band or note always takes focus back, as their old dispatch arms did. */
  function applyParagraphCommand(command = noParagraphCommand(), label = "formatting", refocus = true) {
    const surface = activeSurface();
    const model = doc.value;
    if (!command || !surface.available() || !model || !editable.value) {
      return;
    }
    const selection = surface.selection() ?? surface.fallbackSelection();
    const target = surface.resolveModel(model);
    surface.apply(command(target, selection), selection, label);
    if (refocus || !surface.honorsRefocus) {
      surface.rootEl()?.focus();
    }
  }

  function handleAlign(alignment = "") {
    applyParagraphCommand(
      (model, selection) => setAlignment(model, selection, asParagraphAlignment(alignment)),
      "alignment",
    );
  }

  function handleIndent(direction = 1) {
    applyParagraphCommand(
      (model, selection) => adjustIndent(model, selection, direction * INDENT_STEP_CM),
      "indent",
    );
  }

  // --- clear formatting (SPEC-020) --------------------------------------------

  /** The "Clear formatting" toolbar button and Ctrl/Cmd+\ (AC-1, AC-2, AC-3):
   * one command result, so — like every other paragraph command — it is one
   * undo step (AC-18). */
  function handleClearFormatting() {
    applyParagraphCommand(clearFormatting, "clear formatting");
  }

  /** A line-spacing preset (Single/1.15/1.5/Double) sets the multiple and
   * clears any absolute height so the preset actually takes effect (AC-11). */
  function handleSpacingPreset(multiple = 1) {
    applyParagraphCommand(
      (model, selection) =>
        applyParagraphFormat(model, selection, {
          lineSpacing: multiple,
          lineSpacingPt: null,
          lineSpacingRule: null,
        }),
      "line spacing",
    );
  }

  function handleToggleSpaceBefore() {
    applyParagraphCommand(toggleSpaceBefore, "space before");
  }

  /** The bulleted/numbered toolbar toggles (SPEC-018 AC-1, AC-2). */
  function handleToggleList(type = "") {
    applyParagraphCommand(
      (model, selection) => toggleList(model, selection, asListType(type)),
      type === "bullet" ? "bulleted list" : "numbered list",
    );
  }

  /** A marker style picked from a list button's dropdown (SPEC-018 AC-13). */
  function handleListMarker(style = "") {
    applyParagraphCommand(
      (model, selection) => setListMarker(model, selection, asListMarkerStyle(style)),
      "list marker",
    );
  }

  // --- multilevel lists and custom tab stops (SPEC-032) -----------------------

  /** A gallery preset picked from the Multilevel list menu (AC-1..4). */
  function handleMultilevelPreset(key = "") {
    applyParagraphCommand(
      (model, selection) => applyMultilevelPreset(model, selection, key),
      "multilevel list",
    );
  }

  /** "Number headings 1–3" (AC-13). */
  function handleLinkHeadingNumbering() {
    applyParagraphCommand(
      (model, selection) => linkHeadingNumbering(model, selection),
      "number headings",
    );
  }

  /** "Define new multilevel list…" (AC-5..12): prefilled from the selection's
   * own definition when it sits in a defined list. */
  function openDefineList() {
    const root = contentRoot.value as HTMLElement | null;
    const model = doc.value;
    if (!root || !model) {
      return;
    }
    defineListInitial.value = selectionListDefinitionLevels(model, paragraphSelection(root));
    defineListOpen.value = true;
  }

  function handleDefineListApply(levels = asListLevelDefinitions([])) {
    defineListOpen.value = false;
    applyParagraphCommand(
      (model, selection) => applyListDefinition(model, selection, asListLevelDefinitions(levels)),
      "list definition",
    );
  }

  /** The Tabs dialog (AC-25), prefilled from the first touched paragraph. */
  function openTabsDialog() {
    const root = contentRoot.value as HTMLElement | null;
    const model = doc.value;
    if (!root || !model) {
      return;
    }
    tabsDialogInitial.value = selectionTabStops(model, paragraphSelection(root));
    tabsDialogDefaultInterval.value = documentDefaultTabStopCm(model);
    tabsDialogOpen.value = true;
  }

  function handleTabsApply(payload = { stops: asTabStops([]), defaultIntervalCm: 1.27 }) {
    tabsDialogOpen.value = false;
    const stops = asTabStops(payload.stops ?? []);
    const defaultIntervalCm = payload.defaultIntervalCm ?? 1.27;
    applyParagraphCommand(
      (model, selection) => setTabStopsWithDefault(model, selection, stops, defaultIntervalCm),
      "tab stops",
    );
  }

  function handleToggleDirection() {
    const rtl = getParagraphState().direction !== "rtl";
    applyParagraphCommand(
      (model, selection) => applyParagraphFormat(model, selection, { direction: rtl ? "rtl" : "ltr" }),
      "text direction",
    );
  }

  function openCustomSpacing() {
    const formatting = getParagraphFormatting();
    customSpacingInitial.value =
      formatting.lineSpacingPt !== undefined
        ? {
            mode: formatting.lineSpacingRule === "exact" ? "exact" : "atLeast",
            value: formatting.lineSpacingPt,
            beforePt: formatting.spaceBeforePt ?? 0,
            afterPt: formatting.spaceAfterPt ?? 0,
          }
        : {
            mode: "multiple",
            value: formatting.lineSpacing ?? 1,
            beforePt: formatting.spaceBeforePt ?? 0,
            afterPt: formatting.spaceAfterPt ?? 0,
          };
    customSpacingOpen.value = true;
  }

  /** OK in the custom-spacing dialog (AC-13): one patch, applied to every
   * touched paragraph. The unused line-spacing representation is cleared so
   * the chosen mode wins regardless of what was set before. */
  function handleCustomSpacingApply(payload = { mode: "multiple", value: 1, beforePt: 0, afterPt: 0 }) {
    customSpacingOpen.value = false;
    applyParagraphCommand(
      (model, selection) =>
        payload.mode === "multiple"
          ? applyParagraphFormat(model, selection, {
              lineSpacing: payload.value,
              lineSpacingPt: null,
              lineSpacingRule: null,
              spaceBeforePt: payload.beforePt,
              spaceAfterPt: payload.afterPt,
            })
          : applyParagraphFormat(model, selection, {
              lineSpacing: null,
              lineSpacingPt: payload.value,
              lineSpacingRule: payload.mode === "exact" ? "exact" : null,
              spaceBeforePt: payload.beforePt,
              spaceAfterPt: payload.afterPt,
            }),
      "paragraph spacing",
    );
  }

  /** A ruler marker drag (AC-23): indent updates go to every touched
   * paragraph as logical values; margin updates go to the page margins. No
   * refocus — grabbing focus mid-drag would break the pointer capture. */
  function handleRulerUpdate(
    update = { kind: "", valueCm: 0, boundary: 0, deltaCm: 0, positionCm: 0, alignment: "left", index: 0 },
  ) {
    const valueCm = update.valueCm;
    if (update.kind === "indentStartCm") {
      applyParagraphCommand(
        (model, selection) => applyParagraphFormat(model, selection, { indentStartCm: valueCm }),
        "indent",
        false,
      );
    } else if (update.kind === "indentEndCm") {
      applyParagraphCommand(
        (model, selection) => applyParagraphFormat(model, selection, { indentEndCm: valueCm }),
        "indent",
        false,
      );
    } else if (update.kind === "indentFirstLineCm") {
      applyParagraphCommand(
        (model, selection) => applyParagraphFormat(model, selection, { indentFirstLineCm: valueCm }),
        "indent",
        false,
      );
    } else if (update.kind === "marginLeftCm") {
      applyParagraphCommand(
        (model, selection) => setPageMargins(model, selection, { leftCm: valueCm }),
        "margins",
        false,
      );
    } else if (update.kind === "marginRightCm") {
      applyParagraphCommand(
        (model, selection) => setPageMargins(model, selection, { rightCm: valueCm }),
        "margins",
        false,
      );
    } else if (update.kind === "column-boundary") {
      const boundary = update.boundary ?? 0;
      const deltaCm = update.deltaCm ?? 0;
      applyParagraphCommand(
        (model, selection) => setColumnBoundary(model, selection, boundary, deltaCm),
        "column width",
        false,
      );
    } else if (update.kind === "add-tab") {
      // A click in the ruler's tab lane lays a stop of the armed type
      // (SPEC-032 AC-24).
      const positionCm = update.positionCm ?? 0;
      const alignment = asTabStopAlignment(update.alignment ?? "left");
      applyParagraphCommand(
        (model, selection) =>
          setTabStops(model, selection, [
            ...selectionTabStops(model, selection),
            { positionCm, alignment },
          ]),
        "tab stops",
        false,
      );
    } else if (update.kind === "move-tab") {
      // Releasing a dragged handle on the ruler moves its stop live
      // (SPEC-054 AC-4). The dragged stop is identified by its index in the
      // first touched paragraph's list.
      const index = update.index ?? 0;
      const positionCm = update.positionCm ?? 0;
      applyParagraphCommand(
        (model, selection) => {
          const stops = selectionTabStops(model, selection);
          if (stops[index]) {
            stops[index] = { ...stops[index], positionCm };
          }
          return setTabStops(model, selection, stops);
        },
        "tab stops",
        false,
      );
    } else if (update.kind === "remove-tab") {
      // Clicking a laid handle, or dragging it off the ruler, removes its stop
      // (SPEC-032 AC-24, SPEC-054 AC-5).
      const index = update.index ?? 0;
      applyParagraphCommand(
        (model, selection) => {
          const stops = selectionTabStops(model, selection);
          stops.splice(index, 1);
          return setTabStops(model, selection, stops);
        },
        "tab stops",
        false,
      );
    }
  }

  /** The paragraph pagination menu's toggles (AC-22..28): each key flips
   * against the first touched paragraph's current state, applied to every
   * touched paragraph. */
  function handleTogglePagination(key = "") {
    applyParagraphCommand((model, selection) => {
      const state = paragraphPaginationState(model, selection);
      if (key === "keepNext") {
        return applyParagraphFormat(model, selection, { keepNext: !state.keepNext });
      }
      if (key === "keepLines") {
        return applyParagraphFormat(model, selection, { keepLines: !state.keepLines });
      }
      if (key === "pageBreakBefore") {
        return applyParagraphFormat(model, selection, { pageBreakBefore: !state.pageBreakBefore });
      }
      return applyParagraphFormat(model, selection, { widowControl: !state.widowControl });
    }, "paragraph pagination");
  }

  return {
    toggleCharacterFormat,
    applyRunFormatPatch,
    handleFontFamilySelect,
    handleFontSizeChange,
    handleTextColorSelect,
    handleHighlightSelect,
    recentColors,
    applyParagraphCommand,
    handleAlign,
    handleIndent,
    handleClearFormatting,
    handleSpacingPreset,
    handleToggleSpaceBefore,
    handleToggleList,
    handleListMarker,
    handleMultilevelPreset,
    handleLinkHeadingNumbering,
    defineListOpen,
    defineListInitial,
    openDefineList,
    handleDefineListApply,
    tabsDialogOpen,
    tabsDialogInitial,
    tabsDialogDefaultInterval,
    openTabsDialog,
    handleTabsApply,
    handleToggleDirection,
    customSpacingOpen,
    customSpacingInitial,
    openCustomSpacing,
    handleCustomSpacingApply,
    handleRulerUpdate,
    handleTogglePagination,
  };
}
