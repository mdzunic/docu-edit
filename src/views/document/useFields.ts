// SPEC-053 field editing wired into DocumentView: the Insert → Field…,
// Insert → Date and time and Insert → Cross-reference dialogs, the F9 /
// "Update all fields" update paths, the field context-menu/chip actions
// (update, edit, toggle codes, unlink, lock), and click-to-follow for a
// cross-reference. Model logic lives in lib/commands.fields.ts and lib/fields.ts;
// this composable owns only the view wiring, per the DocumentView convention.
import { computed, ref, shallowRef } from "vue";
import type { EditorContext } from "./context";
import { caretSelection, documentStart, insertText } from "../../lib/commands";
import {
  insertFieldCommand,
  setFieldLockedAtCaret,
  unlinkFieldAtCaret,
  updateAllFields,
  updateFieldAtCaret,
  type FieldUpdateContext,
} from "../../lib/commands.fields";
import type { DocumentModel, FieldSpec } from "../../lib/documentModel";
import type { ModelSelection } from "../../lib/commands";
import { DATE_TIME_CATALOGUE, evaluateField, renderDateTime } from "../../lib/fields";
import { caretFromDomPoint } from "../../lib/domCaret";
import { now } from "../../lib/clock";
import { effectiveLanguage } from "../../lib/core/language";
import {
  bookmarkCaretIncludingHidden,
  ensureHeadingBookmark,
  headingAnchor,
  listBookmarks,
  listHeadings,
} from "../../lib/links";

export interface FieldsDeps extends EditorContext {
  /** The heading/bookmark → page-number resolver the view already owns
     (shared with the TOC update path). */
  tocPageNumberFor: (model: DocumentModel, anchor: string) => number | null;
  /** The shared context-menu opener/closer (from useContextMenus), so the
     field section can piggy-back on the same menu (SPEC-053 AC-13). */
  handleContextMenu: (event: MouseEvent) => void;
  closeContextMenu: () => void;
}

export function useFields(deps: FieldsDeps) {
  const fieldDialogOpen = ref(false);
  const dateTimeDialogOpen = ref(false);
  const crossReferenceDialogOpen = ref(false);
  const nowMs = ref(0);

  function model(): DocumentModel | null {
    return deps.doc.value;
  }

  /** The live selection to act on. `deps.currentSelection` needs the canvas
     root to read the DOM selection; called with none it always returns null
     (its `root` default is a typed-null seed), which silently no-oped every
     insert/update path. Passing the root and falling back to the tracked
     caret matches every other composable (useLinks, useTables, …). */
  function activeSelection(): ModelSelection {
    const root = deps.contentRoot.value as HTMLElement | null;
    return deps.currentSelection(root) ?? caretSelection(deps.getCaret() ?? documentStart());
  }

  function fileName(): string {
    return model()?.title ?? "";
  }

  function documentLanguage(): string {
    return effectiveLanguage(undefined, model()?.defaultStyles.language);
  }

  /** The update context every command reads the clock and layout through. */
  function updateContext(): FieldUpdateContext {
    const m = model();
    return {
      now: now(),
      fileName: fileName(),
      pageOfBookmark: (name) => (m ? deps.tocPageNumberFor(m, name) : null),
    };
  }

  // --- Insert → Field… -------------------------------------------------------

  const fieldPreviews = shallowRef<Array<{ format: string; label: string; isTime: boolean }>>([]);
  const propertyNames = shallowRef<string[]>([]);

  function openFieldDialog() {
    if (!deps.editable.value || !model()) {
      return;
    }
    nowMs.value = now().getTime();
    propertyNames.value = (model()?.properties?.customProperties ?? []).map((property) => property.name);
    fieldDialogOpen.value = true;
  }

  function handleFieldInsert(field: FieldSpec) {
    const m = model();
    const selection = activeSelection();
    fieldDialogOpen.value = false;
    if (!m) {
      return;
    }
    const withCached = fillCached(field, m, selection);
    deps.applyEdit(insertFieldCommand(m, selection, withCached), selection, "insert field");
  }

  /** Fills an inserted field's `cached` with its value now (insert is one of
     the update occasions, SPEC-053 Recalculation rule 2). */
  function fillCached(field: FieldSpec, m: DocumentModel, selection: ModelSelection): FieldSpec {
    if (!("cached" in field)) {
      return field;
    }
    const value = evaluateField(field, {
      model: m,
      now: now(),
      fileName: fileName(),
      here: selection.anchor,
      pageOfBookmark: (name) => deps.tocPageNumberFor(m, name),
    });
    return { ...field, cached: value };
  }

  // --- Insert → Date and time ------------------------------------------------

  function openDateTimeDialog() {
    if (!deps.editable.value || !model()) {
      return;
    }
    const language = documentLanguage();
    const instant = now();
    nowMs.value = instant.getTime();
    fieldPreviews.value = DATE_TIME_CATALOGUE.map((entry) => ({
      format: entry.format,
      label: renderDateTime(entry.format, instant, language) ?? entry.format,
      isTime: entry.isTime,
    }));
    dateTimeDialogOpen.value = true;
  }

  function handleDateTimeInsert(payload: { format: string; updateAutomatically: boolean }) {
    const m = model();
    const selection = activeSelection();
    dateTimeDialogOpen.value = false;
    if (!m) {
      return;
    }
    const isTime = DATE_TIME_CATALOGUE.find((entry) => entry.format === payload.format)?.isTime ?? false;
    const rendered = renderDateTime(payload.format, now(), documentLanguage()) ?? "";
    if (!payload.updateAutomatically) {
      // AC-4: plain text, no field.
      deps.applyEdit(insertText(m, selection, rendered), selection, "insert date");
      return;
    }
    // AC-5: a live date/time field, its cached value the rendered instant.
    const field: FieldSpec = isTime
      ? { kind: "time", format: payload.format, cached: rendered }
      : { kind: "date", format: payload.format, cached: rendered };
    deps.applyEdit(insertFieldCommand(m, selection, field), selection, "insert date field");
  }

  // --- Insert → Cross-reference ----------------------------------------------

  const crossReferenceHeadings = shallowRef<Array<{ index: number; text: string; level: number }>>([]);
  const crossReferenceBookmarks = shallowRef<Array<{ name: string }>>([]);

  function openCrossReferenceDialog() {
    const m = model();
    if (!deps.editable.value || !m) {
      return;
    }
    crossReferenceHeadings.value = listHeadings(m).map((heading, index) => ({
      index,
      text: heading.text,
      level: heading.level,
    }));
    crossReferenceBookmarks.value = listBookmarks(m).map((bookmark) => ({ name: bookmark.name }));
    crossReferenceDialogOpen.value = true;
  }

  function handleCrossReferenceInsert(payload: { targetType: string; targetId: string; show: string }) {
    let m = model();
    const selection = activeSelection();
    crossReferenceDialogOpen.value = false;
    if (!m) {
      return;
    }
    let targetName = payload.targetId;
    if (payload.targetType === "heading") {
      const heading = listHeadings(m)[Number(payload.targetId)];
      if (!heading) {
        return;
      }
      const ensured = ensureHeadingBookmark(m, heading.address, headingAnchor(heading.text, Number(payload.targetId)));
      m = ensured.model;
      targetName = ensured.name;
    }
    const field: FieldSpec =
      payload.show === "page"
        ? { kind: "pageRef", target: targetName, show: "page", cached: "" }
        : { kind: "ref", target: targetName, show: payload.show === "position" ? "position" : "text", cached: "" };
    const withCached = fillCached(field, m, selection);
    deps.applyEdit(insertFieldCommand(m, selection, withCached), selection, "insert cross-reference");
  }

  // --- Update paths ----------------------------------------------------------

  /** F9 (AC-15): update the field at the caret. */
  function handleUpdateField() {
    const m = model();
    const selection = activeSelection();
    if (!m || !deps.editable.value) {
      return;
    }
    deps.applyEdit(updateFieldAtCaret(m, selection, updateContext()), selection, "update field");
  }

  /** Tools → Update all fields (AC-16..18, AC-28). */
  function handleUpdateAllFields() {
    const m = model();
    const selection = activeSelection();
    if (!m || !deps.editable.value) {
      return;
    }
    const outcome = updateAllFields(m, selection, updateContext(), (anchor) => deps.tocPageNumberFor(m, anchor));
    deps.applyEdit({ model: outcome.model, selection: outcome.selection }, selection, "update all fields");
    if (outcome.brokenRefs.length > 0) {
      deps.addToast(`Updated fields. ${outcome.brokenRefs.length} cross-reference(s) point to a deleted target.`);
    } else {
      deps.addToast("Updated all fields");
    }
  }

  // --- Context-menu / chip actions ------------------------------------------

  const showFieldCodes = ref(false);

  function handleUnlinkField() {
    const m = model();
    const selection = activeSelection();
    if (!m || !deps.editable.value) {
      return;
    }
    deps.applyEdit(unlinkFieldAtCaret(m, selection, updateContext()), selection, "unlink field");
  }

  function handleLockField(locked: boolean) {
    const m = model();
    const selection = activeSelection();
    if (!m || !deps.editable.value) {
      return;
    }
    deps.applyEdit(setFieldLockedAtCaret(m, selection, locked), selection, locked ? "lock field" : "unlock field");
  }

  function toggleFieldCodes() {
    showFieldCodes.value = !showFieldCodes.value;
  }

  // The right-click field target (AC-13): whether the context menu opened on a
  // field, and whether that field is locked (Lock vs. Unlock).
  const contextMenuOnField = ref(false);
  const contextMenuFieldLocked = ref(false);

  /** Right-click handler: flags a field target, then delegates to the shared
     context-menu opener (which also places the caret on the atom). */
  function onContextMenu(event: MouseEvent) {
    const el = event.target instanceof Element ? event.target.closest(".doc-field") : null;
    contextMenuOnField.value = !!el;
    contextMenuFieldLocked.value = el?.classList.contains("doc-field-locked") ?? false;
    deps.handleContextMenu(event);
  }

  /** Dispatches a field context-menu / toolbar-chip action (AC-13, AC-14). */
  function handleFieldAction(action: string) {
    if (action === "update") {
      handleUpdateField();
    } else if (action === "edit") {
      openFieldDialog();
    } else if (action === "toggle-codes") {
      toggleFieldCodes();
    } else if (action === "unlink") {
      handleUnlinkField();
    } else if (action === "lock") {
      handleLockField(true);
    } else if (action === "unlock") {
      handleLockField(false);
    }
    deps.closeContextMenu();
  }

  /** Clicking a field selects its whole one-character atom (AC-7). A field
     renders as a `contenteditable=false` atom whose value can fill its line, so
     a plain click never lands the browser caret on it — the selection stays
     wherever it was (often the document start). We map the clicked element back
     to its model range and select it, so the field shades (AC-10), the toolbar
     chip appears (AC-14) and typing replaces the whole field (AC-8).
     Cross-references are followed before this (AC-26). Returns true when it
     selected a field. */
  function selectField(target: EventTarget | null): boolean {
    const root = deps.contentRoot.value as HTMLElement | null;
    if (!root || !(target instanceof Element)) {
      return false;
    }
    const atom = target.closest(".doc-field");
    const parent = atom?.parentNode ?? null;
    if (!atom || !parent) {
      return false;
    }
    const index = Array.prototype.indexOf.call(parent.childNodes, atom);
    const start = caretFromDomPoint(root, parent, index);
    if (!start) {
      return false;
    }
    // A field is exactly one model character, so its atom spans [start, start+1].
    deps.applyCaretMove(start);
    deps.applyCaretMove({ ...start, offset: start.offset + 1 }, true);
    return true;
  }

  /** Click-to-follow a cross-reference (AC-26): scroll to its target. */
  function followCrossReference(target: string) {
    const m = model();
    if (!m) {
      return;
    }
    const caret = bookmarkCaretIncludingHidden(m, target);
    if (caret) {
      void deps.scrollModelCaretIntoView(caret, true);
    }
  }

  return {
    fieldDialogOpen,
    dateTimeDialogOpen,
    crossReferenceDialogOpen,
    nowMs,
    fieldPreviews,
    propertyNames,
    crossReferenceHeadings,
    crossReferenceBookmarks,
    showFieldCodes,
    documentLanguage: computed(() => documentLanguage()),
    openFieldDialog,
    handleFieldInsert,
    openDateTimeDialog,
    handleDateTimeInsert,
    openCrossReferenceDialog,
    handleCrossReferenceInsert,
    handleUpdateField,
    handleUpdateAllFields,
    handleUnlinkField,
    handleLockField,
    toggleFieldCodes,
    selectField,
    followCrossReference,
    contextMenuOnField,
    contextMenuFieldLocked,
    onContextMenu,
    handleFieldAction,
  };
}
