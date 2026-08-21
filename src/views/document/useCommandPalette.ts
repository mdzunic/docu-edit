// The command palette wiring for the editor (SPEC-062). DocumentView owns the
// registry (`editorCommands`) and every enablement predicate the menu bar
// already reads; this composable turns those into the three things
// CommandPalette.vue needs — the searchable command list, a live
// id→reason availability record, and the recents list — plus the open/close
// lifecycle that returns focus to where it was (AC-14) without ever touching
// the document selection (AC-15, AC-26).
//
// Kept out of DocumentView.vue to hold its line budget (DocumentView.size
// .spec.ts): the view gains only a `useCommandPalette({ ...editorContext,
// ...extras })` call and the template hookup.
import { computed, ref, type ComputedRef, type Ref } from "vue";
import {
  COMMANDS,
  COMMON_COMMAND_IDS,
  loadRecentCommands,
  pushRecentCommand,
  type CommandMeta,
} from "../../lib/commandCatalogue";
import { stylePickerEntries } from "../../lib/styles";
import type { EditorContext } from "./context";

/** The prefix marking a virtual "Apply style: …" row; the styleId follows. */
export const STYLE_ROW_PREFIX = "style-apply:";

export interface CommandPaletteDeps extends EditorContext {
  /** Whether the document is editable — the menu's `!editable` gate. */
  editable: Ref<boolean>;
  /** The undo/redo availability the toolbar/menu read. */
  undoRedo: Ref<{ canUndo: boolean; canRedo: boolean }>;
  /** Whether "Convert {{…}} text to merge fields" applies. */
  canConvertLiteralChips: Ref<boolean>;
  /** Whether a table holds the active selection ("Chart from table"). */
  tableSelectedForChart: Ref<boolean>;
  /** Whether an image/table/chart is selected ("Caption"). */
  captionTargetSelected: Ref<boolean>;
  /** The theme actually in effect — gates the dark-only "Dim page" item. */
  resolvedTheme: () => string;
  /** Dispatches a registry command by id — the exact call the menu bar's
   * handler makes, so a palette run is identical (AC-16). */
  dispatch: (id: string) => void;
  /** Applies a named style by id (the "Apply style: …" rows). */
  applyStyle: (styleId: string) => void;
}

/** The document commands that stay available even in a read-only document —
 * everything the menu bar renders without an `!editable` gate (print, view
 * toggles, the always-on info dialogs). Every other document command needs an
 * editable document, mirroring SPEC-049's menu. */
const ALWAYS_AVAILABLE = new Set([
  "file.print",
  "file.properties",
  "file.version-history",
  "file.close",
  "edit.copy",
  "edit.selectAll",
  "find.openReplace",
  "view.ruler",
  "view.outline-pane",
  "view.styles-pane",
  "view.comments-pane",
  "view.comments-only",
  "view.field-shading",
  "view.track-changes-pane",
  "view.formattingMarks",
  "view.theme-system",
  "view.theme-light",
  "view.theme-dark",
  "view.dim-page",
  "format.styles-pane",
  "insert.bookmark-list",
  "tools.word-count",
  "tools.dictionary",
  "tools.accessibility",
  "tools.fonts",
]);

export interface CommandPaletteApi {
  commandPaletteOpen: Ref<boolean>;
  paletteCommands: CommandMeta[];
  paletteStyleRows: ComputedRef<CommandMeta[]>;
  paletteAvailability: ComputedRef<Record<string, string>>;
  paletteRecentIds: Ref<string[]>;
  paletteCommonIds: string[];
  openCommandPalette: () => void;
  closeCommandPalette: () => void;
  runCommandFromPalette: (id: string) => void;
}

export function useCommandPalette(deps: CommandPaletteDeps): CommandPaletteApi {
  const commandPaletteOpen = ref(false);
  const paletteRecentIds = ref<string[]>(loadRecentCommands());

  // The document-scope catalogue is static; style rows are per open document.
  const paletteCommands = COMMANDS.filter((command) => command.scope === "document");
  const paletteCommonIds = COMMON_COMMAND_IDS;

  /** The open document's styles as virtual "Apply style: …" rows (AC-22,
   * AC-23), rebuilt live from the style catalogue — never persisted into
   * COMMANDS. */
  const paletteStyleRows = computed<CommandMeta[]>(() => {
    const model = deps.doc.value;
    if (!model) {
      return [];
    }
    return stylePickerEntries(model).map((entry) => ({
      id: `${STYLE_ROW_PREFIX}${entry.id}`,
      label: `Apply style: ${entry.name}`,
      path: ["Styles"],
      scope: "document" as const,
    }));
  });

  /** id → "" when enabled, or a short reason when not (AC-18, AC-19). Reads
   * the same predicates the menu bar does, so the two surfaces cannot
   * disagree. */
  const paletteAvailability = computed<Record<string, string>>(() => {
    const editable = deps.editable.value;
    const record: Record<string, string> = {};
    const readOnlyReason = "document is read-only";
    for (const command of paletteCommands) {
      const { id } = command;
      let reason = "";
      if (!ALWAYS_AVAILABLE.has(id) && !editable) {
        reason = readOnlyReason;
      } else if (id === "edit.undo" && !deps.undoRedo.value.canUndo) {
        reason = "nothing to undo";
      } else if (id === "edit.redo" && !deps.undoRedo.value.canRedo) {
        reason = "nothing to redo";
      } else if (id === "tools.convert-merge-fields" && !deps.canConvertLiteralChips.value) {
        reason = "no {{…}} text to convert";
      } else if (id === "insert.chart-from-table" && !deps.tableSelectedForChart.value) {
        reason = "no table selected";
      } else if (id === "insert.caption" && !deps.captionTargetSelected.value) {
        reason = "no caption target selected";
      } else if (id === "view.dim-page" && deps.resolvedTheme() !== "dark") {
        reason = "dark theme only";
      }
      if (reason) {
        record[id] = reason;
      }
    }
    // Style rows follow the document's editability.
    if (!editable) {
      for (const row of paletteStyleRows.value) {
        record[row.id] = readOnlyReason;
      }
    }
    return record;
  });

  // Focus is returned here on close (AC-14) — captured at open time.
  let restoreFocusTo: HTMLElement | null = null;

  function openCommandPalette(): void {
    if (commandPaletteOpen.value) {
      return;
    }
    const active = document.activeElement;
    restoreFocusTo = active instanceof HTMLElement ? active : null;
    commandPaletteOpen.value = true;
  }

  function closeCommandPalette(): void {
    if (!commandPaletteOpen.value) {
      return;
    }
    commandPaletteOpen.value = false;
    // Returning focus to the canvas restores its DOM selection, so the model
    // selection the palette never touched is usable again (AC-15).
    const target = restoreFocusTo;
    restoreFocusTo = null;
    if (target && typeof target.focus === "function") {
      target.focus();
    }
  }

  function runCommandFromPalette(id: string): void {
    if (id.startsWith(STYLE_ROW_PREFIX)) {
      const styleId = id.slice(STYLE_ROW_PREFIX.length);
      closeCommandPalette();
      deps.applyStyle(styleId);
      return;
    }
    // Record before dispatch so the reason a command may throw never loses the
    // recent; unknown ids (dropped on reload) are harmless.
    paletteRecentIds.value = pushRecentCommand(id, paletteRecentIds.value);
    closeCommandPalette();
    deps.dispatch(id);
  }

  return {
    commandPaletteOpen,
    paletteCommands,
    paletteStyleRows,
    paletteAvailability,
    paletteRecentIds,
    paletteCommonIds,
    openCommandPalette,
    closeCommandPalette,
    runCommandFromPalette,
  };
}
