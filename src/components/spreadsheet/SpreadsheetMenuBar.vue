<script setup lang="ts">
// The workbook's menu bar (SPEC-117 "The menu bar") — PresentationMenuBar's
// pattern for the spreadsheet product: the product-neutral MenuBar with
// slotted menus, every item emitting a command id that useSpreadsheetMenu
// routes, its state coerced by lib/spreadsheet/menuState.ts. SPEC-117's
// items: File — Download (the same action as the app bar's) and Close;
// View — Gridlines (a menuitemcheckbox), SPEC-119's Show formulas (another,
// with its Ctrl+` shortcut — Control on every platform) and Zoom ▸ the
// ZOOM_PRESETS ladder, Zoom in and Zoom out. SPEC-120's Data menu opens with
// Recalculate (F9).
// SPEC-121 adds the Edit menu between File and View — Undo and Redo, each
// naming its step ("Undo Rename sheet") with the registry's shortcut, and
// disabled when there is nothing to undo or redo — and File ▸ "Download a
// copy as …" after Download (disabled, with the reason, until the copy can be
// made). SPEC-123 adds the Format menu after View: Format cells…, Number
// format ▸ (the presets and More formats…), Cell styles ▸ and Clear ▸.
// SPEC-124 adds View ▸ Freeze ▸ (Top row, First column, Freeze at current
// cell — disabled at A1 — and Unfreeze, enabled only when frozen) and, at the
// Format menu's end, Row height…, Column width… and the two autofits.
// SPEC-125 adds Edit ▸ Paste special… and Edit ▸ Fill ▸ Down / Right, each
// with its SHEET_SHORTCUTS binding. SPEC-128 adds to the Data menu, after
// Recalculate and a separator, Sort A to Z, Sort Z to A and Sort range…, then
// after another separator Filter (a menuitemcheckbox, checked while the sheet
// has an autofilter, Ctrl/Cmd+Shift+L), Clear filters and Reapply filter
// (Ctrl/Cmd+Alt+L) — the last two disabled without an autofilter, stating
// "Turn on Filter first". SPEC-129 adds Format ▸ Conditional formatting…
// (after Cell styles ▸), which opens the side panel's rules manager. SPEC-130
// adds Data ▸ "Data validation…" after a separator below the filter items,
// which opens the validation dialog (disabled on a foreign sheet).
// SPEC-131 adds the Insert menu after View — Image… (`sheet-menu-insert-image`,
// the image picker) and Shapes (`sheet-menu-insert-shape`, the shape gallery
// popover), disabled where the workbook cannot hold drawings — and the
// Arrange menu after it: Bring to Front, Bring Forward, Send Backward, Send to
// Back (`sheet-menu-arrange-front`, `-forward`, `-backward`, `-back`) and
// Alt Text… (`sheet-menu-arrange-alt-text`) on the selected drawing, disabled
// without one (and a move that cannot move it disabled). The drawing context
// menu offers the same under `sheet-drawing-*` ids; these popups stay mounted,
// so the two never share an id. SPEC-133 ends the Insert menu, after a
// separator, with Table… (`sheet-menu-insert-table`, the Create Table dialog;
// no shortcut — Ctrl/Cmd+T is the browser's), disabled off a grid sheet.
// SPEC-135 ends the Insert menu with Link… (`sheet-menu-insert-link`,
// Ctrl/Cmd+K — the link dialog), disabled off a grid sheet and, with its
// reason, where links cannot be kept (CSV/TSV, a sheet whose source's links
// stay as read), then after a
// separator Note (`sheet-menu-insert-note`, Shift+F2 — the note editor on the
// active cell) and Comment (`sheet-menu-insert-comment`, Ctrl+Alt+M /
// Cmd+Option+M — the comment composer), each disabled where it cannot be kept
// (Comment, in an .ods workbook, with "OpenDocument spreadsheets support notes
// only"); and ends the View menu, after a separator, with Comments
// (`sheet-menu-comments`, checked while the comments pane is open) and Show
// all notes (`sheet-menu-show-all-notes`, session-only).
// SPEC-133 adds the Formulas menu after Format: Name Manager…
// (`sheet-menu-name-manager`, Ctrl/Cmd+F3), which opens the Name Manager.
// SPEC-136 adds File ▸ Page setup… (`sheet-menu-page-setup`), Print area ▸
// (Set print area / Clear print area, `sheet-menu-set-print-area`,
// `sheet-menu-clear-print-area`) and Print… (`sheet-menu-print`, Ctrl/Cmd+P)
// before Close; View ▸ Page breaks (`sheet-menu-page-breaks`, a
// session-only check item) after Freeze ▸; and, ending the Insert menu after
// a separator, Page break, Remove page break and Reset all page breaks
// (`sheet-menu-insert-page-break`, `-remove-page-break`, `-reset-page-breaks`).
// SPEC-138 adds File ▸ Version history (`sheet-menu-version-history`, checked
// while the side panel's history pane is open) after Print….
// Later specs append menus and items; item test ids keep the `sheet-menu-`
// prefix. Runtime prop declarations per the espree constraint (see
// AppShell.vue).
import { computed } from "vue";
import MenuBar from "../MenuBar.vue";
import MenuCheckItem from "../MenuCheckItem.vue";
import MenuItem from "../MenuItem.vue";
import MenuRadioItem from "../MenuRadioItem.vue";
import MenuSubmenu from "../MenuSubmenu.vue";
import {
  asSpreadsheetMenuState,
  cellStyleCommand,
  FILTER_OFF_REASON,
  numberFormatCommand,
  SHEET_MENU_COMMANDS,
  undoRedoMenuLabel,
  zoomPresetCommand,
} from "../../lib/spreadsheet/menuState";
import { isMacPlatform } from "../../lib/core/platform";
import { shortcutLabel } from "../../lib/shortcuts";
import { NUMBER_FORMAT_PRESETS } from "../../lib/spreadsheet/numberCodes";
import { SHEET_SHORTCUTS } from "../../views/spreadsheet/sheetShortcuts";
import { ZOOM_PRESETS } from "../../lib/zoom";

const props = defineProps({
  /** One object rather than a prop per item — see SpreadsheetMenuState. */
  state: { type: Object, default: () => ({}) },
});

const emit = defineEmits(["command"]);

const s = computed(() => asSpreadsheetMenuState(props.state));

// Accelerators from the SPEC-021 registry, never restated here.
const undoShortcut = shortcutLabel("edit.undo");
const redoShortcut = shortcutLabel("edit.redo");
const formatCellsShortcut = shortcutLabel("sheet.formatCells", isMacPlatform(), SHEET_SHORTCUTS);
const pasteSpecialShortcut = shortcutLabel("sheet.pasteSpecial", isMacPlatform(), SHEET_SHORTCUTS);
const filterShortcut = shortcutLabel("sheet.toggleFilter", isMacPlatform(), SHEET_SHORTCUTS);
const reapplyShortcut = shortcutLabel("sheet.reapplyFilter", isMacPlatform(), SHEET_SHORTCUTS);
const nameManagerShortcut = shortcutLabel("sheet.nameManager", isMacPlatform(), SHEET_SHORTCUTS);
const linkShortcut = shortcutLabel("sheet.insertLink", isMacPlatform(), SHEET_SHORTCUTS);
const noteShortcut = shortcutLabel("sheet.note", isMacPlatform(), SHEET_SHORTCUTS);
const commentShortcut = shortcutLabel("sheet.newComment", isMacPlatform(), SHEET_SHORTCUTS);
const printShortcut = shortcutLabel("sheet.print", isMacPlatform(), SHEET_SHORTCUTS);
// SPEC-128: the Data menu's sorts.
const SORTS = [
  { label: "Sort A to Z", command: SHEET_MENU_COMMANDS.sortAscending, testid: "sheet-menu-sort-asc" },
  { label: "Sort Z to A", command: SHEET_MENU_COMMANDS.sortDescending, testid: "sheet-menu-sort-desc" },
  { label: "Sort range…", command: SHEET_MENU_COMMANDS.sortRange, testid: "sheet-menu-sort-range" },
];
const FILLS = [
  { label: "Down", command: SHEET_MENU_COMMANDS.fillDown, testid: "sheet-menu-fill-down", shortcut: shortcutLabel("sheet.fillDown", isMacPlatform(), SHEET_SHORTCUTS) },
  { label: "Right", command: SHEET_MENU_COMMANDS.fillRight, testid: "sheet-menu-fill-right", shortcut: shortcutLabel("sheet.fillRight", isMacPlatform(), SHEET_SHORTCUTS) },
];
const CLEARS = [
  { label: "Clear all", command: SHEET_MENU_COMMANDS.clearAll, testid: "sheet-menu-clear-all" },
  { label: "Clear formats", command: SHEET_MENU_COMMANDS.clearFormats, testid: "sheet-menu-clear-formats" },
  { label: "Clear contents", command: SHEET_MENU_COMMANDS.clearContents, testid: "sheet-menu-clear-contents" },
];

// SPEC-124: the Format menu's rows and columns.
const SIZES = [
  { label: "Row height…", command: SHEET_MENU_COMMANDS.rowHeight, testid: "sheet-menu-row-height" },
  { label: "Column width…", command: SHEET_MENU_COMMANDS.columnWidth, testid: "sheet-menu-column-width" },
  { label: "Autofit row height", command: SHEET_MENU_COMMANDS.autofitRows, testid: "sheet-menu-autofit-rows" },
  { label: "Autofit column width", command: SHEET_MENU_COMMANDS.autofitColumns, testid: "sheet-menu-autofit-columns" },
];

// SPEC-131: the Arrange menu's z-order moves (`end`: the z-order end at which
// the move cannot move the drawing).
const ARRANGE = [
  { label: "Bring to Front", command: SHEET_MENU_COMMANDS.arrangeFront, testid: "sheet-menu-arrange-front", end: "front" },
  { label: "Bring Forward", command: SHEET_MENU_COMMANDS.arrangeForward, testid: "sheet-menu-arrange-forward", end: "front" },
  { label: "Send Backward", command: SHEET_MENU_COMMANDS.arrangeBackward, testid: "sheet-menu-arrange-backward", end: "back" },
  { label: "Send to Back", command: SHEET_MENU_COMMANDS.arrangeBack, testid: "sheet-menu-arrange-back", end: "back" },
];

const MENUS = [
  { id: "file", label: "File" },
  { id: "edit", label: "Edit" },
  { id: "view", label: "View" },
  { id: "insert", label: "Insert" },
  { id: "arrange", label: "Arrange" },
  { id: "format", label: "Format" },
  { id: "formulas", label: "Formulas" },
  { id: "data", label: "Data" },
];

/** An Arrange move is off without a selected drawing, or at its end. */
function arrangeDisabled(end = "front") {
  return !s.value.drawingSelected || (end === "front" ? s.value.drawingAtFront : s.value.drawingAtBack);
}
</script>

<template>
  <MenuBar
    :menus="MENUS"
    aria-label="Workbook menu"
  >
    <template #file>
      <MenuItem
        label="Download"
        icon="download"
        testid="sheet-menu-download"
        @select="emit('command', SHEET_MENU_COMMANDS.download)"
      />
      <MenuItem
        :label="s.copyLabel"
        testid="sheet-menu-download-copy"
        :disabled="s.copyDisabledReason !== ''"
        :reason="s.copyDisabledReason"
        @select="emit('command', SHEET_MENU_COMMANDS.downloadCopy)"
      />
      <hr class="menu-separator">
      <MenuItem
        label="Page setup…"
        testid="sheet-menu-page-setup"
        :disabled="!s.hasGrid"
        @select="emit('command', SHEET_MENU_COMMANDS.pageSetup)"
      />
      <MenuSubmenu
        label="Print area"
        testid="sheet-menu-print-area"
      >
        <MenuItem
          label="Set print area"
          testid="sheet-menu-set-print-area"
          :disabled="!s.hasGrid"
          @select="emit('command', SHEET_MENU_COMMANDS.setPrintArea)"
        />
        <MenuItem
          label="Clear print area"
          testid="sheet-menu-clear-print-area"
          :disabled="!s.hasGrid || !s.hasPrintArea"
          @select="emit('command', SHEET_MENU_COMMANDS.clearPrintArea)"
        />
      </MenuSubmenu>
      <MenuItem
        label="Print…"
        icon="print"
        testid="sheet-menu-print"
        :shortcut="printShortcut"
        @select="emit('command', SHEET_MENU_COMMANDS.print)"
      />
      <MenuCheckItem
        label="Version history"
        testid="sheet-menu-version-history"
        :checked="s.historyOpen"
        @select="emit('command', SHEET_MENU_COMMANDS.versionHistory)"
      />
      <hr class="menu-separator">
      <MenuItem
        label="Close"
        icon="close"
        testid="sheet-menu-close"
        @select="emit('command', SHEET_MENU_COMMANDS.close)"
      />
    </template>

    <template #edit>
      <MenuItem
        :label="undoRedoMenuLabel('Undo', s.undoLabel)"
        icon="undo"
        testid="sheet-menu-undo"
        :shortcut="undoShortcut"
        :disabled="!s.canUndo"
        @select="emit('command', SHEET_MENU_COMMANDS.undo)"
      />
      <MenuItem
        :label="undoRedoMenuLabel('Redo', s.redoLabel)"
        icon="redo"
        testid="sheet-menu-redo"
        :shortcut="redoShortcut"
        :disabled="!s.canRedo"
        @select="emit('command', SHEET_MENU_COMMANDS.redo)"
      />
      <hr class="menu-separator">
      <MenuItem
        label="Paste special…"
        testid="sheet-menu-paste-special"
        :shortcut="pasteSpecialShortcut"
        :disabled="!s.hasGrid"
        @select="emit('command', SHEET_MENU_COMMANDS.pasteSpecial)"
      />
      <MenuSubmenu
        label="Fill"
        testid="sheet-menu-fill"
      >
        <MenuItem
          v-for="fill in FILLS"
          :key="fill.command"
          :label="fill.label"
          :testid="fill.testid"
          :shortcut="fill.shortcut"
          :disabled="!s.hasGrid"
          @select="emit('command', fill.command)"
        />
      </MenuSubmenu>
    </template>

    <template #view>
      <MenuCheckItem
        label="Gridlines"
        testid="sheet-menu-gridlines"
        :checked="s.gridLines"
        :disabled="!s.hasGrid"
        @select="emit('command', SHEET_MENU_COMMANDS.gridLines)"
      />
      <MenuCheckItem
        label="Show formulas"
        testid="sheet-menu-show-formulas"
        shortcut="Ctrl+`"
        :checked="s.showFormulas"
        :disabled="!s.hasGrid"
        @select="emit('command', SHEET_MENU_COMMANDS.showFormulas)"
      />
      <MenuSubmenu
        label="Zoom"
        testid="sheet-menu-zoom"
      >
        <MenuRadioItem
          v-for="preset in ZOOM_PRESETS"
          :key="preset"
          :label="`${preset}%`"
          :testid="`sheet-menu-zoom-${preset}`"
          :checked="Math.round(s.zoomPercent) === preset"
          :disabled="!s.hasGrid"
          @select="emit('command', zoomPresetCommand(preset))"
        />
        <hr class="menu-separator">
        <MenuItem
          label="Zoom in"
          testid="sheet-menu-zoom-in"
          :disabled="!s.hasGrid"
          @select="emit('command', SHEET_MENU_COMMANDS.zoomIn)"
        />
        <MenuItem
          label="Zoom out"
          testid="sheet-menu-zoom-out"
          :disabled="!s.hasGrid"
          @select="emit('command', SHEET_MENU_COMMANDS.zoomOut)"
        />
      </MenuSubmenu>
      <MenuSubmenu
        label="Freeze"
        testid="sheet-menu-freeze"
      >
        <MenuItem
          label="Top row"
          testid="sheet-menu-freeze-top-row"
          :disabled="!s.hasGrid"
          @select="emit('command', SHEET_MENU_COMMANDS.freezeTopRow)"
        />
        <MenuItem
          label="First column"
          testid="sheet-menu-freeze-first-column"
          :disabled="!s.hasGrid"
          @select="emit('command', SHEET_MENU_COMMANDS.freezeFirstColumn)"
        />
        <MenuItem
          label="Freeze at current cell"
          testid="sheet-menu-freeze-at-cell"
          :disabled="!s.hasGrid || s.freezeAtCellDisabled"
          @select="emit('command', SHEET_MENU_COMMANDS.freezeAtCell)"
        />
        <hr class="menu-separator">
        <MenuItem
          label="Unfreeze"
          testid="sheet-menu-unfreeze"
          :disabled="!s.hasGrid || !s.frozen"
          @select="emit('command', SHEET_MENU_COMMANDS.unfreeze)"
        />
      </MenuSubmenu>
      <MenuCheckItem
        label="Page breaks"
        testid="sheet-menu-page-breaks"
        :checked="s.pageBreaks"
        :disabled="!s.hasGrid"
        @select="emit('command', SHEET_MENU_COMMANDS.pageBreaks)"
      />
      <hr class="menu-separator">
      <MenuCheckItem
        label="Comments"
        testid="sheet-menu-comments"
        :checked="s.commentsOpen"
        @select="emit('command', SHEET_MENU_COMMANDS.comments)"
      />
      <MenuCheckItem
        label="Show all notes"
        testid="sheet-menu-show-all-notes"
        :checked="s.showAllNotes"
        :disabled="!s.canAnnotate"
        @select="emit('command', SHEET_MENU_COMMANDS.showAllNotes)"
      />
    </template>

    <template #insert>
      <MenuItem
        label="Image…"
        icon="image"
        testid="sheet-menu-insert-image"
        :disabled="!s.canInsertDrawings"
        @select="emit('command', SHEET_MENU_COMMANDS.insertImage)"
      />
      <MenuItem
        label="Shapes"
        icon="shape"
        testid="sheet-menu-insert-shape"
        :disabled="!s.canInsertDrawings"
        @select="emit('command', SHEET_MENU_COMMANDS.insertShape)"
      />
      <MenuItem
        label="Chart"
        icon="chart"
        testid="sheet-menu-insert-chart"
        :disabled="!s.canInsertDrawings"
        @select="emit('command', SHEET_MENU_COMMANDS.insertChart)"
      />
      <hr class="menu-separator">
      <MenuItem
        label="Table…"
        testid="sheet-menu-insert-table"
        :disabled="!s.hasGrid"
        @select="emit('command', SHEET_MENU_COMMANDS.insertTable)"
      />
      <MenuItem
        label="Link…"
        icon="link"
        testid="sheet-menu-insert-link"
        :shortcut="linkShortcut"
        :disabled="!s.hasGrid || !s.canLink"
        :reason="s.linkDisabledReason"
        @select="emit('command', SHEET_MENU_COMMANDS.insertLink)"
      />
      <hr class="menu-separator">
      <MenuItem
        label="Note"
        testid="sheet-menu-insert-note"
        :shortcut="noteShortcut"
        :disabled="!s.canAnnotate"
        :reason="s.noteDisabledReason"
        @select="emit('command', SHEET_MENU_COMMANDS.insertNote)"
      />
      <MenuItem
        label="Comment"
        icon="comment"
        testid="sheet-menu-insert-comment"
        :shortcut="commentShortcut"
        :disabled="!s.canComment"
        :reason="s.commentDisabledReason"
        @select="emit('command', SHEET_MENU_COMMANDS.insertComment)"
      />
      <hr class="menu-separator">
      <MenuItem
        label="Page break"
        testid="sheet-menu-insert-page-break"
        :disabled="!s.canInsertPageBreak"
        @select="emit('command', SHEET_MENU_COMMANDS.insertPageBreak)"
      />
      <MenuItem
        label="Remove page break"
        testid="sheet-menu-remove-page-break"
        :disabled="!s.canRemovePageBreak"
        @select="emit('command', SHEET_MENU_COMMANDS.removePageBreak)"
      />
      <MenuItem
        label="Reset all page breaks"
        testid="sheet-menu-reset-page-breaks"
        :disabled="!s.canResetPageBreaks"
        @select="emit('command', SHEET_MENU_COMMANDS.resetPageBreaks)"
      />
    </template>

    <template #arrange>
      <MenuItem
        v-for="move in ARRANGE"
        :key="move.command"
        :label="move.label"
        :testid="move.testid"
        :disabled="arrangeDisabled(move.end)"
        @select="emit('command', move.command)"
      />
      <hr class="menu-separator">
      <MenuItem
        label="Alt Text…"
        testid="sheet-menu-arrange-alt-text"
        :disabled="!s.drawingSelected"
        @select="emit('command', SHEET_MENU_COMMANDS.altText)"
      />
    </template>

    <template #format>
      <MenuItem
        label="Format cells…"
        testid="sheet-menu-format-cells"
        :shortcut="formatCellsShortcut"
        :disabled="!s.hasGrid"
        @select="emit('command', SHEET_MENU_COMMANDS.formatCells)"
      />
      <MenuSubmenu
        label="Number format"
        testid="sheet-menu-number-format"
      >
        <MenuItem
          v-for="preset in NUMBER_FORMAT_PRESETS"
          :key="preset.id"
          :label="preset.label"
          :testid="`sheet-menu-number-format-${preset.id}`"
          :disabled="!s.hasGrid"
          @select="emit('command', numberFormatCommand(preset.id))"
        />
        <hr class="menu-separator">
        <MenuItem
          label="More formats…"
          testid="sheet-menu-number-format-more"
          :disabled="!s.hasGrid"
          @select="emit('command', SHEET_MENU_COMMANDS.moreNumberFormats)"
        />
      </MenuSubmenu>
      <MenuSubmenu
        label="Cell styles"
        testid="sheet-menu-cell-styles"
      >
        <MenuItem
          v-for="name in s.cellStyleNames"
          :key="name"
          :label="name"
          testid="sheet-menu-cell-style"
          :disabled="!s.hasGrid"
          @select="emit('command', cellStyleCommand(name))"
        />
      </MenuSubmenu>
      <MenuItem
        label="Conditional formatting…"
        testid="sheet-menu-conditional-formatting"
        :disabled="!s.hasGrid"
        @select="emit('command', SHEET_MENU_COMMANDS.conditionalFormatting)"
      />
      <MenuSubmenu
        label="Clear"
        testid="sheet-menu-clear"
      >
        <MenuItem
          v-for="clear in CLEARS"
          :key="clear.command"
          :label="clear.label"
          :testid="clear.testid"
          :disabled="!s.hasGrid"
          @select="emit('command', clear.command)"
        />
      </MenuSubmenu>
      <hr class="menu-separator">
      <MenuItem
        v-for="size in SIZES"
        :key="size.command"
        :label="size.label"
        :testid="size.testid"
        :disabled="!s.hasGrid"
        @select="emit('command', size.command)"
      />
    </template>

    <template #formulas>
      <MenuItem
        label="Name Manager…"
        testid="sheet-menu-name-manager"
        :shortcut="nameManagerShortcut"
        @select="emit('command', SHEET_MENU_COMMANDS.nameManager)"
      />
    </template>

    <template #data>
      <MenuItem
        label="Recalculate"
        testid="sheet-menu-recalculate"
        shortcut="F9"
        @select="emit('command', SHEET_MENU_COMMANDS.recalculate)"
      />
      <hr class="menu-separator">
      <MenuItem
        v-for="sort in SORTS"
        :key="sort.command"
        :label="sort.label"
        :testid="sort.testid"
        :disabled="!s.hasGrid"
        @select="emit('command', sort.command)"
      />
      <hr class="menu-separator">
      <MenuCheckItem
        label="Filter"
        testid="sheet-menu-filter"
        :shortcut="filterShortcut"
        :checked="s.filterOn"
        :disabled="!s.hasGrid"
        @select="emit('command', SHEET_MENU_COMMANDS.filter)"
      />
      <MenuItem
        label="Clear filters"
        testid="sheet-menu-clear-filters"
        :disabled="!s.filterOn"
        :reason="s.hasGrid ? FILTER_OFF_REASON : ''"
        @select="emit('command', SHEET_MENU_COMMANDS.clearFilters)"
      />
      <MenuItem
        label="Reapply filter"
        testid="sheet-menu-reapply-filter"
        :shortcut="reapplyShortcut"
        :disabled="!s.filterOn"
        :reason="s.hasGrid ? FILTER_OFF_REASON : ''"
        @select="emit('command', SHEET_MENU_COMMANDS.reapplyFilter)"
      />
      <hr class="menu-separator">
      <MenuItem
        label="Data validation…"
        testid="sheet-menu-data-validation"
        :disabled="!s.hasGrid"
        @select="emit('command', SHEET_MENU_COMMANDS.dataValidation)"
      />
    </template>
  </MenuBar>
</template>
