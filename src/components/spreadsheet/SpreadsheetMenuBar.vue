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
// (after Cell styles ▸), which opens the side panel's rules manager.
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

const MENUS = [
  { id: "file", label: "File" },
  { id: "edit", label: "Edit" },
  { id: "view", label: "View" },
  { id: "format", label: "Format" },
  { id: "data", label: "Data" },
];
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
    </template>
  </MenuBar>
</template>
