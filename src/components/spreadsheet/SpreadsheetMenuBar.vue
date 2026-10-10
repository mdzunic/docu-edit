<script setup lang="ts">
// The workbook's menu bar (SPEC-117 "The menu bar") — PresentationMenuBar's
// pattern for the spreadsheet product: the product-neutral MenuBar with
// slotted menus, every item emitting a command id that useSpreadsheetMenu
// routes, its state coerced by lib/spreadsheet/menuState.ts. SPEC-117's
// items: File — Download (the same action as the app bar's) and Close;
// View — Gridlines (a menuitemcheckbox), SPEC-119's Show formulas (another,
// with its Ctrl+` shortcut — Control on every platform) and Zoom ▸ the
// ZOOM_PRESETS ladder, Zoom in and Zoom out. SPEC-120's Data menu ends, after
// a separator, with Recalculate (F9) — SPEC-128 puts Sort and Filter above
// that separator; until then nothing is, so a leading separator is not drawn.
// SPEC-121 adds the Edit menu between File and View — Undo and Redo, each
// naming its step ("Undo Rename sheet") with the registry's shortcut, and
// disabled when there is nothing to undo or redo — and File ▸ "Download a
// copy as …" after Download (disabled, with the reason, until the copy can be
// made). SPEC-123 adds the Format menu after View: Format cells…, Number
// format ▸ (the presets and More formats…), Cell styles ▸ and Clear ▸.
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
const CLEARS = [
  { label: "Clear all", command: SHEET_MENU_COMMANDS.clearAll, testid: "sheet-menu-clear-all" },
  { label: "Clear formats", command: SHEET_MENU_COMMANDS.clearFormats, testid: "sheet-menu-clear-formats" },
  { label: "Clear contents", command: SHEET_MENU_COMMANDS.clearContents, testid: "sheet-menu-clear-contents" },
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
    </template>

    <template #data>
      <hr class="menu-separator data-recalculate-separator">
      <MenuItem
        label="Recalculate"
        testid="sheet-menu-recalculate"
        shortcut="F9"
        @select="emit('command', SHEET_MENU_COMMANDS.recalculate)"
      />
    </template>
  </MenuBar>
</template>

<style scoped>
/* The separator before Recalculate divides it from the Data commands above
   it (SPEC-128's); with nothing above it yet — first in the popup, or right
   under the compact sheet's heading — it is not drawn. */
.data-recalculate-separator:first-child,
.menu-section-label + .data-recalculate-separator {
  display: none;
}
</style>
