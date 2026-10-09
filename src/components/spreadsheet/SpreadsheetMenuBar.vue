<script setup lang="ts">
// The workbook's menu bar (SPEC-117 "The menu bar") — PresentationMenuBar's
// pattern for the spreadsheet product: the product-neutral MenuBar with
// slotted menus, every item emitting a command id that useSpreadsheetMenu
// routes, its state coerced by lib/spreadsheet/menuState.ts. This spec's
// items: File — Download (the same action as the app bar's) and Close;
// View — Gridlines (a menuitemcheckbox) and Zoom ▸ the ZOOM_PRESETS ladder,
// Zoom in and Zoom out. Later specs append menus and items; item test ids
// keep the `sheet-menu-` prefix. Runtime prop declarations per the espree
// constraint (see AppShell.vue).
import { computed } from "vue";
import MenuBar from "../MenuBar.vue";
import MenuCheckItem from "../MenuCheckItem.vue";
import MenuItem from "../MenuItem.vue";
import MenuRadioItem from "../MenuRadioItem.vue";
import MenuSubmenu from "../MenuSubmenu.vue";
import { asSpreadsheetMenuState, SHEET_MENU_COMMANDS, zoomPresetCommand } from "../../lib/spreadsheet/menuState";
import { ZOOM_PRESETS } from "../../lib/zoom";

const props = defineProps({
  /** One object rather than a prop per item — see SpreadsheetMenuState. */
  state: { type: Object, default: () => ({}) },
});

const emit = defineEmits(["command"]);

const s = computed(() => asSpreadsheetMenuState(props.state));

const MENUS = [
  { id: "file", label: "File" },
  { id: "view", label: "View" },
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
      <hr class="menu-separator">
      <MenuItem
        label="Close"
        icon="close"
        testid="sheet-menu-close"
        @select="emit('command', SHEET_MENU_COMMANDS.close)"
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
  </MenuBar>
</template>
