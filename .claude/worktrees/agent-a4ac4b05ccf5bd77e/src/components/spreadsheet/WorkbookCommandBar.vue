<script setup lang="ts">
// The workbook's command rows in AppShell's menu-bar region (outside the
// scrolling canvas, so their popups are never clipped): SPEC-117's menu bar
// over SPEC-123's sheet formatting strip. Extracted so SpreadsheetView keeps
// its line budget (SpreadsheetView.size.spec.ts). Runtime prop declarations
// per the espree constraint (see AppShell.vue).
import { computed } from "vue";
import SheetFormatToolbar from "./SheetFormatToolbar.vue";
import SpreadsheetMenuBar from "./SpreadsheetMenuBar.vue";

const props = defineProps({
  /** useSpreadsheetMenu's return. */
  menu: { type: Object, required: true },
  /** useFormatting's return. */
  formatting: { type: Object, required: true },
  /** useFormatPainter's return. */
  painter: { type: Object, required: true },
  /** SPEC-124: useMerges' return (the strip's merge split button). */
  merges: { type: Object, default: null },
});

const state = computed(() => Object(props.menu).state.value);

function run(command = "") {
  Object(props.menu).run(command);
}
</script>

<template>
  <SpreadsheetMenuBar
    :state="state"
    @command="run"
  />
  <SheetFormatToolbar
    :formatting="props.formatting"
    :painter="props.painter"
    :merges="props.merges"
  />
</template>
