<script setup lang="ts">
import TableStyleGalleryDialog from "./TableStyleGalleryDialog.vue";
import TableSortDialog from "./TableSortDialog.vue";
import TableSeparatorDialog from "./TableSeparatorDialog.vue";
import { asConvertSeparatorId } from "../lib/tableConvert";
import { asTableLookKey } from "../lib/tableStyles";

// The Table menu's three dialogs, hosted together (SPEC-076 AC-79). The
// whole API arrives as one `tools` object — useTableTools' return value —
// so DocumentView's template wiring stays a single binding, exactly as
// PointEditOverlay's `controller` does for its own size budget. Nothing is
// decided here: this only routes `tools.dialog` to the dialog it names and
// hands each dialog's result straight back to the composable.
// Runtime prop declaration per the .vue espree constraint.
defineProps({ tools: { type: Object, required: true } });
</script>

<template>
  <TableStyleGalleryDialog
    v-if="tools.dialog.value === 'gallery' && tools.caretTable.value"
    :table="tools.caretTable.value"
    :styles="tools.docStyles.value"
    @apply="(id = null) => tools.applyGalleryStyle(id)"
    @toggle-look="(key = '') => tools.toggleLookFlag(asTableLookKey(key))"
    @cancel="tools.closeDialog()"
  />
  <TableSortDialog
    v-else-if="tools.dialog.value === 'sort' && tools.caretTable.value"
    :table="tools.caretTable.value"
    :initial-keys="tools.sortKeys.value"
    :initial-has-header="tools.sortHasHeader.value"
    :blocked-reason="tools.sortBlockedReason(tools.sortHasHeader.value)"
    @header-change="(value = false) => tools.setSortHasHeader(value)"
    @apply="(keys = [], hasHeader = false) => tools.applySort(keys, hasHeader)"
    @cancel="tools.closeDialog()"
  />
  <TableSeparatorDialog
    v-else-if="tools.dialog.value === 'text-to-table'"
    direction="text-to-table"
    :paragraphs="tools.convertSourceParagraphs.value"
    @apply="(id = 'tab', custom = '') => tools.applyTextToTable(asConvertSeparatorId(id), custom)"
    @cancel="tools.closeDialog()"
  />
  <TableSeparatorDialog
    v-else-if="tools.dialog.value === 'table-to-text'"
    direction="table-to-text"
    @apply="(id = 'tab', custom = '') => tools.applyTableToText(asConvertSeparatorId(id), custom)"
    @cancel="tools.closeDialog()"
  />
</template>
