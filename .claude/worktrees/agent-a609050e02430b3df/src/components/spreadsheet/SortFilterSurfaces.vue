<script setup lang="ts">
// SPEC-128's surfaces in the workbook body, bound to useSortFilter (the
// grid context's `sortFilter`): the Sort dialog, the filter popover with its
// custom-filter and Top 10 dialogs, and the refusal alert (`sheet-alert`).
// Extracted so WorkbookBody only places it. Runtime prop declarations per the
// espree constraint (see AppShell.vue).
import { computed } from "vue";
import { asSortFilter } from "../../views/spreadsheet/useSortFilter";
import FilterCustomDialog from "./FilterCustomDialog.vue";
import FilterPopover from "./FilterPopover.vue";
import FilterTop10Dialog from "./FilterTop10Dialog.vue";
import SheetAlertDialog from "./SheetAlertDialog.vue";
import SortDialog from "./SortDialog.vue";

const props = defineProps({
  /** useSortFilter's return. */
  sortFilter: { type: Object, required: true },
});

const sf = computed(() => asSortFilter(props.sortFilter));
const popover = computed(() => sf.value.popover);

/** The popover's cancel: how it was dismissed. */
function cancelPopover(how = "done") {
  popover.value.close(how === "escape" ? "escape" : how === "outside" ? "outside" : "done");
}
</script>

<template>
  <SortDialog
    v-if="sf.sortDialog.value"
    :with-header="sf.sortDialog.value.withHeader"
    :without-header="sf.sortDialog.value.withoutHeader"
    :has-header="sf.sortDialog.value.hasHeader"
    :header-locked="sf.sortDialog.value.headerLocked"
    :key-col="sf.sortDialog.value.keyCol"
    :error="sf.sortDialog.value.error"
    @confirm="sf.confirmSort"
    @cancel="sf.cancelSort()"
  />
  <FilterPopover
    v-if="popover.view.value"
    :key="`${popover.view.value.col}`"
    :title="popover.view.value.title"
    :sort-asc-label="popover.view.value.sortAscLabel"
    :sort-desc-label="popover.view.value.sortDescLabel"
    :conditions-label="popover.view.value.conditionsLabel"
    :conditions="popover.view.value.conditions"
    :entries="popover.view.value.entries"
    :truncated="popover.view.value.truncated"
    :has-criteria="popover.view.value.hasCriteria"
    :preserved="popover.view.value.preserved"
    :anchor="popover.view.value.anchor"
    @sort="popover.sort"
    @clear="popover.clear()"
    @condition="popover.condition"
    @apply="popover.apply"
    @cancel="cancelPopover"
  />
  <FilterCustomDialog
    v-if="popover.custom.value"
    :title="popover.custom.value.title"
    :kind="popover.custom.value.kind"
    :draft="popover.custom.value.draft"
    @confirm="popover.confirmCustom"
    @cancel="popover.cancelCustom()"
  />
  <FilterTop10Dialog
    v-if="popover.top10.value"
    :title="popover.top10.value.title"
    @confirm="popover.confirmTop10"
    @cancel="popover.cancelTop10()"
  />
  <SheetAlertDialog
    v-if="sf.alert.value"
    :message="sf.alert.value"
    @close="sf.closeAlert()"
  />
</template>
