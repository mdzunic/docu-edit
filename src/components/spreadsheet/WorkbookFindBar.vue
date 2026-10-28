<script setup lang="ts">
// The workbook view's find dock wiring (SPEC-134): SPEC-017's FindBar, bound
// to useSheetFind's `bar` and `barEvents`, with the sheet's find options in
// its options slot, inside SheetFindResults (Find all's list below it).
// Extracted by SPEC-138 so SpreadsheetView keeps its line budget
// (SpreadsheetView.size.spec.ts) as the saving and history wiring arrived —
// the move WorkbookAppBar and WorkbookStatusBar made. Runtime prop
// declarations per the espree constraint (see AppShell.vue).
import { computed } from "vue";
import { asSheetFind } from "../../views/spreadsheet/useSheetFind";
import FindBar from "../FindBar.vue";
import SheetFindOptions from "./SheetFindOptions.vue";
import SheetFindResults from "./SheetFindResults.vue";

const props = defineProps({
  /** useSheetFind's return. */
  find: { type: Object, required: true },
});

const find = computed(() => asSheetFind(props.find));
</script>

<template>
  <SheetFindResults :find="props.find">
    <FindBar
      v-bind="find.bar.value"
      v-on="find.barEvents"
    >
      <template #options>
        <SheetFindOptions :find="props.find" />
      </template>
    </FindBar>
  </SheetFindResults>
</template>
