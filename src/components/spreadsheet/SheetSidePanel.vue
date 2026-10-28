<script setup lang="ts">
// The spreadsheet view's one side panel (SPEC-129 "The side panel"): the
// shared SidePanel.vue chrome (SPEC-065) driven by useWorkbookEditing's
// `sidePanel` (useSheetSidePanel), with the active pane's component in its
// default slot — this spec's `conditionalFormats` pane, the rules manager
// over useWorkbookEditing's `cfManager`. Later spreadsheet panes register
// here (SPEC-135, SPEC-138, SPEC-139) rather than mounting panels of their
// own. SPEC-135 adds the `comments` pane: SheetCommentsPanel over
// useWorkbookEditing's `comments`; SPEC-138 the `history` pane:
// WorkbookHistoryPanel over useWorkbookHistory. Runtime prop declarations per
// the espree constraint (see AppShell.vue).
import { computed } from "vue";
import { clampSheetPanelWidth } from "../../views/spreadsheet/useSheetSidePanel";
import { asWorkbookEditing } from "../../views/spreadsheet/useWorkbookEditing";
import SidePanel from "../SidePanel.vue";
import ConditionalFormatsPane from "./ConditionalFormatsPane.vue";
import SheetCommentsPanel from "./SheetCommentsPanel.vue";
import WorkbookHistoryPanel from "./WorkbookHistoryPanel.vue";

const props = defineProps({
  /** useWorkbookEditing's return. */
  editing: { type: Object, required: true },
  /** SPEC-138: useWorkbookHistory's return (the `history` pane). */
  history: { type: Object, default: null },
});

const editing = computed(() => asWorkbookEditing(props.editing));
const panel = computed(() => editing.value.sidePanel);

function activate(id = "") {
  if (id === "conditionalFormats" || id === "comments" || id === "history") {
    panel.value.open(id);
  }
}

function toggleSide() {
  panel.value.side.value = panel.value.side.value === "right" ? "left" : "right";
}

function toggleCollapse() {
  panel.value.collapsed.value = !panel.value.collapsed.value;
}

function setWidth(width = 0) {
  panel.value.widthPx.value = clampSheetPanelWidth(width);
}

function nudgeWidth(delta = 0) {
  panel.value.widthPx.value = clampSheetPanelWidth(panel.value.widthPx.value + delta);
}
</script>

<template>
  <SidePanel
    v-if="panel.activePane.value"
    :side="panel.side.value"
    :width-px="panel.widthPx.value"
    :collapsed="panel.collapsed.value"
    :tabs="panel.tabs.value"
    :active-pane="panel.activePane.value"
    @activate="activate"
    @close="panel.close()"
    @toggle-side="toggleSide"
    @toggle-collapse="toggleCollapse"
    @update:width-px="setWidth"
    @commit-width="setWidth(panel.widthPx.value)"
    @nudge-width="nudgeWidth"
  >
    <ConditionalFormatsPane
      v-if="panel.activePane.value === 'conditionalFormats'"
      :manager="editing.cfManager"
    />
    <SheetCommentsPanel
      v-else-if="panel.activePane.value === 'comments'"
      :controller="editing.comments"
    />
    <WorkbookHistoryPanel
      v-else-if="panel.activePane.value === 'history' && props.history"
      :controller="props.history"
    />
  </SidePanel>
</template>
