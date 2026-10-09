<script setup lang="ts">
// The sheet tabs (SPEC-116): one tab per visible sheet, under the grid as in
// every spreadsheet. The ARIA tablist pattern: role="tablist"/"tab",
// aria-selected, a roving tabindex (only the selected tab is in the Tab
// order), and Left/Right/Home/End moving between tabs inside the focused
// tablist — selection follows focus, since showing a sheet is instant.
// `tabs` is [{ index, id, name }] (index = the sheet's position in the
// model). Runtime prop declarations per the espree constraint.
import { computed, nextTick, useTemplateRef } from "vue";
import { asSheetTabs } from "../../views/spreadsheet/useWorkbookLoad";

const props = defineProps({
  tabs: { type: Array, default: () => [] },
  /** The model index of the sheet on show. */
  activeIndex: { type: Number, default: 0 },
  /** The id of the panel the tabs control. */
  panelId: { type: String, default: "" },
});

const emit = defineEmits(["select"]);

const list = useTemplateRef("list");
const sheetTabs = computed(() => asSheetTabs(props.tabs));

async function selectAt(position = 0) {
  const tab = sheetTabs.value[position];
  if (!tab) {
    return;
  }
  emit("select", tab.index);
  await nextTick();
  const buttons = list.value?.querySelectorAll("[role='tab']");
  const button = buttons?.[position];
  if (button instanceof HTMLElement) {
    button.focus();
  }
}

function onKeydown(event = new KeyboardEvent("keydown")) {
  const count = sheetTabs.value.length;
  if (count === 0) {
    return;
  }
  const current = Math.max(0, sheetTabs.value.findIndex((tab) => tab.index === props.activeIndex));
  const targets = {
    ArrowRight: (current + 1) % count,
    ArrowLeft: (current - 1 + count) % count,
    Home: 0,
    End: count - 1,
  };
  const target = Object.entries(targets).find(([key]) => key === event.key)?.[1];
  if (target === undefined) {
    return;
  }
  event.preventDefault();
  void selectAt(target);
}
</script>

<template>
  <div
    ref="list"
    class="sheet-tabs"
    role="tablist"
    aria-label="Sheets"
    aria-orientation="horizontal"
    data-testid="sheet-tabs"
    @keydown="onKeydown"
  >
    <button
      v-for="tab in sheetTabs"
      :id="`sheet-tab-${tab.id}`"
      :key="tab.id"
      type="button"
      class="sheet-tab"
      :class="{ 'is-selected': tab.index === props.activeIndex }"
      role="tab"
      data-testid="sheet-tab"
      :aria-selected="tab.index === props.activeIndex"
      :aria-controls="props.panelId || undefined"
      :tabindex="tab.index === props.activeIndex ? 0 : -1"
      @click="emit('select', tab.index)"
    >
      {{ tab.name }}
    </button>
  </div>
</template>

<style scoped>
.sheet-tabs {
  display: flex;
  flex: none;
  gap: var(--space-1);
  padding: var(--space-1) var(--space-3);
  overflow-x: auto;
  background: var(--surface-sunken);
  border-top: var(--border-width) solid var(--border-default);
}

.sheet-tab {
  flex: none;
  max-width: 16rem;
  padding: var(--space-1) var(--space-4);
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
  border: var(--border-width) solid transparent;
  border-radius: var(--radius-sm);
  background: transparent;
  color: var(--text-secondary);
  font: inherit;
  font-size: var(--text-sm);
  cursor: pointer;
}

.sheet-tab:hover {
  background: var(--state-hover);
}

.sheet-tab.is-selected {
  background: var(--surface-raised);
  border-color: var(--border-default);
  color: var(--text-primary);
  font-weight: var(--weight-medium);
}

.sheet-tab:focus-visible {
  outline: var(--space-1) solid var(--focus-ring);
  outline-offset: 1px;
}
</style>
