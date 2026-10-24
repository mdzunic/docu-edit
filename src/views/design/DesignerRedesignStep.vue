<script setup lang="ts">
// SPEC-115: Redesign's plan step. One row per existing slide — its original
// number and its read-only title ("Slide N" when it has none) — with a layout
// select of the six kinds plus "Keep current", Alt+↑/Alt+↓ (or the arrow
// buttons) to reorder, and the deck's theme select ("Keep current theme" plus
// every built-in). Deliberately no add-slide or delete-slide control and no
// editable text: Redesign changes only order, layouts and theme.
// Runtime prop declarations per the espree constraint (see AppShell.vue).
import { computed, nextTick, useTemplateRef } from "vue";
import UiButton from "../../components/ui/UiButton.vue";
import UiSelect from "../../components/ui/UiSelect.vue";
import DesignerDrops from "./DesignerDrops.vue";
import { asDeckDesigner, reorderDelta } from "./designerFields";

const props = defineProps({
  /** The useDeckDesigner() session. */
  controller: { type: Object, required: true },
});

const c = asDeckDesigner(props.controller);
const list = useTemplateRef("list");

const layoutOptions = computed(() => [{ value: "", label: "Keep current" }, ...c.layoutOptions]);
const themeOptions = computed(() => [{ value: "", label: "Keep current theme" }, ...c.themeNames]);

async function move(index = 0, delta = 0, testid = "") {
  const target = c.moveRow(index, delta);
  if (target === index) {
    return;
  }
  await nextTick();
  const row = list.value?.querySelectorAll('[data-testid="designer-redesign-row"]')[target];
  const control = testid ? row?.querySelector(`[data-testid="${testid}"]`) : null;
  if (control instanceof HTMLElement) {
    control.focus();
  }
}

function onRowKeydown(index = 0, event = new KeyboardEvent("keydown")) {
  const delta = reorderDelta(event);
  if (delta === 0) {
    return;
  }
  event.preventDefault();
  const testid = event.target instanceof HTMLElement ? (event.target.dataset.testid ?? "") : "";
  void move(index, delta, testid);
}
</script>

<template>
  <div class="redesign">
    <label class="redesign-theme">
      <span>Theme</span>
      <UiSelect
        data-testid="designer-redesign-theme"
        :model-value="c.redesignTheme.value"
        :options="themeOptions"
        @update:model-value="c.setRedesignTheme($event)"
      />
    </label>

    <DesignerDrops :records="c.dropped.value" />

    <p class="redesign-hint">
      Slides keep their text and notes; only their order, layouts and the theme change. Alt+↑ and Alt+↓ move a slide.
    </p>
    <ol
      ref="list"
      class="redesign-list"
      data-testid="designer-redesign"
      aria-label="The deck's slides, in their new order"
    >
      <li
        v-for="(row, index) in c.rows.value"
        :key="row.key"
        class="redesign-row"
        data-testid="designer-redesign-row"
        :data-source="row.source"
        @keydown="onRowKeydown(index, $event)"
      >
        <span
          class="redesign-number"
          data-testid="designer-row-number"
        >Slide {{ row.source + 1 }}</span>
        <span
          class="redesign-title"
          data-testid="designer-row-title"
        >{{ row.title || `Slide ${row.source + 1}` }}</span>
        <UiSelect
          data-testid="designer-row-layout"
          :label="`Layout for slide ${row.source + 1}`"
          :model-value="row.layout ?? ''"
          :options="layoutOptions"
          @update:model-value="c.setRowLayout(index, $event)"
        />
        <UiButton
          size="sm"
          data-testid="designer-row-up"
          :aria-label="`Move slide ${row.source + 1} up`"
          :disabled="index === 0"
          @click="move(index, -1, 'designer-row-up')"
        >
          ↑
        </UiButton>
        <UiButton
          size="sm"
          data-testid="designer-row-down"
          :aria-label="`Move slide ${row.source + 1} down`"
          :disabled="index === c.rows.value.length - 1"
          @click="move(index, 1, 'designer-row-down')"
        >
          ↓
        </UiButton>
      </li>
    </ol>
  </div>
</template>

<style scoped>
.redesign {
  display: flex;
  flex-direction: column;
  gap: var(--space-3);
}

.redesign-theme {
  display: flex;
  align-items: center;
  gap: var(--space-3);
  color: var(--text-secondary);
  font-size: var(--text-sm);
}

.redesign-hint {
  margin: 0;
  color: var(--text-secondary);
  font-size: var(--text-sm);
}

.redesign-list {
  display: flex;
  flex-direction: column;
  gap: var(--space-2);
  margin: 0;
  padding: 0;
  list-style: none;
}

.redesign-row {
  display: grid;
  grid-template-columns: 5rem minmax(0, 1fr) auto auto auto;
  align-items: center;
  gap: var(--space-3);
  padding: var(--space-2) var(--space-3);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-sm);
}

.redesign-row:focus-within {
  border-color: var(--border-strong);
}

.redesign-number {
  color: var(--text-secondary);
  font-family: var(--font-caption);
  font-size: var(--text-sm);
}

.redesign-title {
  overflow: hidden;
  color: var(--text-primary);
  text-overflow: ellipsis;
  white-space: nowrap;
}
</style>
