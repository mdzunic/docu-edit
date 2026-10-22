<script setup lang="ts">
// SPEC-127's function autocomplete: the registered names starting with the
// identifier being typed (`function-autocomplete`, a listbox; each name a
// `function-autocomplete-item` option with `data-name`), the highlighted
// one marked and its summary below (`function-autocomplete-summary`). At
// most ten rows show; the rest scroll. A press chooses a name without taking
// focus from the editor. State and keys are useFunctionAssist's. Runtime
// prop declarations per the espree constraint (see AppShell.vue).
import { nextTick, useTemplateRef, watch } from "vue";
import { AUTOCOMPLETE_ID } from "../../views/spreadsheet/useFunctionAssist";

const props = defineProps({
  /** The matching names, alphabetically. */
  names: { type: Array, default: () => [] },
  /** The highlighted name. */
  highlighted: { type: String, default: "" },
  /** Its summary. */
  summary: { type: String, default: "" },
  /** The fixed position below the editor. */
  anchorStyle: { type: Object, default: () => ({}) },
});

const emit = defineEmits(["choose"]);
const list = useTemplateRef("list");

// Keep the highlighted row in view as Up/Down move it.
watch(
  () => props.highlighted,
  async () => {
    await nextTick();
    const row = list.value?.querySelector('[aria-selected="true"]');
    if (row instanceof HTMLElement && typeof row.scrollIntoView === "function") {
      row.scrollIntoView({ block: "nearest" });
    }
  },
);
</script>

<template>
  <div
    class="function-autocomplete-popup"
    :style="props.anchorStyle"
  >
    <ul
      :id="AUTOCOMPLETE_ID"
      ref="list"
      class="function-autocomplete"
      data-testid="function-autocomplete"
      role="listbox"
      aria-label="Functions"
    >
      <li
        v-for="name in props.names"
        :id="`${AUTOCOMPLETE_ID}-${name}`"
        :key="String(name)"
        class="function-autocomplete-item"
        :class="{ 'is-highlighted': name === props.highlighted }"
        data-testid="function-autocomplete-item"
        role="option"
        :data-name="name"
        :aria-selected="name === props.highlighted ? 'true' : 'false'"
        @mousedown.prevent="emit('choose', name)"
      >
        {{ name }}
      </li>
    </ul>
    <p
      class="function-autocomplete-summary"
      data-testid="function-autocomplete-summary"
    >
      {{ props.summary }}
    </p>
  </div>
</template>

<style scoped>
.function-autocomplete-popup {
  position: fixed;
  z-index: 40;
  width: 18rem;
  background: var(--surface-raised);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-sm);
  box-shadow: var(--shadow-2);
  font-family: var(--font-ui);
  font-size: var(--text-sm);
  color: var(--text-primary);
}

/* Ten rows, then a scroll. */
.function-autocomplete {
  max-height: calc(10 * 1.75rem);
  margin: 0;
  padding: var(--space-1) 0;
  overflow-y: auto;
  list-style: none;
}

.function-autocomplete-item {
  height: 1.75rem;
  padding: 0 var(--space-2);
  line-height: 1.75rem;
  cursor: pointer;
}

.function-autocomplete-item.is-highlighted {
  background: var(--accent-subtle, var(--surface-sunken));
  font-weight: var(--weight-button);
}

.function-autocomplete-summary {
  margin: 0;
  padding: var(--space-2);
  border-top: var(--border-width) solid var(--border-default);
  color: var(--text-secondary);
}
</style>
