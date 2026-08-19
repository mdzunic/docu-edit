<script setup lang="ts">
// The "Keyboard shortcuts" reference dialog (SPEC-021 AC-6..10): every entry
// is generated from the shortcuts registry, so it can never drift from what
// the app actually binds. BaseDialog owns the shell, Escape and initial
// focus (the [autofocus] search input). Runtime prop/emit declarations, not
// the defineProps<{...}>()/defineEmits<{...}>() generics — see AppShell.vue's
// comment on the same espree constraint.
import { computed, nextTick, ref, useTemplateRef } from "vue";
import { formatBinding, groupShortcuts, RESERVED_BROWSER_KEYS, searchShortcuts } from "../lib/shortcuts";
import BaseDialog from "./BaseDialog.vue";

const emit = defineEmits(["close"]);

const query = ref("");
const activeIndex = ref(0);
const listEl = useTemplateRef("listEl");

// Filtered once, then grouped — grouping never changes membership, only
// how the same filtered set is arranged, so search and category grouping
// can't disagree about which entries are visible.
const filtered = computed(() => searchShortcuts(query.value));
const groups = computed(() => groupShortcuts(filtered.value));

// Arrow-key navigation moves a highlight across category boundaries, so it
// works over a flat index rather than a per-group one.
const flatIds = computed(() => filtered.value.map((definition) => definition.id));

function rowId(id = "") {
  return `shortcut-row-${id}`;
}

function onQueryInput() {
  activeIndex.value = 0;
}

function moveActive(delta = 0) {
  if (flatIds.value.length === 0) {
    return;
  }
  activeIndex.value = (activeIndex.value + delta + flatIds.value.length) % flatIds.value.length;
  void nextTick(() => {
    const id = flatIds.value[activeIndex.value];
    const row = id ? listEl.value?.querySelector(`[data-testid="shortcut-row-${id}"]`) : null;
    row?.scrollIntoView?.({ block: "nearest" });
  });
}

function onNavigationKeydown(event = new KeyboardEvent("keydown")) {
  if (event.key === "ArrowDown") {
    event.preventDefault();
    moveActive(1);
  } else if (event.key === "ArrowUp") {
    event.preventDefault();
    moveActive(-1);
  } else if (event.key === "Home") {
    event.preventDefault();
    activeIndex.value = 0;
  } else if (event.key === "End") {
    event.preventDefault();
    activeIndex.value = Math.max(0, flatIds.value.length - 1);
  }
}

function isActive(id = "") {
  return flatIds.value.length > 0 && flatIds.value[activeIndex.value] === id;
}

const reservedKeysLabel = RESERVED_BROWSER_KEYS.map((key) => `Ctrl/Cmd+${key.toUpperCase()}`).join(", ");
</script>

<template>
  <BaseDialog
    title="Keyboard shortcuts"
    testid="shortcuts-dialog"
    @dismiss="emit('close')"
  >
    <input
      v-model="query"
      type="search"
      class="search-input"
      data-testid="shortcuts-search"
      aria-label="Search shortcuts"
      placeholder="Search shortcuts"
      aria-controls="shortcuts-list"
      autofocus
      :aria-activedescendant="flatIds[activeIndex] ? rowId(flatIds[activeIndex]) : undefined"
      @input="onQueryInput"
      @keydown="onNavigationKeydown"
    >
    <div
      id="shortcuts-list"
      ref="listEl"
      class="shortcuts-list"
      role="listbox"
      aria-label="Registered shortcuts"
      data-testid="shortcuts-list"
    >
      <p
        v-if="groups.length === 0"
        class="empty-state"
        data-testid="shortcuts-empty"
      >
        No shortcuts match "{{ query }}"
      </p>
      <div
        v-for="group in groups"
        :key="group.category"
        class="shortcut-group"
      >
        <h3 class="group-title">
          {{ group.category }}
        </h3>
        <div
          v-for="definition in group.items"
          :id="rowId(definition.id)"
          :key="definition.id"
          role="option"
          class="shortcut-row"
          :class="{ active: isActive(definition.id) }"
          :aria-selected="isActive(definition.id) ? 'true' : 'false'"
          :data-testid="`shortcut-row-${definition.id}`"
        >
          <span class="shortcut-label">
            {{ definition.label }}
            <span
              v-if="definition.overridesBrowser"
              class="shortcut-note"
            >(overrides browser default)</span>
          </span>
          <span class="shortcut-keys">
            <span
              v-for="binding in definition.bindings"
              :key="`${binding.key}-${binding.mod}-${binding.shift}`"
              class="key-chip"
            >{{ formatBinding(binding) }}</span>
          </span>
        </div>
      </div>
    </div>
    <p
      class="reserved-note"
      data-testid="shortcuts-reserved-note"
    >
      {{ reservedKeysLabel }} always stay the browser's — this app never binds them.
    </p>
    <template #actions>
      <button
        type="button"
        data-testid="shortcuts-close"
        @click="emit('close')"
      >
        Close
      </button>
    </template>
  </BaseDialog>
</template>

<style scoped>
/* Fixed 32rem panel laid out as a column so .shortcuts-list flex-scrolls
   between the search input and the reserved-keys note, capped at 80vh. */
:deep(.dialog) {
  display: flex;
  flex-direction: column;
  width: 32rem;
  max-height: 80vh;
}

.search-input {
  font: inherit;
  width: 100%;
  box-sizing: border-box;
  padding: var(--space-3) var(--space-3);
  border: var(--border-width) solid var(--border-strong);
  border-radius: var(--radius-sm);
  margin-bottom: var(--space-4);
}

.shortcuts-list {
  overflow-y: auto;
  flex: 1 1 auto;
  min-height: 0;
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-sm);
  padding: var(--space-2) 0;
}

.empty-state {
  margin: 0;
  padding: var(--space-4);
  color: var(--text-secondary);
  font-size: var(--text-md);
}

.group-title {
  margin: var(--space-3) 0 var(--space-2);
  padding: 0 var(--space-4);
  font-size: var(--text-sm);
  font-weight: var(--weight-bold);
  text-transform: uppercase;
  letter-spacing: 0.03em;
  color: var(--text-secondary);
}

.shortcut-row {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: var(--space-5);
  padding: var(--space-3) var(--space-4);
  font-size: var(--text-md);
  color: var(--text-primary);
}

.shortcut-row.active {
  background: var(--surface-selected);
}

.shortcut-note {
  color: var(--text-secondary);
  font-size: var(--text-sm);
}

.shortcut-keys {
  display: flex;
  gap: var(--space-2);
  flex: none;
}

.key-chip {
  font-size: var(--text-sm);
  padding: var(--space-1) var(--space-3);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-sm);
  background: var(--surface-sunken);
  color: var(--text-primary);
  white-space: nowrap;
}

.reserved-note {
  margin: var(--space-4) 0 0;
  font-size: var(--text-sm);
  color: var(--text-secondary);
}

.dialog-actions button {
  font: inherit;
  padding: var(--space-3) var(--space-5);
  border: var(--border-width) solid var(--border-strong);
  border-radius: var(--radius-sm);
  background: var(--surface-base);
  cursor: pointer;
}
</style>
