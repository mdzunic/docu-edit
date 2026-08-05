<script setup lang="ts">
// The "Keyboard shortcuts" reference dialog (SPEC-021 AC-6..10): every entry
// is generated from the shortcuts registry, so it can never drift from what
// the app actually binds. Runtime prop/emit declarations, not the
// defineProps<{...}>()/defineEmits<{...}>() generics — see AppShell.vue's
// comment on the same espree constraint.
import { computed, nextTick, onMounted, ref, useTemplateRef } from "vue";
import { formatBinding, groupShortcuts, RESERVED_BROWSER_KEYS, searchShortcuts } from "../lib/shortcuts";

const emit = defineEmits(["close"]);

const query = ref("");
const activeIndex = ref(0);
const searchInput = useTemplateRef("searchInput");
const listEl = useTemplateRef("listEl");

onMounted(async () => {
  await nextTick();
  searchInput.value?.focus();
});

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
  <div
    class="dialog-backdrop"
    data-testid="shortcuts-dialog"
    @click.self="emit('close')"
    @keydown.escape="emit('close')"
  >
    <div
      class="dialog"
      role="dialog"
      aria-modal="true"
      aria-label="Keyboard shortcuts"
    >
      <h2 class="dialog-title">
        Keyboard shortcuts
      </h2>
      <input
        ref="searchInput"
        v-model="query"
        type="search"
        class="search-input"
        data-testid="shortcuts-search"
        aria-label="Search shortcuts"
        placeholder="Search shortcuts"
        aria-controls="shortcuts-list"
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
      <div class="dialog-actions">
        <button
          type="button"
          data-testid="shortcuts-close"
          @click="emit('close')"
        >
          Close
        </button>
      </div>
    </div>
  </div>
</template>

<style scoped>
.dialog-backdrop {
  position: fixed;
  inset: 0;
  background: rgb(0 0 0 / 0.4);
  display: flex;
  align-items: center;
  justify-content: center;
  z-index: 50;
}

.dialog {
  display: flex;
  flex-direction: column;
  background: #fff;
  border-radius: 6px;
  box-shadow: 0 4px 16px rgb(0 0 0 / 0.3);
  padding: 1.25rem 1.5rem;
  width: 32rem;
  max-width: 90vw;
  max-height: 80vh;
  box-sizing: border-box;
}

.dialog-title {
  margin: 0 0 1rem;
  font-size: 1rem;
  font-weight: 600;
  color: #202124;
}

.search-input {
  font: inherit;
  width: 100%;
  box-sizing: border-box;
  padding: 0.375rem 0.5rem;
  border: 1px solid #bdbdbd;
  border-radius: 4px;
  margin-bottom: 0.75rem;
}

.shortcuts-list {
  overflow-y: auto;
  flex: 1 1 auto;
  min-height: 0;
  border: 1px solid #e0e0e0;
  border-radius: 4px;
  padding: 0.25rem 0;
}

.empty-state {
  margin: 0;
  padding: 0.75rem;
  color: #5f6368;
  font-size: 0.875rem;
}

.group-title {
  margin: 0.5rem 0 0.25rem;
  padding: 0 0.75rem;
  font-size: 0.75rem;
  font-weight: 600;
  text-transform: uppercase;
  letter-spacing: 0.03em;
  color: #5f6368;
}

.shortcut-row {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 1rem;
  padding: 0.375rem 0.75rem;
  font-size: 0.875rem;
  color: #202124;
}

.shortcut-row.active {
  background: #e8f0fe;
}

.shortcut-note {
  color: #5f6368;
  font-size: 0.75rem;
}

.shortcut-keys {
  display: flex;
  gap: 0.25rem;
  flex: none;
}

.key-chip {
  font-size: 0.75rem;
  padding: 0.0625rem 0.375rem;
  border: 1px solid #dadce0;
  border-radius: 4px;
  background: #f5f5f5;
  color: #3c4043;
  white-space: nowrap;
}

.reserved-note {
  margin: 0.75rem 0 0;
  font-size: 0.75rem;
  color: #5f6368;
}

.dialog-actions {
  display: flex;
  justify-content: flex-end;
  gap: 0.5rem;
  margin-top: 1rem;
}

.dialog-actions button {
  font: inherit;
  padding: 0.375rem 1rem;
  border: 1px solid #bdbdbd;
  border-radius: 4px;
  background: #fff;
  cursor: pointer;
}
</style>
