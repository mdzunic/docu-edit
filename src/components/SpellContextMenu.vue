<script setup lang="ts">
// The spell-check context menu (SPEC-035 AC-2, AC-3, AC-3.1): up to five
// suggestions for the misspelled word right-clicked, plus Ignore/Ignore
// all/Add to dictionary — always available, even with no suggestions.
// Same positioning/dismissal shape as EditorContextMenu.vue (a separate
// component since the trigger condition — landing on a flagged misspelling
// — and action set are unrelated to Cut/Copy/Paste).
import { onBeforeUnmount, onMounted } from "vue";

const props = defineProps({
  left: { type: Number, required: true },
  top: { type: Number, required: true },
  word: { type: String, required: true },
  /** Up to five ranked suggestions (SPEC-035 AC-2); empty shows the
   * disabled "No suggestions" item (AC-3). */
  suggestions: { type: Array, default: () => [] },
});

const emit = defineEmits(["suggestion", "ignore", "ignore-all", "add-to-dictionary", "dismiss"]);

function menuStyle() {
  return { left: `${props.left}px`, top: `${props.top}px` };
}

function handleOutsidePointerDown() {
  emit("dismiss");
}

function handleKeyDown(event = new KeyboardEvent("keydown")) {
  if (event.key === "Escape") {
    emit("dismiss");
  }
}

onMounted(() => {
  document.addEventListener("mousedown", handleOutsidePointerDown);
  document.addEventListener("keydown", handleKeyDown);
});

onBeforeUnmount(() => {
  document.removeEventListener("mousedown", handleOutsidePointerDown);
  document.removeEventListener("keydown", handleKeyDown);
});
</script>

<template>
  <div
    class="context-menu"
    data-testid="spell-context-menu"
    role="menu"
    :aria-label="`Spelling suggestions for ${word}`"
    :style="menuStyle()"
    @mousedown.stop
  >
    <button
      v-for="(suggestion, index) in suggestions"
      :key="index"
      type="button"
      class="context-menu-item"
      role="menuitem"
      data-testid="spell-suggestion"
      @click="emit('suggestion', suggestion)"
    >
      <span>{{ suggestion }}</span>
    </button>
    <button
      v-if="suggestions.length === 0"
      type="button"
      class="context-menu-item"
      role="menuitem"
      disabled
      data-testid="spell-no-suggestions"
    >
      <span>No suggestions</span>
    </button>
    <div
      class="context-menu-separator"
      role="separator"
    />
    <button
      type="button"
      class="context-menu-item"
      role="menuitem"
      data-testid="spell-ignore"
      @click="emit('ignore')"
    >
      <span>Ignore</span>
    </button>
    <button
      type="button"
      class="context-menu-item"
      role="menuitem"
      data-testid="spell-ignore-all"
      @click="emit('ignore-all')"
    >
      <span>Ignore all</span>
    </button>
    <button
      type="button"
      class="context-menu-item"
      role="menuitem"
      data-testid="spell-add-to-dictionary"
      @click="emit('add-to-dictionary')"
    >
      <span>Add to dictionary</span>
    </button>
  </div>
</template>

<style scoped>
.context-menu {
  position: fixed;
  z-index: 40;
  min-width: 12rem;
  padding: var(--space-2);
  background: var(--surface-base);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-md);
  box-shadow: var(--shadow-2);
}

.context-menu-item {
  font: inherit;
  font-size: var(--text-md);
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: var(--space-6);
  width: 100%;
  padding: var(--space-3) var(--space-4);
  border: none;
  border-radius: var(--radius-sm);
  background: transparent;
  color: var(--text-primary);
  cursor: pointer;
  text-align: left;
}

.context-menu-item:hover:not(:disabled) {
  background: var(--surface-canvas);
}

.context-menu-item:disabled {
  color: var(--text-disabled);
  cursor: default;
}

.context-menu-separator {
  height: 1px;
  margin: var(--space-2) 0;
  background: var(--border-default);
}
</style>
