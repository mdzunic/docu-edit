<script setup lang="ts">
// The measurement-unit picker shown when either ruler is right-clicked
// (SPEC-064 AC-10a): the four units, the current one checked. It emits the
// chosen unit — the view writes the document's `measurementUnit` preference,
// the same stored value the Page setup dialog's selector writes (AC-10b).
// Dismissed by a click anywhere outside or Escape, like EditorContextMenu.
import { onBeforeUnmount, onMounted } from "vue";

defineProps({
  /** Viewport coordinates of the right-click that opened the menu. */
  left: { type: Number, required: true },
  top: { type: Number, required: true },
  /** The document's current unit, shown checked. */
  unit: { type: String, default: "cm" },
});

const emit = defineEmits(["select", "dismiss"]);

const UNITS = [
  { key: "cm", label: "Centimetres" },
  { key: "mm", label: "Millimetres" },
  { key: "in", label: "Inches" },
  { key: "pt", label: "Points" },
];

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
    data-testid="ruler-unit-menu"
    role="menu"
    aria-label="Measurement unit"
    :style="{ left: `${left}px`, top: `${top}px` }"
    @mousedown.stop
  >
    <button
      v-for="option in UNITS"
      :key="option.key"
      type="button"
      class="context-menu-item"
      role="menuitemradio"
      :aria-checked="unit === option.key"
      :data-testid="`ruler-unit-${option.key}`"
      @click="emit('select', option.key)"
    >
      <span class="check">{{ unit === option.key ? "✓" : "" }}</span>
      <span>{{ option.label }}</span>
    </button>
  </div>
</template>

<style scoped>
.context-menu {
  position: fixed;
  z-index: 40;
  min-width: 12rem;
  padding: var(--space-1);
  background: var(--surface-base);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-md);
  box-shadow: var(--shadow-2);
}

.context-menu-item {
  display: flex;
  align-items: center;
  gap: var(--space-3);
  width: 100%;
  padding: var(--space-1) var(--space-3);
  border: none;
  border-radius: var(--radius-sm);
  background: transparent;
  color: var(--text-primary);
  font: inherit;
  font-size: var(--text-md);
  text-align: left;
  cursor: pointer;
}

.context-menu-item:hover {
  background: var(--state-hover);
}

.check {
  width: 1rem;
  color: var(--text-accent);
}
</style>
