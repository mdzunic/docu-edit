<script setup lang="ts">
// The editor's context menu (SPEC-016 AC-4..6): Cut, Copy and Paste at the
// right-click point. Emits the chosen action only — the editor owns the
// clipboard work. Dismissed by a click anywhere outside or Escape.
import { onBeforeUnmount, onMounted } from "vue";
import { modifierKeyLabel } from "../lib/platform";

const props = defineProps({
  /** Viewport coordinates of the right-click that opened the menu. */
  left: { type: Number, required: true },
  top: { type: Number, required: true },
});

const emit = defineEmits(["cut", "copy", "paste", "dismiss"]);

const mod = modifierKeyLabel();

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
    data-testid="editor-context-menu"
    role="menu"
    aria-label="Editing"
    :style="menuStyle()"
    @mousedown.stop
  >
    <button
      type="button"
      class="context-menu-item"
      role="menuitem"
      data-testid="context-menu-cut"
      @click="emit('cut')"
    >
      <span>Cut</span>
      <span class="context-menu-shortcut">{{ mod }}+X</span>
    </button>
    <button
      type="button"
      class="context-menu-item"
      role="menuitem"
      data-testid="context-menu-copy"
      @click="emit('copy')"
    >
      <span>Copy</span>
      <span class="context-menu-shortcut">{{ mod }}+C</span>
    </button>
    <button
      type="button"
      class="context-menu-item"
      role="menuitem"
      data-testid="context-menu-paste"
      @click="emit('paste')"
    >
      <span>Paste</span>
      <span class="context-menu-shortcut">{{ mod }}+V</span>
    </button>
  </div>
</template>

<style scoped>
.context-menu {
  position: fixed;
  z-index: 40;
  min-width: 12rem;
  padding: 0.25rem;
  background: #fff;
  border: 1px solid #dadce0;
  border-radius: 6px;
  box-shadow: 0 2px 10px rgb(0 0 0 / 0.25);
}

.context-menu-item {
  font: inherit;
  font-size: 0.875rem;
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 1.5rem;
  width: 100%;
  padding: 0.375rem 0.75rem;
  border: none;
  border-radius: 4px;
  background: transparent;
  color: #202124;
  cursor: pointer;
  text-align: left;
}

.context-menu-item:hover {
  background: #f1f3f4;
}

.context-menu-shortcut {
  color: #5f6368;
  font-size: 0.75rem;
}
</style>
