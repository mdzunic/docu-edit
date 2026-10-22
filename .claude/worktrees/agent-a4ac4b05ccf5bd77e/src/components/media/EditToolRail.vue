<script setup lang="ts">
// The image editor's tool rail (SPEC-112): crop & straighten, rotate/flip/
// resize, adjust, redact and annotate — native buttons in a vertical
// toolbar, the active one aria-pressed. Arrow keys move between them, as
// in any toolbar.
// Runtime prop declarations — see AppShell.vue.
import { computed, useTemplateRef } from "vue";
import { asImageEditor, EDIT_TOOL_LABELS, EDIT_TOOLS } from "../../views/media/useImageEditor";

const props = defineProps({
  /** The useImageEditor controller. */
  editor: { type: Object, required: true },
});

const e = computed(() => asImageEditor(props.editor));
const rail = useTemplateRef("rail");

function onKeydown(event = new KeyboardEvent("keydown")) {
  const forward = event.key === "ArrowDown" || event.key === "ArrowRight";
  const back = event.key === "ArrowUp" || event.key === "ArrowLeft";
  if (!forward && !back) {
    return;
  }
  const buttons = rail.value ? Array.from(rail.value.querySelectorAll("button")) : [];
  const at = buttons.findIndex((button) => button === document.activeElement);
  const next = buttons[(at + (forward ? 1 : -1) + buttons.length) % buttons.length];
  if (next) {
    event.preventDefault();
    next.focus();
  }
}
</script>

<template>
  <div
    ref="rail"
    class="tool-rail"
    role="toolbar"
    aria-label="Edit tools"
    aria-orientation="vertical"
    @keydown="onKeydown"
  >
    <button
      v-for="tool in EDIT_TOOLS"
      :key="tool"
      type="button"
      class="tool-button"
      :class="{ 'is-active': e.tool.value === tool }"
      :data-testid="`tool-${tool}`"
      :aria-pressed="e.tool.value === tool"
      @click="e.setTool(tool)"
    >
      {{ EDIT_TOOL_LABELS[tool] }}
    </button>
  </div>
</template>

<style scoped>
.tool-rail {
  flex: none;
  width: 9rem;
  display: flex;
  flex-direction: column;
  gap: var(--space-1);
  padding: var(--space-2);
  border-right: var(--border-width) solid var(--border-subtle);
  background: var(--surface-base);
}

.tool-button {
  font: inherit;
  font-size: var(--text-sm);
  text-align: left;
  min-height: var(--height-control);
  padding: var(--space-1) var(--space-2);
  border: var(--border-width) solid transparent;
  border-radius: var(--radius-sm);
  background: transparent;
  color: var(--text-primary);
  cursor: pointer;
}

.tool-button:hover {
  background: var(--state-hover);
}

.tool-button.is-active {
  background: var(--surface-selected);
  color: var(--text-accent);
  font-weight: var(--weight-bold);
}

.tool-button:focus-visible {
  outline: var(--space-1) solid var(--focus-ring);
  outline-offset: calc(-1 * var(--space-1));
}
</style>
