<script setup lang="ts">
// SPEC-121 "Tab colour": the core TextColorPicker — the theme row from the
// workbook's theme, the standard palette, recent colours — opened at a
// viewport point (the tab menu's), plus a "No colour" row inside its menu.
// Its trigger is laid out at the point but never shown, so the menu opens
// just above it (the tabs sit at the foot of the view). Emits what the
// picker emits for a swatch (a hex, or a ThemeColorRef kept as a reference),
// `clear` for "No colour", and `close` when dismissed without a choice.
// Runtime prop declarations per the espree constraint (see AppShell.vue).
import { nextTick, onMounted, onUnmounted, useTemplateRef } from "vue";
import TextColorPicker from "../TextColorPicker.vue";

const props = defineProps({
  /** The viewport point the picker opens at. */
  x: { type: Number, default: 0 },
  y: { type: Number, default: 0 },
  /** The workbook's DocumentTheme (its theme row); null for none. */
  theme: { type: Object, default: null },
  /** Recently applied colours, most recent first. */
  recent: { type: Array, default: () => [] },
});

const emit = defineEmits(["select", "clear", "close"]);

const root = useTemplateRef("root");
const picker = useTemplateRef("picker");
const previouslyFocused = document.activeElement instanceof HTMLElement ? document.activeElement : null;

onMounted(async () => {
  picker.value?.openMenu();
  await nextTick();
  const none = root.value?.querySelector('[data-testid="sheet-tab-color-none"]');
  if (none instanceof HTMLElement) {
    none.focus();
  }
});

onUnmounted(() => {
  const active = document.activeElement;
  if (previouslyFocused?.isConnected && (active === null || active === document.body)) {
    previouslyFocused.focus();
  }
});
</script>

<template>
  <div
    ref="root"
    class="sheet-tab-color-picker"
    data-testid="sheet-tab-color-picker"
    :style="{ left: `${props.x}px`, top: `${props.y}px` }"
  >
    <TextColorPicker
      ref="picker"
      :theme="props.theme"
      :recent="props.recent"
      placement="above"
      @select="emit('select', $event)"
      @close="emit('close')"
    >
      <template #extra>
        <button
          type="button"
          class="tab-color-none"
          data-testid="sheet-tab-color-none"
          @mousedown.prevent
          @click="emit('clear')"
        >
          No colour
        </button>
      </template>
    </TextColorPicker>
  </div>
</template>

<style scoped>
/* A zero-height anchor at the point: the trigger keeps its place in layout
   (the menu measures it) but is neither seen nor reachable. */
.sheet-tab-color-picker {
  --control-height: 0;
  position: fixed;
  z-index: 61;
}

.sheet-tab-color-picker :deep([data-testid="text-color-button"]) {
  visibility: hidden;
  overflow: hidden;
}

.tab-color-none {
  display: block;
  width: 100%;
  margin-bottom: var(--space-2);
  padding: var(--space-1) var(--space-3);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-sm);
  background: var(--surface-base);
  color: var(--text-primary);
  font: inherit;
  font-size: var(--text-md);
  text-align: left;
  cursor: pointer;
}

.tab-color-none:hover {
  background: var(--state-hover);
}

.tab-color-none:focus-visible {
  outline: var(--space-1) solid var(--focus-ring);
  outline-offset: 1px;
}
</style>
