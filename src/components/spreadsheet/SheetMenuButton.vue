<script setup lang="ts">
// A sheet toolbar button that opens a popup (SPEC-123): the number-format
// dropdown, the rotation, borders, fill, cell-styles and Clear menus all hang
// off one of these — the trigger (an icon or a text value, `aria-haspopup`,
// `aria-expanded`) and a UiMenuSurface positioned by useFloatingMenu, which
// owns outside-click and Escape dismissal. The default slot is the popup's
// content and receives `close`. Runtime prop declarations per the espree
// constraint (see AppShell.vue).
import { ref, useTemplateRef } from "vue";
import { useFloatingMenu } from "../../composables/useFloatingMenu";
import AppIcon from "../AppIcon.vue";
import UiMenuSurface from "../ui/UiMenuSurface.vue";

const props = defineProps({
  /** Accessible name and tooltip. */
  label: { type: String, required: true },
  icon: { type: String, default: "" },
  /** A text value shown on the trigger (the number-format category). */
  text: { type: String, default: "" },
  testid: { type: String, required: true },
  menuTestid: { type: String, required: true },
  /** The popup's role ("menu" for item lists, "dialog" for swatch grids). */
  role: { type: String, default: "menu" },
  disabled: { type: Boolean, default: false },
});

const open = ref(false);
const trigger = useTemplateRef("trigger");
const surface = useTemplateRef("surface");
const { menuStyle } = useFloatingMenu({
  open: () => open.value,
  anchor: () => trigger.value,
  menu: () => surface.value?.$el ?? null,
});

function toggle() {
  if (!props.disabled) {
    open.value = !open.value;
  }
}

function close() {
  open.value = false;
}

defineExpose({ open: () => (open.value = true), close });
</script>

<template>
  <span class="sheet-menu-button">
    <button
      ref="trigger"
      type="button"
      class="trigger"
      :class="{ 'has-text': props.text !== '' }"
      :data-testid="props.testid"
      :aria-label="props.label"
      :title="props.label"
      :aria-haspopup="props.role === 'menu' ? 'menu' : 'dialog'"
      :aria-expanded="open ? 'true' : 'false'"
      :disabled="props.disabled"
      @mousedown.prevent
      @click="toggle"
    >
      <AppIcon
        v-if="props.icon"
        :name="props.icon"
      />
      <span
        v-if="props.text"
        class="text"
      >{{ props.text }}</span>
      <AppIcon
        name="chevron-down"
        class="chevron"
      />
    </button>
    <UiMenuSurface
      ref="surface"
      :open="open"
      :anchor="trigger"
      class="sheet-menu"
      :role="props.role"
      :aria-label="props.label"
      :data-testid="props.menuTestid"
      :style="menuStyle"
      @dismiss="close"
    >
      <slot :close="close" />
    </UiMenuSurface>
  </span>
</template>

<style scoped>
.sheet-menu-button {
  display: inline-flex;
}

.trigger {
  display: inline-flex;
  align-items: center;
  gap: var(--space-1);
  height: var(--control-height, var(--height-control));
  padding: 0 var(--space-1);
  font: inherit;
  font-size: var(--text-md);
  color: var(--text-primary);
  background: transparent;
  border: var(--border-width) solid transparent;
  border-radius: var(--radius-sm);
  cursor: pointer;
}

.trigger.has-text {
  min-width: 7rem;
  justify-content: space-between;
  border-color: var(--border-default);
}

.trigger:not(:disabled):hover {
  background: var(--state-hover);
}

.trigger:disabled {
  opacity: 0.5;
  cursor: default;
}

.trigger:focus-visible {
  outline: var(--space-1) solid var(--focus-ring);
  outline-offset: var(--space-1);
}

.chevron {
  width: 0.75rem;
  height: 0.75rem;
}

.sheet-menu {
  z-index: 30;
  min-width: 12rem;
  max-height: min(28rem, 80vh);
  overflow-y: auto;
}

.sheet-menu :deep([role="menuitem"]),
.sheet-menu :deep([role="menuitemradio"]) {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: var(--space-4);
  width: 100%;
  padding: var(--space-1) var(--space-3);
  font: inherit;
  font-size: var(--text-md);
  text-align: left;
  color: var(--text-primary);
  background: transparent;
  border: 0;
  border-radius: var(--radius-sm);
  cursor: pointer;
}

.sheet-menu :deep([role="menuitem"]:hover),
.sheet-menu :deep([role="menuitemradio"]:hover) {
  background: var(--state-hover);
}

.sheet-menu :deep(.item-hint) {
  color: var(--text-secondary);
  font-variant-numeric: tabular-nums;
}
</style>
