<script setup lang="ts">
// One command in a MenuBar popup (SPEC-049). A role="menuitem" button that
// MenuBar finds by role for keyboard focus. Its accelerator label is passed
// in already resolved from the SPEC-021 registry (AC-6) — this component
// never restates a binding itself. Inapplicable commands are `disabled`
// (present, not hidden — AC-7). tabindex is -1 so the menu bar stays a
// single tab stop (AC-8).
//
// Runtime prop/emit declarations per the espree constraint (see AppShell.vue).
defineProps({
  label: { type: String, required: true },
  testid: { type: String, default: "" },
  /** Pre-formatted accelerator, e.g. "Ctrl+K"; "" shows none. */
  shortcut: { type: String, default: "" },
  disabled: { type: Boolean, default: false },
});

const emit = defineEmits(["select"]);
</script>

<template>
  <button
    type="button"
    role="menuitem"
    class="menu-item"
    :data-testid="testid || undefined"
    :disabled="disabled"
    tabindex="-1"
    @mousedown.prevent
    @click="emit('select')"
  >
    <span class="menu-item-label">{{ label }}</span>
    <span
      v-if="shortcut"
      class="menu-shortcut"
    >{{ shortcut }}</span>
  </button>
</template>
