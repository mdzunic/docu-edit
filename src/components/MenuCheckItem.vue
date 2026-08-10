<script setup lang="ts">
// A toggle command in a MenuBar popup (SPEC-049): role="menuitemcheckbox"
// with a live check mark (AC-11). Used for the View menu's ruler / outline /
// comments toggles and the Edit menu's autocorrect switches. tabindex -1 so
// the menu bar remains one tab stop (AC-8).
//
// Runtime prop/emit declarations per the espree constraint (see AppShell.vue).
defineProps({
  label: { type: String, required: true },
  testid: { type: String, default: "" },
  checked: { type: Boolean, default: false },
  disabled: { type: Boolean, default: false },
  shortcut: { type: String, default: "" },
});

const emit = defineEmits(["select"]);
</script>

<template>
  <button
    type="button"
    role="menuitemcheckbox"
    class="menu-item"
    :data-testid="testid || undefined"
    :aria-checked="checked ? 'true' : 'false'"
    :disabled="disabled"
    tabindex="-1"
    @mousedown.prevent
    @click="emit('select')"
  >
    <span
      class="menu-check"
      aria-hidden="true"
    >{{ checked ? "✓" : "" }}</span>
    <span class="menu-item-label">{{ label }}</span>
    <span
      v-if="shortcut"
      class="menu-shortcut"
    >{{ shortcut }}</span>
  </button>
</template>
