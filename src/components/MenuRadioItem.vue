<script setup lang="ts">
// A single-choice command in a MenuBar popup (SPEC-059): role="menuitemradio"
// with a live selection mark, for the View → Appearance theme choices
// (System / Light / Dark). Mirrors MenuCheckItem.vue but carries radio
// semantics so assistive tech reads the group as a single-select set and the
// current choice shows (AC-48). tabindex -1 keeps the menu bar one tab stop.
//
// Runtime prop/emit declarations per the espree constraint (see AppShell.vue).
defineProps({
  label: { type: String, required: true },
  testid: { type: String, default: "" },
  checked: { type: Boolean, default: false },
  disabled: { type: Boolean, default: false },
});

const emit = defineEmits(["select"]);
</script>

<template>
  <button
    type="button"
    role="menuitemradio"
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
    >{{ checked ? "●" : "" }}</span>
    <span class="menu-item-label">{{ label }}</span>
  </button>
</template>
