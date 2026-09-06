<script setup lang="ts">
// One command in a MenuBar popup (SPEC-049). A role="menuitem" button that
// MenuBar finds by role for keyboard focus. Its accelerator label is passed
// in already resolved from the SPEC-021 registry (AC-6) — this component
// never restates a binding itself. Inapplicable commands are `disabled`
// (present, not hidden — AC-7). tabindex is -1 so the menu bar stays a
// single tab stop (AC-8).
//
// Runtime prop/emit declarations per the espree constraint (see AppShell.vue).
import AppIcon from "./AppIcon.vue";

defineProps({
  label: { type: String, required: true },
  /** The handoff draws every menu row with a 16px glyph gutter, so labels line
   * up in one column whether or not a row has an icon, a check mark or a
   * submenu arrow. Named from the shared set (components/icons/paths.ts);
   * omitted rows still reserve the gutter, and an unknown name renders
   * nothing, so a typo is a blank gutter rather than a broken row. */
  icon: { type: String, default: "" },
  testid: { type: String, default: "" },
  /** Pre-formatted accelerator, e.g. "Ctrl+K"; "" shows none. */
  shortcut: { type: String, default: "" },
  disabled: { type: Boolean, default: false },
  /** SPEC-068 AC-12/13: when a command is disabled because it cannot work on
   * this screen size (not because it is inapplicable), the stated reason is
   * shown in the item — e.g. "not available on this screen size" — so the menu
   * stays a stable map of the product and nothing is silently dropped. */
  reason: { type: String, default: "" },
  /** A destructive command — "Delete document" sat in the File menu in the
   * same ink as Save as template and Preferences, one row from Close. The
   * confirmation dialog and the Trash make it recoverable, but a command that
   * removes something should still look unlike the ones that do not. Matches
   * the library row menu's own `is-danger`. */
  danger: { type: Boolean, default: false },
});

const emit = defineEmits(["select"]);
</script>

<template>
  <button
    type="button"
    role="menuitem"
    class="menu-item"
    :class="{ 'is-danger': danger }"
    :data-testid="testid || undefined"
    :disabled="disabled"
    :title="disabled && reason ? reason : undefined"
    tabindex="-1"
    @mousedown.prevent
    @click="emit('select')"
  >
    <span
      class="menu-icon"
      aria-hidden="true"
    ><AppIcon
      v-if="icon"
      :name="icon"
    /></span>
    <span class="menu-item-label">{{ label }}</span>
    <span
      v-if="disabled && reason"
      class="menu-item-reason"
      data-testid="menu-item-reason"
    >{{ reason }}</span>
    <span
      v-else-if="shortcut"
      class="menu-shortcut"
    >{{ shortcut }}</span>
  </button>
</template>

<style scoped>
.menu-item.is-danger:not(:disabled) .menu-item-label {
  color: var(--text-danger);
}

/* The stated reason a command is disabled on this screen size (SPEC-068):
   pushed to the trailing edge like a shortcut, but muted and italic so it
   reads as an explanation, not an accelerator. */
.menu-item-reason {
  margin-left: auto;
  padding-left: var(--space-6);
  color: var(--text-disabled);
  font-size: var(--text-xs);
  font-style: italic;
  white-space: nowrap;
}
</style>
