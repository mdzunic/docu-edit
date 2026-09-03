<script setup lang="ts">
import { computed } from "vue";

// A toggle command in a MenuBar popup (SPEC-049): role="menuitemcheckbox"
// with a live check mark (AC-11). Used for the View menu's ruler / outline /
// comments toggles and the Edit menu's autocorrect switches. tabindex -1 so
// the menu bar remains one tab stop (AC-8). `checked` also takes the ARIA
// tri-state "mixed" — a group toggle (View → Rulers) whose members are only
// partly on — rendered as aria-checked="mixed" with a dash for its mark.
//
// Runtime prop/emit declarations per the espree constraint (see AppShell.vue).
const props = defineProps({
  label: { type: String, required: true },
  testid: { type: String, default: "" },
  checked: { type: [Boolean, String], default: false },
  disabled: { type: Boolean, default: false },
  shortcut: { type: String, default: "" },
  /** A trailing descriptive note, e.g. "Screen only" (SPEC-059 AC-71) — shown
   * in the shortcut gutter but styled as plain helper text, not a key. */
  hint: { type: String, default: "" },
  /** Why this toggle is disabled, e.g. "the caret is not in a table"
   * (SPEC-076 AC-18) — stated in the item, exactly as MenuItem states it, so
   * a disabled command is never silently greyed out. */
  reason: { type: String, default: "" },
});

const emit = defineEmits(["select"]);

const ariaChecked = computed(() =>
  props.checked === "mixed" ? "mixed" : props.checked === true ? "true" : "false",
);
const mark = computed(() => (props.checked === "mixed" ? "–" : props.checked === true ? "✓" : ""));
</script>

<template>
  <button
    type="button"
    role="menuitemcheckbox"
    class="menu-item"
    :data-testid="testid || undefined"
    :aria-checked="ariaChecked"
    :disabled="disabled"
    :title="disabled && reason ? reason : undefined"
    tabindex="-1"
    @mousedown.prevent
    @click="emit('select')"
  >
    <span
      class="menu-check"
      aria-hidden="true"
    >{{ mark }}</span>
    <span class="menu-item-label">{{ label }}</span>
    <span
      v-if="disabled && reason"
      class="menu-item-reason"
      data-testid="menu-item-reason"
    >{{ reason }}</span>
    <span
      v-else-if="hint"
      class="menu-hint"
    >{{ hint }}</span>
    <span
      v-if="shortcut && !(disabled && reason)"
      class="menu-shortcut"
    >{{ shortcut }}</span>
  </button>
</template>

<style scoped>
/* The stated reason a command is disabled (SPEC-076 AC-18), styled exactly
   as MenuItem states its own — muted, italic, at the trailing edge. */
.menu-item-reason {
  margin-left: auto;
  padding-left: var(--space-6);
  color: var(--text-disabled);
  font-size: var(--text-xs);
  font-style: italic;
  white-space: nowrap;
}

/* The trailing helper note (e.g. "Screen only", AC-71): pushed to the trailing
   edge like a shortcut but read as secondary descriptive text. */
.menu-hint {
  margin-left: auto;
  padding-left: var(--space-6);
  color: var(--text-secondary);
  font-size: var(--text-sm);
  font-style: italic;
}
</style>
