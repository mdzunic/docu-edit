<script setup lang="ts">
import UiButton from "./ui/UiButton.vue";
// The selected-field toolbar chip (SPEC-053 AC-14): the same five actions the
// field context menu offers — Update field, Edit field…, Toggle field codes,
// Unlink field and Lock/Unlock field — surfaced as a compact chip while the
// caret sits on a field. Emits the action id only; DocumentView runs it
// through the same handleFieldAction dispatcher the context menu uses.
// Runtime prop/emit declarations per the espree constraint (see AppShell.vue).
const props = defineProps({
  /** Whether the field under the caret is locked (Lock vs. Unlock label). */
  locked: { type: Boolean, default: false },
  /** Whether the field is read-only (unsupported): only "Toggle field codes"
     and "Unlink field" apply — it is never recomputed or edited. */
  readOnly: { type: Boolean, default: false },
});

const emit = defineEmits(["field-action"]);
</script>

<template>
  <div
    class="field-toolbar"
    role="toolbar"
    aria-label="Field actions"
    data-testid="field-toolbar"
  >
    <UiButton
      variant="ghost"
      data-testid="field-toolbar-update"
      :disabled="props.readOnly"
      @click="emit('field-action', 'update')"
    >
      Update field
    </UiButton>
    <UiButton
      variant="ghost"
      data-testid="field-toolbar-edit"
      :disabled="props.readOnly"
      @click="emit('field-action', 'edit')"
    >
      Edit field…
    </UiButton>
    <UiButton
      variant="ghost"
      data-testid="field-toolbar-toggle-codes"
      @click="emit('field-action', 'toggle-codes')"
    >
      Toggle field codes
    </UiButton>
    <UiButton
      variant="ghost"
      data-testid="field-toolbar-unlink"
      @click="emit('field-action', 'unlink')"
    >
      Unlink field
    </UiButton>
    <UiButton
      variant="ghost"
      data-testid="field-toolbar-lock"
      @click="emit('field-action', props.locked ? 'unlock' : 'lock')"
    >
      {{ props.locked ? "Unlock field" : "Lock field" }}
    </UiButton>
  </div>
</template>

<style scoped>
.field-toolbar {
  display: inline-flex;
  gap: var(--space-2);
  align-items: center;
  padding: var(--space-2) var(--space-3);
  background: var(--surface-canvas);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-md);
}
</style>
