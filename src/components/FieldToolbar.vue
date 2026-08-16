<script setup lang="ts">
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
    <button
      type="button"
      data-testid="field-toolbar-update"
      :disabled="props.readOnly"
      @click="emit('field-action', 'update')"
    >
      Update field
    </button>
    <button
      type="button"
      data-testid="field-toolbar-edit"
      :disabled="props.readOnly"
      @click="emit('field-action', 'edit')"
    >
      Edit field…
    </button>
    <button
      type="button"
      data-testid="field-toolbar-toggle-codes"
      @click="emit('field-action', 'toggle-codes')"
    >
      Toggle field codes
    </button>
    <button
      type="button"
      data-testid="field-toolbar-unlink"
      @click="emit('field-action', 'unlink')"
    >
      Unlink field
    </button>
    <button
      type="button"
      data-testid="field-toolbar-lock"
      @click="emit('field-action', props.locked ? 'unlock' : 'lock')"
    >
      {{ props.locked ? "Unlock field" : "Lock field" }}
    </button>
  </div>
</template>

<style scoped>
.field-toolbar {
  display: inline-flex;
  gap: 0.25rem;
  align-items: center;
  padding: 0.25rem 0.375rem;
  background: #f1f3f4;
  border: 1px solid #dadce0;
  border-radius: 6px;
}

.field-toolbar button {
  font: inherit;
  font-size: 0.8125rem;
  padding: 0.2rem 0.5rem;
  border: 1px solid transparent;
  border-radius: 4px;
  background: transparent;
  color: #202124;
  cursor: pointer;
}

.field-toolbar button:hover:not(:disabled) {
  background: #e0e0e0;
}

.field-toolbar button:disabled {
  opacity: 0.45;
  cursor: not-allowed;
}
</style>
