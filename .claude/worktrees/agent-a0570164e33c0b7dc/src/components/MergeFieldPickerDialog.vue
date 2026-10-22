<script setup lang="ts">
// The named field list (SPEC-046 "The named field list", AC-12, AC-30):
// shared between "Insert -> Merge field…" (every scalar/array name) and
// "Table -> Repeat row for each…" (array names only — the caller passes an
// already-narrowed `fields` list either way). A template can be written
// before its data exists, so a name not yet in the list can be typed too,
// validated against the same grammar chip insertion enforces. BaseDialog
// owns the shell, Escape and initial focus (the [autofocus] name input) —
// which also fixes the old bug where Escape was bound to the input only
// and did nothing while a button or the field list had focus.
import { computed, ref } from "vue";
import { isValidChipName } from "../lib/mergeChips";
import { asNamedFieldList } from "../lib/merge";
import BaseDialog from "./BaseDialog.vue";

const props = defineProps({
  /** Dialog heading and accessible name. */
  heading: { type: String, required: true },
  /** { name, array }[] — in "The named field list"'s order: source columns
   * first, then names already used by a chip in this document. */
  fields: { type: Array, default: () => [] },
});

const emit = defineEmits(["insert", "cancel"]);

// Runtime Array prop carries no element type (the espree constraint bars
// PropType<NamedField[]> here) — this typed pass-through gives the template
// below typed access.
const fields = computed(() => asNamedFieldList(props.fields));

const customName = ref("");

const trimmed = computed(() => customName.value.trim());
const customNameError = computed(() => {
  if (trimmed.value.length === 0) {
    return "";
  }
  return isValidChipName(trimmed.value) ? "" : "A field name starts with a letter or underscore and holds only letters, digits, underscores and dots.";
});

function choose(name = "") {
  emit("insert", name);
}

function insertCustom() {
  if (trimmed.value.length === 0 || customNameError.value) {
    return;
  }
  emit("insert", trimmed.value);
}
</script>

<template>
  <BaseDialog
    :title="props.heading"
    testid="merge-field-picker-dialog"
    @dismiss="emit('cancel')"
  >
    <ul
      v-if="fields.length > 0"
      class="field-list"
      data-testid="merge-field-picker-list"
    >
      <li
        v-for="field in fields"
        :key="field.name"
      >
        <button
          type="button"
          class="field-option"
          data-testid="merge-field-picker-option"
          @click="choose(field.name)"
        >
          {{ field.name }}
        </button>
      </li>
    </ul>
    <p
      v-else
      class="field-list-empty"
    >
      No fields yet — type a name below.
    </p>
    <div class="custom-field-row">
      <input
        v-model="customName"
        class="dialog-input"
        data-testid="merge-field-picker-custom-input"
        aria-label="Field name"
        placeholder="field_name"
        autofocus
        @keydown.enter="insertCustom"
      >
      <button
        type="button"
        class="primary"
        data-testid="merge-field-picker-custom-insert"
        :disabled="trimmed.length === 0 || !!customNameError"
        @click="insertCustom"
      >
        Insert
      </button>
    </div>
    <p
      v-if="customNameError"
      class="dialog-error"
      data-testid="merge-field-picker-error"
    >
      {{ customNameError }}
    </p>
    <template #actions>
      <button
        type="button"
        data-testid="merge-field-picker-cancel"
        @click="emit('cancel')"
      >
        Cancel
      </button>
    </template>
  </BaseDialog>
</template>

<style scoped>
/* The field list can be long, so the panel keeps its own scroll cap. */
:deep(.dialog) {
  max-height: 80vh;
  overflow-y: auto;
}

.field-list {
  list-style: none;
  margin: 0 0 var(--space-5);
  padding: 0;
  max-height: 14rem;
  overflow-y: auto;
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-sm);
}

.field-list-empty {
  margin: 0 0 var(--space-5);
  color: var(--text-secondary);
  font-size: var(--text-md);
}

.field-option {
  font: inherit;
  display: block;
  width: 100%;
  text-align: left;
  padding: var(--space-3) var(--space-4);
  border: none;
  border-bottom: var(--border-width) solid var(--border-subtle);
  background: var(--surface-base);
  cursor: pointer;
  color: var(--text-primary);
}

.field-option:last-child {
  border-bottom: none;
}

.field-option:hover {
  background: var(--surface-canvas);
}

.custom-field-row {
  display: flex;
  gap: var(--space-3);
}

.dialog-input {
  font: inherit;
  flex: 1;
  box-sizing: border-box;
  padding: var(--space-3) var(--space-3);
  border: var(--border-width) solid var(--border-strong);
  border-radius: var(--radius-sm);
}

.dialog-error {
  margin: var(--space-3) 0 0;
  color: var(--text-danger);
  font-size: var(--text-md);
}

.dialog-actions button,
.custom-field-row button {
  font: inherit;
  padding: var(--space-3) var(--space-5);
  border: var(--border-width) solid var(--border-strong);
  border-radius: var(--radius-sm);
  background: var(--surface-base);
  cursor: pointer;
}

.custom-field-row button.primary {
  background: var(--surface-brand);
  border-color: var(--surface-brand);
  color: var(--text-on-brand);
}

.custom-field-row button:disabled {
  opacity: 0.5;
  cursor: not-allowed;
}
</style>
