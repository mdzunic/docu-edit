<script setup lang="ts">
// The named field list (SPEC-046 "The named field list", AC-12, AC-30):
// shared between "Insert -> Merge field…" (every scalar/array name) and
// "Table -> Repeat row for each…" (array names only — the caller passes an
// already-narrowed `fields` list either way). A template can be written
// before its data exists, so a name not yet in the list can be typed too,
// validated against the same grammar chip insertion enforces.
import { computed, nextTick, onMounted, ref, useTemplateRef } from "vue";
import { isValidChipName } from "../lib/mergeChips";
import { asNamedFieldList } from "../lib/merge";

const props = defineProps({
  /** Dialog heading and aria-label. */
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
const input = useTemplateRef("input");

onMounted(async () => {
  await nextTick();
  input.value?.focus();
});

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
  <div
    class="dialog-backdrop"
    data-testid="merge-field-picker-dialog"
    @click.self="emit('cancel')"
  >
    <div
      class="dialog"
      role="dialog"
      aria-modal="true"
      :aria-label="props.heading"
    >
      <h2 class="dialog-title">
        {{ props.heading }}
      </h2>
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
          ref="input"
          v-model="customName"
          class="dialog-input"
          data-testid="merge-field-picker-custom-input"
          aria-label="Field name"
          placeholder="field_name"
          @keydown.enter="insertCustom"
          @keydown.escape="emit('cancel')"
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
      <div class="dialog-actions">
        <button
          type="button"
          data-testid="merge-field-picker-cancel"
          @click="emit('cancel')"
        >
          Cancel
        </button>
      </div>
    </div>
  </div>
</template>

<style scoped>
.dialog-backdrop {
  position: fixed;
  inset: 0;
  background: rgb(0 0 0 / 0.4);
  display: flex;
  align-items: center;
  justify-content: center;
  z-index: 50;
}

.dialog {
  background: #fff;
  border-radius: 6px;
  box-shadow: 0 4px 16px rgb(0 0 0 / 0.3);
  padding: 1.25rem 1.5rem;
  min-width: 20rem;
  max-width: 90vw;
  max-height: 80vh;
  overflow-y: auto;
  box-sizing: border-box;
}

.dialog-title {
  margin: 0 0 1rem;
  font-size: 1rem;
  font-weight: 600;
  color: #202124;
}

.field-list {
  list-style: none;
  margin: 0 0 1rem;
  padding: 0;
  max-height: 14rem;
  overflow-y: auto;
  border: 1px solid #e0e0e0;
  border-radius: 4px;
}

.field-list-empty {
  margin: 0 0 1rem;
  color: #5f6368;
  font-size: 0.875rem;
}

.field-option {
  font: inherit;
  display: block;
  width: 100%;
  text-align: left;
  padding: 0.5rem 0.75rem;
  border: none;
  border-bottom: 1px solid #f0f0f0;
  background: #fff;
  cursor: pointer;
  color: #202124;
}

.field-option:last-child {
  border-bottom: none;
}

.field-option:hover {
  background: #f1f3f4;
}

.custom-field-row {
  display: flex;
  gap: 0.5rem;
}

.dialog-input {
  font: inherit;
  flex: 1;
  box-sizing: border-box;
  padding: 0.375rem 0.5rem;
  border: 1px solid #bdbdbd;
  border-radius: 4px;
}

.dialog-error {
  margin: 0.5rem 0 0;
  color: #b3261e;
  font-size: 0.8125rem;
}

.dialog-actions {
  display: flex;
  justify-content: flex-end;
  gap: 0.5rem;
  margin-top: 1.25rem;
}

.dialog-actions button,
.custom-field-row button {
  font: inherit;
  padding: 0.375rem 1rem;
  border: 1px solid #bdbdbd;
  border-radius: 4px;
  background: #fff;
  cursor: pointer;
}

.custom-field-row button.primary {
  background: #2c3e50;
  border-color: #2c3e50;
  color: #fff;
}

.custom-field-row button:disabled {
  opacity: 0.5;
  cursor: not-allowed;
}
</style>
