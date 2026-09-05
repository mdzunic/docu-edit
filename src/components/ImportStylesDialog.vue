<script setup lang="ts">
import { computed, useTemplateRef } from "vue";
import { asStyleImportConflictRows, noStyleImportConflictRows } from "../lib/customStyles";
import BaseDialog from "./BaseDialog.vue";

// "Import styles from another document" (SPEC-042 AC-19..24): pick a
// .docx/.odt (AC-19) -> the parsed plan (conflicts to resolve, AC-21/AC-22;
// a non-conflicting count that merges automatically, AC-23) -> a summary
// once applied (AC-24). BaseDialog owns the shell and Escape — no field is
// marked [autofocus] (this dialog never focused one), so the panel itself
// takes focus and Escape works immediately. Runtime prop/emit declarations
// — see AppShell.vue's espree comment.
const props = defineProps({
  /** "pick" | "plan" | "summary" */
  phase: { type: String, default: "pick" },
  fileName: { type: String, default: "" },
  error: { type: String, default: "" },
  nonConflictingCount: { type: Number, default: 0 },
  /** Array<{ sourceId: string; name: string }> */
  conflicts: { type: Array, default: () => noStyleImportConflictRows() },
  /** Record<sourceId, "keep" | "replace" | "skip"> */
  resolutions: { type: Object, default: () => ({}) },
  /** { added: string[]; replaced: string[]; skipped: string[] } | null */
  summary: { type: Object, default: null },
});

const emit = defineEmits(["choose-file", "set-resolution", "import", "close"]);

const fileInput = useTemplateRef("fileInput");

function handleFileChange(event = new Event("change")) {
  const target = event.target;
  const file = target instanceof HTMLInputElement ? target.files?.[0] : undefined;
  if (file) {
    emit("choose-file", file);
  }
}

function triggerPick() {
  fileInput.value?.click();
}

const conflictRows = computed(() => asStyleImportConflictRows(props.conflicts));

function resolutionFor(sourceId = "") {
  return props.resolutions[sourceId] ?? "keep";
}

const canImport = computed(() => props.phase === "plan" && props.fileName.length > 0);
</script>

<template>
  <BaseDialog
    title="Import styles from another document"
    testid="import-styles-dialog"
    @dismiss="emit('close')"
  >
    <template v-if="phase === 'pick' || phase === 'plan'">
      <input
        ref="fileInput"
        type="file"
        tabindex="-1"
        aria-hidden="true"
        accept=".docx,.odt"
        class="hidden-file-input"
        data-testid="import-styles-file-input"
        @change="handleFileChange"
      >
      <button
        type="button"
        data-testid="import-styles-choose-file"
        @click="triggerPick"
      >
        {{ fileName || "Choose file…" }}
      </button>
      <p
        v-if="error"
        class="dialog-error"
        data-testid="import-styles-error"
      >
        {{ error }}
      </p>

      <div
        v-if="phase === 'plan'"
        class="import-plan"
      >
        <p data-testid="import-styles-non-conflicting-count">
          {{ nonConflictingCount }} style(s) will be added automatically.
        </p>
        <ul
          v-if="conflictRows.length > 0"
          class="conflict-list"
        >
          <li
            v-for="row in conflictRows"
            :key="row.sourceId"
            class="conflict-row"
            :data-testid="`import-conflict-${row.sourceId}`"
          >
            <span class="conflict-name">{{ row.name }}</span>
            <label>
              <input
                type="radio"
                :name="`resolution-${row.sourceId}`"
                value="keep"
                :checked="resolutionFor(row.sourceId) === 'keep'"
                :data-testid="`import-conflict-${row.sourceId}-keep`"
                @change="emit('set-resolution', { sourceId: row.sourceId, resolution: 'keep' })"
              >
              Keep existing
            </label>
            <label>
              <input
                type="radio"
                :name="`resolution-${row.sourceId}`"
                value="replace"
                :checked="resolutionFor(row.sourceId) === 'replace'"
                :data-testid="`import-conflict-${row.sourceId}-replace`"
                @change="emit('set-resolution', { sourceId: row.sourceId, resolution: 'replace' })"
              >
              Replace
            </label>
            <label>
              <input
                type="radio"
                :name="`resolution-${row.sourceId}`"
                value="skip"
                :checked="resolutionFor(row.sourceId) === 'skip'"
                :data-testid="`import-conflict-${row.sourceId}-skip`"
                @change="emit('set-resolution', { sourceId: row.sourceId, resolution: 'skip' })"
              >
              Skip
            </label>
          </li>
        </ul>
      </div>
    </template>

    <template v-else-if="phase === 'summary' && summary">
      <p data-testid="import-styles-summary">
        Added {{ summary.added.length }}, replaced {{ summary.replaced.length }},
        skipped {{ summary.skipped.length }}.
      </p>
    </template>

    <template #actions>
      <template v-if="phase === 'pick' || phase === 'plan'">
        <button
          type="button"
          data-testid="import-styles-cancel"
          @click="emit('close')"
        >
          Cancel
        </button>
        <button
          type="button"
          class="primary"
          data-testid="import-styles-import"
          :disabled="!canImport"
          @click="emit('import')"
        >
          Import
        </button>
      </template>
      <template v-else-if="phase === 'summary' && summary">
        <button
          type="button"
          class="primary"
          data-testid="import-styles-done"
          @click="emit('close')"
        >
          Done
        </button>
      </template>
    </template>
  </BaseDialog>
</template>

<style scoped>
/* This dialog sized its panel with a fixed width (not the usual min-width
   knob the base exposes), plus its own scroll cap for long conflict lists —
   both stay as overrides on the base's panel. */
:deep(.dialog) {
  width: 26rem;
  max-height: 85vh;
  overflow-y: auto;
}

.hidden-file-input {
  display: none;
}

.dialog-error {
  margin: var(--space-3) 0 0;
  color: var(--text-danger);
  font-size: var(--text-md);
}

.import-plan {
  margin-top: var(--space-5);
  font-size: var(--text-md);
}

.conflict-list {
  list-style: none;
  margin: var(--space-3) 0 0;
  padding: 0;
}

.conflict-row {
  display: flex;
  align-items: center;
  gap: var(--space-4);
  padding: var(--space-3) 0;
  border-bottom: var(--border-width) solid var(--border-subtle);
  flex-wrap: wrap;
}

.conflict-name {
  flex: 1 0 100%;
  font-weight: var(--weight-bold);
}

.conflict-row label {
  display: inline-flex;
  align-items: center;
  gap: var(--space-2);
}

.dialog-actions button {
  font: inherit;
  padding: var(--space-3) var(--space-5);
  border: var(--border-width) solid var(--border-strong);
  border-radius: var(--radius-sm);
  background: var(--surface-base);
  cursor: pointer;
}

.dialog-actions button.primary {
  background: var(--surface-brand);
  border-color: var(--surface-brand);
  color: var(--text-on-brand);
}

.dialog-actions button:disabled {
  opacity: 0.5;
  cursor: not-allowed;
}
</style>
