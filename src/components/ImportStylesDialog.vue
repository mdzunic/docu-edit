<script setup lang="ts">
import { computed, useTemplateRef } from "vue";
import { asStyleImportConflictRows, noStyleImportConflictRows } from "../lib/customStyles";

// "Import styles from another document" (SPEC-042 AC-19..24): pick a
// .docx/.odt (AC-19) -> the parsed plan (conflicts to resolve, AC-21/AC-22;
// a non-conflicting count that merges automatically, AC-23) -> a summary
// once applied (AC-24). Runtime prop/emit declarations — see AppShell.vue's
// espree comment.
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
  <div
    class="dialog-backdrop"
    data-testid="import-styles-dialog"
    @click.self="emit('close')"
    @keydown.escape="emit('close')"
  >
    <div
      class="dialog"
      role="dialog"
      aria-modal="true"
      aria-label="Import styles from another document"
    >
      <h2 class="dialog-title">
        Import styles from another document
      </h2>

      <template v-if="phase === 'pick' || phase === 'plan'">
        <input
          ref="fileInput"
          type="file"
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

        <div class="dialog-actions">
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
        </div>
      </template>

      <template v-else-if="phase === 'summary' && summary">
        <p data-testid="import-styles-summary">
          Added {{ summary.added.length }}, replaced {{ summary.replaced.length }},
          skipped {{ summary.skipped.length }}.
        </p>
        <div class="dialog-actions">
          <button
            type="button"
            class="primary"
            data-testid="import-styles-done"
            @click="emit('close')"
          >
            Done
          </button>
        </div>
      </template>
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
  width: 26rem;
  max-width: 90vw;
  max-height: 85vh;
  overflow-y: auto;
  box-sizing: border-box;
}

.dialog-title {
  margin: 0 0 1rem;
  font-size: 1rem;
  font-weight: 600;
  color: #202124;
}

.hidden-file-input {
  display: none;
}

.dialog-error {
  margin: 0.5rem 0 0;
  color: #b3261e;
  font-size: 0.8125rem;
}

.import-plan {
  margin-top: 1rem;
  font-size: 0.8125rem;
}

.conflict-list {
  list-style: none;
  margin: 0.5rem 0 0;
  padding: 0;
}

.conflict-row {
  display: flex;
  align-items: center;
  gap: 0.75rem;
  padding: 0.375rem 0;
  border-bottom: 1px solid #e8eaed;
  flex-wrap: wrap;
}

.conflict-name {
  flex: 1 0 100%;
  font-weight: 600;
}

.conflict-row label {
  display: inline-flex;
  align-items: center;
  gap: 0.25rem;
}

.dialog-actions {
  display: flex;
  justify-content: flex-end;
  gap: 0.5rem;
  margin-top: 1.25rem;
}

.dialog-actions button {
  font: inherit;
  padding: 0.375rem 1rem;
  border: 1px solid #bdbdbd;
  border-radius: 4px;
  background: #fff;
  cursor: pointer;
}

.dialog-actions button.primary {
  background: #2c3e50;
  border-color: #2c3e50;
  color: #fff;
}

.dialog-actions button:disabled {
  opacity: 0.5;
  cursor: not-allowed;
}
</style>
