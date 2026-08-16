<script setup lang="ts">
import { ref } from "vue";
import BaseDialog from "./BaseDialog.vue";

// The footnote settings dialog (SPEC-031 AC-17..20): numbering format
// (1/i/a/*), start number, and restart behaviour (continuous, per page,
// per section) — applied to footnotes or, via the stream selector at the
// top (Word's own dialog shape), to endnotes. BaseDialog owns the shell,
// Escape and initial focus (the [autofocus] number-format select, as
// before — not the stream selector above it).
const props = defineProps({
  /** Prefill: { kind, footnote: { numberFormat, startAt, restart },
   * endnote: { numberFormat, startAt, restart } } — the effective (default-
   * resolved) settings of both streams. */
  initial: { type: Object, required: true },
});

const emit = defineEmits(["apply", "cancel"]);

const kind = ref(String(props.initial.kind ?? "footnote"));
const settingsOf = (which = "footnote") =>
  (which === "footnote" ? props.initial.footnote : props.initial.endnote) ?? {};
const numberFormat = ref(String(settingsOf(kind.value).numberFormat ?? "decimal"));
const startAt = ref(String(settingsOf(kind.value).startAt ?? 1));
const restart = ref(String(settingsOf(kind.value).restart ?? "continuous"));

function switchKind(next = "footnote") {
  kind.value = next;
  const settings = settingsOf(next);
  numberFormat.value = String(settings.numberFormat ?? (next === "footnote" ? "decimal" : "lowerRoman"));
  startAt.value = String(settings.startAt ?? 1);
  restart.value = String(settings.restart ?? "continuous");
}

function onKindChange(event = new Event("change")) {
  const target = event.target;
  if (target instanceof HTMLSelectElement) {
    switchKind(target.value);
  }
}

function apply() {
  const parsed = Number(startAt.value);
  emit("apply", {
    kind: kind.value,
    numberFormat: numberFormat.value,
    startAt: Number.isFinite(parsed) ? Math.max(1, Math.floor(parsed)) : 1,
    restart: restart.value,
  });
}
</script>

<template>
  <BaseDialog
    title="Footnote settings"
    testid="note-settings-dialog"
    width="22rem"
    @dismiss="emit('cancel')"
  >
    <label class="field">
      <span>Apply to</span>
      <select
        data-testid="note-settings-kind"
        :value="kind"
        @change="onKindChange"
      >
        <option value="footnote">Footnotes</option>
        <option value="endnote">Endnotes</option>
      </select>
    </label>
    <label class="field">
      <span>Number format</span>
      <select
        v-model="numberFormat"
        data-testid="note-settings-format"
        autofocus
      >
        <option value="decimal">1, 2, 3, …</option>
        <option value="lowerRoman">i, ii, iii, …</option>
        <option value="lowerAlpha">a, b, c, …</option>
        <option value="symbol">*, †, ‡, §, …</option>
      </select>
    </label>
    <label class="field">
      <span>Start at</span>
      <input
        v-model="startAt"
        class="field-input start-input"
        data-testid="note-settings-start-at"
        type="number"
        min="1"
        step="1"
      >
    </label>
    <fieldset class="restart-group">
      <legend>Numbering</legend>
      <label class="radio-field">
        <input
          v-model="restart"
          type="radio"
          value="continuous"
          data-testid="note-settings-restart-continuous"
        >
        <span>Continuous</span>
      </label>
      <label class="radio-field">
        <input
          v-model="restart"
          type="radio"
          value="page"
          data-testid="note-settings-restart-page"
        >
        <span>Restart each page</span>
      </label>
      <label class="radio-field">
        <input
          v-model="restart"
          type="radio"
          value="section"
          data-testid="note-settings-restart-section"
        >
        <span>Restart each section</span>
      </label>
    </fieldset>
    <template #actions>
      <button
        type="button"
        data-testid="note-settings-cancel"
        @click="emit('cancel')"
      >
        Cancel
      </button>
      <button
        type="button"
        class="primary"
        data-testid="note-settings-apply"
        @click="apply"
      >
        Apply
      </button>
    </template>
  </BaseDialog>
</template>

<style scoped>
.field {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 0.75rem;
  font-size: 0.8125rem;
  color: #3c4043;
  margin-bottom: 0.625rem;
}

.field select,
.field-input {
  font: inherit;
  box-sizing: border-box;
  padding: 0.375rem 0.5rem;
  border: 1px solid #bdbdbd;
  border-radius: 4px;
  background: #fff;
}

.start-input {
  width: 6rem;
}

.restart-group {
  margin: 0.75rem 0 0;
  padding: 0.5rem 0.75rem 0.625rem;
  border: 1px solid #e0e0e0;
  border-radius: 4px;
  font-size: 0.8125rem;
  color: #3c4043;
}

.restart-group legend {
  padding: 0 0.25rem;
  font-size: 0.75rem;
  color: #5f6368;
}

.radio-field {
  display: flex;
  align-items: center;
  gap: 0.5rem;
  margin-bottom: 0.25rem;
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
</style>
