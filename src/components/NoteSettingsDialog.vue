<script setup lang="ts">
import { ref } from "vue";
import BaseDialog from "./BaseDialog.vue";
import UiButton from "./ui/UiButton.vue";

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
      <UiButton
        variant="secondary"
        data-testid="note-settings-cancel"
        @click="emit('cancel')"
      >
        Cancel
      </UiButton>
      <UiButton
        variant="primary"
        data-testid="note-settings-apply"
        @click="apply"
      >
        Apply
      </UiButton>
    </template>
  </BaseDialog>
</template>

<style scoped>
.field {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: var(--space-4);
  font-size: var(--text-md);
  color: var(--text-primary);
  margin-bottom: var(--space-3);
}

.field select,
.field-input {
  font: inherit;
  box-sizing: border-box;
  padding: var(--space-3) var(--space-3);
  border: var(--border-width) solid var(--border-strong);
  border-radius: var(--radius-sm);
  background: var(--surface-base);
}

.start-input {
  width: 6rem;
}

.restart-group {
  margin: var(--space-4) 0 0;
  padding: var(--space-3) var(--space-4) var(--space-3);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-sm);
  font-size: var(--text-md);
  color: var(--text-primary);
}

.restart-group legend {
  padding: 0 var(--space-2);
  font-size: var(--text-sm);
  color: var(--text-secondary);
}

.radio-field {
  display: flex;
  align-items: center;
  gap: var(--space-3);
  margin-bottom: var(--space-2);
}
</style>
