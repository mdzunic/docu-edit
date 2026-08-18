<script setup lang="ts">
// The Tabs dialog (SPEC-032 AC-25): lists the selection's custom tab stops
// and adds/removes them — position in cm, alignment (left/centre/right/
// decimal, AC-19..22) and leader (AC-23). OK emits the full list; the
// setTabStops command normalises and applies it. BaseDialog owns the shell
// and Escape; @dismiss maps onto cancel. Runtime prop/emit declarations —
// this script is parsed by espree.
import { computed, ref } from "vue";
import { asTabStopAlignment, asTabStops } from "../lib/documentModel";
import BaseDialog from "./BaseDialog.vue";

const props = defineProps({
  /** The selection's current stops (commands.selectionTabStops). */
  initial: { type: Array, default: () => [] },
  /** The document's default tab interval in cm (SPEC-054 AC-10, AC-13/14) —
   * commands.documentDefaultTabStopCm, 1.27 when the document sets none. */
  defaultIntervalCm: { type: Number, default: 1.27 },
});

const emit = defineEmits(["apply", "cancel"]);

const stops = ref(
  asTabStops(props.initial).map((stop) => ({
    positionCm: stop.positionCm,
    alignment: stop.alignment,
    leader: stop.leader ?? "",
  })),
);

const position = ref("");
const alignment = ref("left");
const leader = ref("");
// SPEC-054 AC-10: the per-document default tab interval, editable here.
const defaultInterval = ref(String(props.defaultIntervalCm));

const defaultIntervalError = computed(() => {
  const parsed = Number(defaultInterval.value);
  if (defaultInterval.value === "" || !Number.isFinite(parsed) || parsed <= 0) {
    return "Default interval must be a positive centimetre value";
  }
  return "";
});

const positionError = computed(() => {
  if (position.value === "") {
    return "";
  }
  const parsed = Number(position.value);
  if (!Number.isFinite(parsed) || parsed < 0) {
    return "Position must be a non-negative centimetre value";
  }
  return "";
});

function describe(stop = { positionCm: 0, alignment: "left", leader: "" }) {
  const parts = [
    `${Math.round(stop.positionCm * 100) / 100} cm`,
    stop.alignment === "center" ? "centre" : stop.alignment,
  ];
  if (stop.leader !== "") {
    parts.push(stop.leader);
  }
  return parts.join(" · ");
}

function addStop() {
  const parsed = Number(position.value);
  if (position.value === "" || !Number.isFinite(parsed) || parsed < 0) {
    return;
  }
  stops.value = stops.value
    .filter((stop) => stop.positionCm !== parsed)
    .concat([
      { positionCm: parsed, alignment: asTabStopAlignment(alignment.value), leader: leader.value },
    ])
    .sort((a, b) => a.positionCm - b.positionCm);
  position.value = "";
}

function removeStop(index = 0) {
  stops.value = stops.value.filter((unused, i) => i !== index);
}

function clearAll() {
  stops.value = [];
}

function apply() {
  const parsedInterval = Number(defaultInterval.value);
  emit("apply", {
    stops: stops.value.map((stop) => ({
      positionCm: stop.positionCm,
      alignment: stop.alignment,
      ...(stop.leader !== "" ? { leader: stop.leader } : {}),
    })),
    defaultIntervalCm:
      defaultIntervalError.value === "" ? parsedInterval : props.defaultIntervalCm,
  });
}
</script>

<template>
  <BaseDialog
    title="Tabs"
    testid="tabs-dialog"
    width="26rem"
    @dismiss="emit('cancel')"
  >
    <div class="field-row">
      <label class="field">
        <span class="field-label">Position (cm)</span>
        <input
          v-model="position"
          class="field-input"
          data-testid="tab-position-input"
          type="number"
          min="0"
          step="0.25"
          @keydown.enter="addStop"
        >
      </label>
      <label class="field">
        <span class="field-label">Alignment</span>
        <select
          v-model="alignment"
          class="field-input"
          data-testid="tab-alignment-select"
        >
          <option value="left">Left</option>
          <option value="center">Centre</option>
          <option value="right">Right</option>
          <option value="decimal">Decimal</option>
          <option value="bar">Bar</option>
        </select>
      </label>
      <label class="field">
        <span class="field-label">Leader</span>
        <select
          v-model="leader"
          class="field-input"
          data-testid="tab-leader-select"
        >
          <option value="">None</option>
          <option value="dot">Dot (…)</option>
          <option value="hyphen">Hyphen (---)</option>
          <option value="underscore">Underscore (___)</option>
        </select>
      </label>
      <div class="field add-field">
        <span class="field-label">&nbsp;</span>
        <button
          type="button"
          class="add-button"
          data-testid="tab-add"
          :disabled="position === '' || !!positionError"
          @click="addStop"
        >
          Add
        </button>
      </div>
    </div>
    <p
      v-if="positionError"
      class="dialog-error"
      data-testid="tab-position-error"
    >
      {{ positionError }}
    </p>
    <ul
      class="stop-list"
      data-testid="tab-stop-list"
    >
      <li
        v-for="(stop, index) in stops"
        :key="`${stop.positionCm}-${index}`"
        class="stop-entry"
        data-testid="tab-stop-entry"
      >
        <span>{{ describe(stop) }}</span>
        <button
          type="button"
          class="remove-button"
          :data-testid="`tab-remove-${index}`"
          :aria-label="`Remove tab stop at ${stop.positionCm} cm`"
          @click="removeStop(index)"
        >
          Remove
        </button>
      </li>
      <li
        v-if="stops.length === 0"
        class="stop-empty"
      >
        No custom tab stops
      </li>
    </ul>
    <label class="field default-interval-field">
      <span class="field-label">Default tab interval (cm)</span>
      <input
        v-model="defaultInterval"
        class="field-input"
        data-testid="tab-default-interval-input"
        type="number"
        min="0"
        step="0.25"
      >
    </label>
    <p
      v-if="defaultIntervalError"
      class="dialog-error"
      data-testid="tab-default-interval-error"
    >
      {{ defaultIntervalError }}
    </p>
    <!-- Deliberately NOT the base's #actions slot: this row is Clear all on
         the left, a spacer, then Cancel/OK — align-items: center without
         justify-content: flex-end, a real layout the base's flex-end
         actions row cannot express. It stays in the body with its own
         scoped .dialog-actions styles. -->
    <div class="dialog-actions">
      <button
        type="button"
        data-testid="tabs-clear-all"
        :disabled="stops.length === 0"
        @click="clearAll"
      >
        Clear all
      </button>
      <span class="actions-spacer" />
      <button
        type="button"
        data-testid="tabs-cancel"
        @click="emit('cancel')"
      >
        Cancel
      </button>
      <button
        type="button"
        class="primary"
        data-testid="tabs-apply"
        @click="apply"
      >
        OK
      </button>
    </div>
  </BaseDialog>
</template>

<style scoped>
.field-row {
  display: flex;
  gap: var(--space-4);
  margin-bottom: var(--space-4);
  align-items: flex-end;
}

.field {
  flex: 1;
  display: flex;
  flex-direction: column;
  gap: var(--space-2);
}

.add-field {
  flex: 0 0 auto;
}

.default-interval-field {
  margin-top: var(--space-4);
  max-width: 14rem;
}

.field-label {
  font-size: var(--text-md);
  color: var(--text-primary);
}

.field-input {
  font: inherit;
  box-sizing: border-box;
  width: 100%;
  padding: var(--space-3) var(--space-3);
  border: var(--border-width) solid var(--border-strong);
  border-radius: var(--radius-sm);
  background: var(--surface-base);
}

.add-button,
.remove-button {
  font: inherit;
  padding: var(--space-3) var(--space-4);
  border: var(--border-width) solid var(--border-strong);
  border-radius: var(--radius-sm);
  background: var(--surface-base);
  cursor: pointer;
}

.add-button:disabled {
  opacity: 0.5;
  cursor: not-allowed;
}

.stop-list {
  list-style: none;
  margin: 0;
  padding: 0;
  max-height: 12rem;
  overflow-y: auto;
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-sm);
}

.stop-entry {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: var(--space-5);
  padding: var(--space-3) var(--space-4);
  border-bottom: var(--border-width) solid var(--border-subtle);
  font-size: var(--text-md);
}

.stop-entry:last-child {
  border-bottom: none;
}

.stop-empty {
  padding: var(--space-3) var(--space-4);
  color: var(--text-secondary);
  font-size: var(--text-md);
}

.dialog-error {
  margin: var(--space-3) 0 0;
  color: var(--text-danger);
  font-size: var(--text-md);
}

/* This dialog's own actions row (see the template comment) — lives in the
   body slot, so these scoped rules apply to it directly. */
.dialog-actions {
  display: flex;
  align-items: center;
  gap: var(--space-3);
  margin-top: var(--space-6);
}

.actions-spacer {
  flex: 1;
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
