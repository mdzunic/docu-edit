<script setup lang="ts">
// The Tabs dialog (SPEC-032 AC-25): lists the selection's custom tab stops
// and adds/removes them — position in cm, alignment (left/centre/right/
// decimal, AC-19..22) and leader (AC-23). OK emits the full list; the
// setTabStops command normalises and applies it. Runtime prop/emit
// declarations — this script is parsed by espree.
import { computed, ref } from "vue";
import { asTabStopAlignment, asTabStops } from "../lib/documentModel";

const props = defineProps({
  /** The selection's current stops (commands.selectionTabStops). */
  initial: { type: Array, default: () => [] },
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
  emit(
    "apply",
    stops.value.map((stop) => ({
      positionCm: stop.positionCm,
      alignment: stop.alignment,
      ...(stop.leader !== "" ? { leader: stop.leader } : {}),
    })),
  );
}
</script>

<template>
  <div
    class="dialog-backdrop"
    data-testid="tabs-dialog"
    @click.self="emit('cancel')"
    @keydown.escape="emit('cancel')"
  >
    <div
      class="dialog"
      role="dialog"
      aria-modal="true"
      aria-label="Tabs"
    >
      <h2 class="dialog-title">
        Tabs
      </h2>
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
  min-width: 26rem;
  max-width: 90vw;
  box-sizing: border-box;
}

.dialog-title {
  margin: 0 0 1rem;
  font-size: 1rem;
  font-weight: 600;
  color: #202124;
}

.field-row {
  display: flex;
  gap: 0.75rem;
  margin-bottom: 0.75rem;
  align-items: flex-end;
}

.field {
  flex: 1;
  display: flex;
  flex-direction: column;
  gap: 0.25rem;
}

.add-field {
  flex: 0 0 auto;
}

.field-label {
  font-size: 0.8125rem;
  color: #3c4043;
}

.field-input {
  font: inherit;
  box-sizing: border-box;
  width: 100%;
  padding: 0.375rem 0.5rem;
  border: 1px solid #bdbdbd;
  border-radius: 4px;
  background: #fff;
}

.add-button,
.remove-button {
  font: inherit;
  padding: 0.375rem 0.75rem;
  border: 1px solid #bdbdbd;
  border-radius: 4px;
  background: #fff;
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
  border: 1px solid #e0e0e0;
  border-radius: 4px;
}

.stop-entry {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 1rem;
  padding: 0.375rem 0.75rem;
  border-bottom: 1px solid #f1f3f4;
  font-size: 0.875rem;
}

.stop-entry:last-child {
  border-bottom: none;
}

.stop-empty {
  padding: 0.5rem 0.75rem;
  color: #5f6368;
  font-size: 0.875rem;
}

.dialog-error {
  margin: 0.5rem 0 0;
  color: #b3261e;
  font-size: 0.8125rem;
}

.dialog-actions {
  display: flex;
  align-items: center;
  gap: 0.5rem;
  margin-top: 1.25rem;
}

.actions-spacer {
  flex: 1;
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
