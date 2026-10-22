<script setup lang="ts">
// Data ▸ Sort range… (SPEC-128 "UI"): `sort-dialog` on BaseDialog, SPEC-076's
// dialog decisions carried over to the sheet. One `sort-level` row per key —
// "Sort by", then "Then by" — each with its column (`sort-level-column`) and
// order (`sort-level-order`) choices; clicking or focusing a row selects it,
// and `sort-delete-level`, `sort-copy-level`, `sort-move-up` and
// `sort-move-down` act on the selected one; `sort-add-level` appends a level
// (disabled at MAX_SORT_KEYS = 64, stating "Up to 64 levels"). Column choices
// read the header text while "My data has headers" (`sort-has-headers`,
// seeded by the header guess; ticked and disabled over an autofilter's range,
// whose header row is never data, as Excel does) is on — a blank header "Column B" — and
// "Column B" otherwise; orders read "A to Z"/"Z to A", or "Smallest to
// Largest"/"Largest to Smallest" for a column of mostly numbers or dates.
// "Case sensitive" (`sort-case-sensitive`). OK (`sort-ok`) emits confirm({
// keys, hasHeader, caseSensitive }) — a duplicated column is refused here,
// and the view's refusals arrive as `error` — both inline in
// `sort-dialog-error`; Cancel (`sort-cancel`) and Escape emit cancel. Runtime
// prop declarations per the espree constraint (see AppShell.vue).
import { computed, ref } from "vue";
import { duplicateKeyMessage, MAX_SORT_KEYS } from "../../lib/spreadsheet/sort";
import { asSortColumnChoices, sortOrderLabels } from "../../views/spreadsheet/sortFilterLabels";
import AppIcon from "../AppIcon.vue";
import BaseDialog from "../BaseDialog.vue";
import UiButton from "../ui/UiButton.vue";

const props = defineProps({
  /** The column choices with a header row (SortColumnChoice[]). */
  withHeader: { type: Array, default: () => [] },
  /** … and without one. */
  withoutHeader: { type: Array, default: () => [] },
  /** The header guess: the checkbox's initial state. */
  hasHeader: { type: Boolean, default: false },
  /** An autofilter's range: the checkbox ticked and disabled. */
  headerLocked: { type: Boolean, default: false },
  /** The first level's column (absolute). */
  keyCol: { type: Number, default: 0 },
  /** The view's refusal of the last OK ("" for none). */
  error: { type: String, default: "" },
});

const emit = defineEmits(["confirm", "cancel"]);

const MAX_LEVELS_REASON = `Up to ${MAX_SORT_KEYS} levels`;

const headers = ref(props.hasHeader || props.headerLocked);
const caseSensitive = ref(false);
let nextId = 1;
const levels = ref([{ id: 0, col: props.keyCol, descending: false }]);
const selected = ref(0);
const localError = ref("");

const choices = computed(() => asSortColumnChoices(headers.value ? props.withHeader : props.withoutHeader));
const atMax = computed(() => levels.value.length >= MAX_SORT_KEYS);
const errorText = computed(() => localError.value || props.error);

/** A level's order labels, by its column's type. */
function orderLabels(col = 0) {
  const choice = choices.value.find((each) => each.col === col);
  return sortOrderLabels(choice ? choice.numeric : false);
}

function setColumn(index = 0, event = new Event("change")) {
  const level = levels.value[index];
  if (level && event.target instanceof HTMLSelectElement) {
    level.col = Number(event.target.value);
  }
}

function setOrder(index = 0, event = new Event("change")) {
  const level = levels.value[index];
  if (level && event.target instanceof HTMLSelectElement) {
    level.descending = event.target.value === "desc";
  }
}

function addLevel() {
  if (atMax.value) {
    return;
  }
  const used = new Set(levels.value.map((level) => level.col));
  const free = choices.value.find((choice) => !used.has(choice.col)) ?? choices.value[0];
  levels.value.push({ id: nextId++, col: free ? free.col : props.keyCol, descending: false });
  selected.value = levels.value.length - 1;
}

function copyLevel() {
  const level = levels.value[selected.value];
  if (!level || atMax.value) {
    return;
  }
  levels.value.splice(selected.value + 1, 0, { id: nextId++, col: level.col, descending: level.descending });
  selected.value += 1;
}

function deleteLevel() {
  if (levels.value.length <= 1) {
    return;
  }
  levels.value.splice(selected.value, 1);
  selected.value = Math.min(selected.value, levels.value.length - 1);
}

function move(delta = 1) {
  const from = selected.value;
  const to = from + delta;
  if (to < 0 || to >= levels.value.length) {
    return;
  }
  const [level] = levels.value.splice(from, 1);
  if (level) {
    levels.value.splice(to, 0, level);
    selected.value = to;
  }
}

function confirm() {
  const duplicate = levels.value.find((level, index) => levels.value.findIndex((other) => other.col === level.col) !== index);
  if (duplicate) {
    localError.value = duplicateKeyMessage(duplicate.col);
    return;
  }
  localError.value = "";
  emit("confirm", {
    keys: levels.value.map((level) => ({ col: level.col, descending: level.descending })),
    hasHeader: headers.value || props.headerLocked,
    caseSensitive: caseSensitive.value,
  });
}
</script>

<template>
  <BaseDialog
    title="Sort"
    testid="sort-dialog"
    width="34rem"
    @dismiss="emit('cancel')"
  >
    <div class="sort-tools">
      <div
        class="sort-tool-buttons"
        role="toolbar"
        aria-label="Levels"
      >
        <UiButton
          size="sm"
          data-testid="sort-add-level"
          :disabled="atMax"
          :title="atMax ? MAX_LEVELS_REASON : undefined"
          @click="addLevel"
        >
          Add level
        </UiButton>
        <UiButton
          size="sm"
          data-testid="sort-delete-level"
          :disabled="levels.length <= 1"
          @click="deleteLevel"
        >
          Delete level
        </UiButton>
        <UiButton
          size="sm"
          data-testid="sort-copy-level"
          :disabled="atMax"
          :title="atMax ? MAX_LEVELS_REASON : undefined"
          @click="copyLevel"
        >
          Copy level
        </UiButton>
        <UiButton
          size="sm"
          data-testid="sort-move-up"
          aria-label="Move up"
          title="Move up"
          :disabled="selected === 0"
          @click="move(-1)"
        >
          <AppIcon name="chevron-up" />
        </UiButton>
        <UiButton
          size="sm"
          data-testid="sort-move-down"
          aria-label="Move down"
          title="Move down"
          :disabled="selected >= levels.length - 1"
          @click="move(1)"
        >
          <AppIcon name="chevron-down" />
        </UiButton>
        <span
          v-if="atMax"
          class="sort-reason"
          data-testid="sort-add-level-reason"
        >{{ MAX_LEVELS_REASON }}</span>
      </div>
      <div class="sort-options">
        <label class="sort-check">
          <input
            v-model="headers"
            type="checkbox"
            :disabled="props.headerLocked"
            data-testid="sort-has-headers"
          >
          <span>My data has headers</span>
        </label>
        <label class="sort-check">
          <input
            v-model="caseSensitive"
            type="checkbox"
            data-testid="sort-case-sensitive"
          >
          <span>Case sensitive</span>
        </label>
      </div>
    </div>
    <ol
      class="sort-levels"
      aria-label="Sort levels"
    >
      <li
        v-for="(level, index) in levels"
        :key="level.id"
        class="sort-level"
        :class="{ 'is-selected': index === selected }"
        data-testid="sort-level"
        :aria-current="index === selected ? 'true' : undefined"
        @click="selected = index"
        @focusin="selected = index"
      >
        <span class="sort-level-label">{{ index === 0 ? "Sort by" : "Then by" }}</span>
        <select
          class="sort-select"
          data-testid="sort-level-column"
          :aria-label="`${index === 0 ? 'Sort by' : 'Then by'} column`"
          :value="level.col"
          @change="setColumn(index, $event)"
        >
          <option
            v-for="choice in choices"
            :key="choice.col"
            :value="choice.col"
          >
            {{ choice.label }}
          </option>
        </select>
        <select
          class="sort-select"
          data-testid="sort-level-order"
          aria-label="Order"
          :value="level.descending ? 'desc' : 'asc'"
          @change="setOrder(index, $event)"
        >
          <option value="asc">
            {{ orderLabels(level.col)[0] }}
          </option>
          <option value="desc">
            {{ orderLabels(level.col)[1] }}
          </option>
        </select>
      </li>
    </ol>
    <p
      v-if="errorText"
      class="sort-error"
      role="alert"
      data-testid="sort-dialog-error"
    >
      {{ errorText }}
    </p>
    <template #actions>
      <UiButton
        variant="secondary"
        data-testid="sort-cancel"
        @click="emit('cancel')"
      >
        Cancel
      </UiButton>
      <UiButton
        variant="primary"
        data-testid="sort-ok"
        @click="confirm"
      >
        OK
      </UiButton>
    </template>
  </BaseDialog>
</template>

<style scoped>
.sort-tools {
  display: flex;
  flex-wrap: wrap;
  align-items: center;
  justify-content: space-between;
  gap: var(--space-3);
}

.sort-tool-buttons,
.sort-options {
  display: flex;
  flex-wrap: wrap;
  align-items: center;
  gap: var(--space-2);
}

.sort-reason {
  color: var(--text-disabled);
  font-size: var(--text-xs);
  font-style: italic;
}

.sort-check {
  display: inline-flex;
  align-items: center;
  gap: var(--space-1);
  color: var(--text-primary);
  font-size: var(--text-sm);
}

.sort-levels {
  display: flex;
  flex-direction: column;
  gap: var(--space-1);
  max-height: 16rem;
  margin: var(--space-4) 0 0;
  padding: 0;
  overflow-y: auto;
  list-style: none;
}

.sort-level {
  display: grid;
  grid-template-columns: 5rem 1fr 1fr;
  align-items: center;
  gap: var(--space-2);
  padding: var(--space-1) var(--space-2);
  border: var(--border-width) solid transparent;
  border-radius: var(--radius-sm);
}

.sort-level.is-selected {
  background: color-mix(in srgb, var(--accent) 12%, transparent);
  border-color: var(--accent);
}

.sort-level-label {
  color: var(--text-secondary);
  font-size: var(--text-sm);
}

.sort-select {
  height: var(--height-control);
  padding: 0 var(--space-2);
  font: inherit;
  color: var(--text-primary);
  background: var(--surface-raised);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-sm);
}

.sort-error {
  margin: var(--space-3) 0 0;
  color: var(--text-danger);
  font-size: var(--text-sm);
}
</style>
