<script setup lang="ts">
// The workbook's find dock (SPEC-134): the top-right of the sheet area
// holding the find bar — SPEC-017's FindBar, passed in the default slot by
// the view — and, below it, Find all's results (`sheet-find-results`): one
// `sheet-find-result` row per matching cell in search order (`data-sheet` its
// sheet's name, `data-ref` its address) showing Sheet, Cell, Value and
// Formula, capped at FIND_ALL_LIMIT with the cap message. Activating a row (a
// click, Enter or Space) selects that cell exactly as find-next does. The
// dock stacks the bar and the list, so the bar sits in the flow here rather
// than at its own absolute place. Runtime prop declarations per the espree
// constraint (see AppShell.vue).
import { computed } from "vue";
import { asFindResults } from "../../views/spreadsheet/sheetFindRows";
import { asSheetFind } from "../../views/spreadsheet/useSheetFind";

const props = defineProps({
  /** useSheetFind's return. */
  find: { type: Object, required: true },
});

const f = computed(() => asSheetFind(props.find));
const results = computed(() => asFindResults(f.value.results.value));

function onRowKey(index = 0, event = new KeyboardEvent("keydown")) {
  if (event.key === "Enter" || event.key === " ") {
    event.preventDefault();
    f.value.activate(index);
  }
}
</script>

<template>
  <div
    v-if="f.open.value"
    class="sheet-find-dock"
  >
    <div class="sheet-find-stack">
      <slot />
      <section
        v-if="results"
        class="sheet-find-results"
        data-testid="sheet-find-results"
        aria-label="Find all results"
      >
        <table class="sheet-find-table">
          <thead>
            <tr>
              <th scope="col">
                Sheet
              </th>
              <th scope="col">
                Cell
              </th>
              <th scope="col">
                Value
              </th>
              <th scope="col">
                Formula
              </th>
            </tr>
          </thead>
          <tbody>
            <tr
              v-for="row in results.rows"
              :key="row.index"
              class="sheet-find-result"
              :class="{ 'is-current': row.index === f.currentIndex.value }"
              data-testid="sheet-find-result"
              :data-sheet="row.sheetName"
              :data-ref="row.ref"
              tabindex="0"
              @click="f.activate(row.index)"
              @keydown="onRowKey(row.index, $event)"
            >
              <td>{{ row.sheetName }}</td>
              <td>{{ row.ref }}</td>
              <td>{{ row.value }}</td>
              <td>{{ row.formula }}</td>
            </tr>
          </tbody>
        </table>
        <p
          v-if="results.rows.length === 0"
          class="sheet-find-note"
        >
          No results
        </p>
        <p
          v-if="results.cap"
          class="sheet-find-note"
          data-testid="sheet-find-results-cap"
        >
          {{ results.cap }}
        </p>
      </section>
    </div>
  </div>
</template>

<style scoped>
/* A zero-height strip at the top of the sheet area: the stack hangs from its
   right end over the grid, like the text editor's find dock. */
.sheet-find-dock {
  position: relative;
  z-index: 30;
  height: 0;
}

.sheet-find-stack {
  position: absolute;
  top: 0.5rem;
  right: 1.5rem;
  display: flex;
  flex-direction: column;
  align-items: stretch;
  gap: var(--space-2);
  max-width: calc(100vw - 3rem);
}

/* FindBar places itself absolutely; in the dock it stacks. */
.sheet-find-stack :deep(.find-bar) {
  position: static;
}

.sheet-find-results {
  max-height: 16rem;
  overflow: auto;
  background: var(--surface-base);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-md);
  box-shadow: var(--shadow-2);
  font-size: var(--text-sm);
}

.sheet-find-table {
  width: 100%;
  border-collapse: collapse;
}

.sheet-find-table th {
  position: sticky;
  top: 0;
  padding: var(--space-1) var(--space-3);
  background: var(--surface-sunken);
  color: var(--text-secondary);
  font-weight: 600;
  text-align: left;
}

.sheet-find-result td {
  padding: var(--space-1) var(--space-3);
  border-top: var(--border-width) solid var(--border-default);
  color: var(--text-primary);
  white-space: nowrap;
}

.sheet-find-result {
  cursor: pointer;
}

.sheet-find-result:hover {
  background: var(--state-hover);
}

.sheet-find-result.is-current {
  background: var(--surface-selected);
}

.sheet-find-result:focus-visible {
  outline: var(--space-1) solid var(--focus-ring);
  outline-offset: calc(-1 * var(--space-1));
}

.sheet-find-note {
  margin: 0;
  padding: var(--space-2) var(--space-3);
  color: var(--text-secondary);
}
</style>
