<script setup lang="ts">
// Formulas ▸ Name Manager… (SPEC-133): `name-manager` on BaseDialog, bound to
// useNames. A grid of the visible names of every scope — one
// `name-manager-row` each (`-row-name`, `-row-value`, `-row-refers-to`,
// `-row-scope`; a click selects it, aria-selected) — and New…, Edit… and
// Delete (`name-manager-new`, `-edit`, `-delete`; Edit and Delete act on the
// selected row and are disabled without one) and Close (`name-manager-close`,
// Escape too). Runtime prop declarations per the espree constraint (see
// AppShell.vue).
import { computed } from "vue";
import { asNames } from "../../views/spreadsheet/useNames";
import BaseDialog from "../BaseDialog.vue";
import UiButton from "../ui/UiButton.vue";

const props = defineProps({
  /** useNames' return. */
  names: { type: Object, required: true },
});

const n = computed(() => asNames(props.names));
const hasSelection = computed(() => n.value.rows.value.some((row) => row.key === n.value.selectedKey.value));

/** A double-click edits the row. */
function editRow(key = "") {
  n.value.select(key);
  n.value.openEdit();
}
</script>

<template>
  <BaseDialog
    title="Name Manager"
    testid="name-manager"
    width="40rem"
    @dismiss="n.closeManager()"
  >
    <div class="name-manager-tools">
      <UiButton
        variant="secondary"
        data-testid="name-manager-new"
        autofocus
        @click="n.openNew()"
      >
        New…
      </UiButton>
      <UiButton
        variant="secondary"
        data-testid="name-manager-edit"
        :disabled="!hasSelection"
        @click="n.openEdit()"
      >
        Edit…
      </UiButton>
      <UiButton
        variant="secondary"
        data-testid="name-manager-delete"
        :disabled="!hasSelection"
        @click="n.deleteSelected()"
      >
        Delete
      </UiButton>
    </div>
    <div
      class="name-manager-table"
      role="grid"
      aria-label="Names"
      aria-readonly="true"
    >
      <div
        class="name-manager-head"
        role="row"
      >
        <span role="columnheader">Name</span>
        <span role="columnheader">Value</span>
        <span role="columnheader">Refers to</span>
        <span role="columnheader">Scope</span>
      </div>
      <div
        v-for="row in n.rows.value"
        :key="row.key"
        class="name-manager-row"
        role="row"
        tabindex="0"
        data-testid="name-manager-row"
        :aria-selected="row.key === n.selectedKey.value ? 'true' : 'false'"
        @click="n.select(row.key)"
        @keydown.enter.prevent="n.select(row.key)"
        @keydown.space.prevent="n.select(row.key)"
        @dblclick="editRow(row.key)"
      >
        <span
          role="gridcell"
          data-testid="name-manager-row-name"
        >{{ row.name }}</span>
        <span
          role="gridcell"
          data-testid="name-manager-row-value"
        >{{ row.value }}</span>
        <span
          role="gridcell"
          data-testid="name-manager-row-refers-to"
        >{{ row.refersTo }}</span>
        <span
          role="gridcell"
          data-testid="name-manager-row-scope"
        >{{ row.scope }}</span>
      </div>
      <p
        v-if="n.rows.value.length === 0"
        class="name-manager-empty"
      >
        No names are defined yet.
      </p>
    </div>
    <template #actions>
      <UiButton
        variant="primary"
        data-testid="name-manager-close"
        @click="n.closeManager()"
      >
        Close
      </UiButton>
    </template>
  </BaseDialog>
</template>

<style scoped>
.name-manager-tools {
  display: flex;
  gap: var(--space-2);
  margin-bottom: var(--space-3);
}

.name-manager-table {
  max-height: 20rem;
  overflow: auto;
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-sm);
  color: var(--text-primary);
  font-size: var(--text-sm);
}

.name-manager-head,
.name-manager-row {
  display: grid;
  grid-template-columns: 1fr 1fr 1.5fr 0.8fr;
  gap: var(--space-2);
  padding: var(--space-1) var(--space-2);
}

.name-manager-head {
  position: sticky;
  top: 0;
  background: var(--surface-raised);
  color: var(--text-secondary);
  font-weight: var(--weight-button);
}

.name-manager-row {
  cursor: pointer;
}

.name-manager-row > span {
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.name-manager-row:hover {
  background: var(--surface-sunken);
}

.name-manager-row[aria-selected="true"] {
  background: var(--surface-selected, var(--surface-sunken));
  outline: var(--border-width) solid var(--focus-ring);
  outline-offset: -1px;
}

.name-manager-row:focus-visible {
  outline: var(--space-1) solid var(--focus-ring);
  outline-offset: -2px;
}

.name-manager-empty {
  margin: 0;
  padding: var(--space-3);
  color: var(--text-secondary);
}
</style>
