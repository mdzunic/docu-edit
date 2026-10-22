<script setup lang="ts">
// The image editor's history list (SPEC-112, `media-edit-history`): one row
// per edit in the order the user made it — a record of intent; the stack it
// normalises to is what renders — each with its values and a remove button.
// Runtime prop declarations — see AppShell.vue.
import { computed } from "vue";
import { describeEdit, describeEditDetail } from "../../lib/core/imageEdits";
import { asImageEditor } from "../../views/media/useImageEditor";

const props = defineProps({
  /** The useImageEditor controller. */
  editor: { type: Object, required: true },
});

const e = computed(() => asImageEditor(props.editor));
const rows = computed(() =>
  e.value.history.value.map((edit, index) => ({ index, label: describeEdit(edit), detail: describeEditDetail(edit) })),
);
</script>

<template>
  <aside
    class="edit-history"
    aria-labelledby="edit-history-heading"
  >
    <h2
      id="edit-history-heading"
      class="heading"
    >
      Edits
    </h2>
    <ol
      class="rows"
      data-testid="media-edit-history"
      aria-label="Applied edits"
    >
      <li
        v-for="row in rows"
        :key="`${row.index}-${row.label}`"
        class="row"
        data-testid="edit-history-row"
      >
        <span class="label">{{ row.label }}</span>
        <span
          v-if="row.detail"
          class="detail"
        >{{ row.detail }}</span>
        <button
          type="button"
          class="remove"
          data-testid="edit-history-remove"
          :aria-label="`Remove ${row.label}`"
          :title="`Remove ${row.label}`"
          @click="e.removeEntry(row.index)"
        >
          Remove
        </button>
      </li>
    </ol>
    <p
      v-if="rows.length === 0"
      class="empty"
      data-testid="edit-history-empty"
    >
      No edits yet — the image is the original.
    </p>
  </aside>
</template>

<style scoped>
.edit-history {
  flex: none;
  width: 14rem;
  padding: var(--space-3);
  border-left: var(--border-width) solid var(--border-subtle);
  background: var(--surface-base);
  color: var(--text-primary);
  overflow-y: auto;
}

.heading {
  margin: 0 0 var(--space-3);
  font-size: var(--text-md);
  font-weight: var(--weight-bold);
}

.rows {
  margin: 0;
  padding: 0;
  list-style: none;
  display: flex;
  flex-direction: column;
  gap: var(--space-2);
}

.row {
  display: grid;
  grid-template-columns: 1fr auto;
  gap: 0 var(--space-2);
  align-items: center;
  font-size: var(--text-sm);
}

.detail {
  grid-column: 1;
  color: var(--text-secondary);
  font-size: var(--text-xs);
  font-variant-numeric: tabular-nums;
}

.remove {
  grid-column: 2;
  grid-row: 1;
  font: inherit;
  font-size: var(--text-xs);
  padding: 0 var(--space-2);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-sm);
  background: var(--surface-base);
  color: var(--text-primary);
  cursor: pointer;
}

.empty {
  margin: 0;
  color: var(--text-secondary);
  font-size: var(--text-sm);
}
</style>
