<script setup lang="ts">
// "What changed" (SPEC-111): what the last merge could not carry, as a
// disclosure — one row per PdfDropRecord, grouped by kind with a count per
// group — so nothing a merge drops is lost silently. A merge that dropped
// nothing renders nothing at all. Runtime prop declarations (espree lints
// .vue scripts).
import { computed } from "vue";
import { dropRowText } from "../../lib/pdfMerge";
import { asDropGroups } from "../../views/assemble/usePdfMerge";

const props = defineProps({
  /** The drop records grouped by kind (groupDrops). */
  groups: { type: Array, default: () => [] },
});

const typed = computed(() => asDropGroups(props.groups));
const total = computed(() => typed.value.reduce((sum, group) => sum + group.records.length, 0));
</script>

<template>
  <details
    v-if="total > 0"
    class="merge-drops"
    data-testid="merge-drops"
  >
    <summary>What changed ({{ total }})</summary>
    <section
      v-for="group in typed"
      :key="group.kind"
      class="merge-drops-group"
      data-testid="merge-drop-group"
      :data-kind="group.kind"
    >
      <h3 class="merge-drops-heading">
        {{ group.label }} <span data-testid="merge-drop-count">({{ group.records.length }})</span>
      </h3>
      <ul class="merge-drops-list">
        <li
          v-for="(record, index) in group.records"
          :key="index"
          data-testid="merge-drop"
          :data-kind="record.kind"
        >
          {{ dropRowText(record) }}
        </li>
      </ul>
    </section>
  </details>
</template>

<style scoped>
.merge-drops {
  font-size: var(--text-sm);
  color: var(--text-primary);
  border: var(--border-width) solid var(--border-subtle);
  border-radius: var(--radius-md);
  padding: var(--space-2) var(--space-3);
  background: var(--surface-base);
  max-height: 14rem;
  overflow-y: auto;
}

.merge-drops summary {
  cursor: pointer;
  font-weight: var(--weight-medium);
  min-height: var(--height-control);
  display: flex;
  align-items: center;
}

.merge-drops summary:focus-visible {
  outline: var(--space-1) solid var(--focus-ring);
  outline-offset: 0;
}

.merge-drops-heading {
  margin: var(--space-2) 0 var(--space-1);
  font-size: var(--text-sm);
  font-weight: var(--weight-medium);
}

.merge-drops-list {
  margin: 0;
  padding-left: var(--space-4);
  color: var(--text-secondary);
}
</style>
