<script setup lang="ts">
// SPEC-115: the plan step's "What we changed" list — every substitution and
// cut the validators made to the model's reply (a closed-vocabulary value
// replaced by its nearest default, a slide or bullet cut, a redesign entry
// discarded or appended), and, live, what Build would not carry from the plan
// as edited: a title slide's bullets past its subtitle above all. Shown before
// anything is built, so nothing is dropped silently.
// Runtime prop declarations per the espree constraint (see AppShell.vue).
import { dropLine } from "./designerFields";

defineProps({
  /** DeckDropRecords from the reply's validation. */
  records: { type: Array, default: () => [] },
  /** DeckDropRecords the edited plan would incur at Build. */
  pending: { type: Array, default: () => [] },
});
</script>

<template>
  <section
    v-if="records.length > 0 || pending.length > 0"
    class="drops"
    data-testid="designer-drops"
    aria-labelledby="designer-drops-title"
  >
    <h3
      id="designer-drops-title"
      class="drops-title"
    >
      What we changed
    </h3>
    <ul class="drops-list">
      <li
        v-for="(record, index) in records"
        :key="`r${index}`"
        data-testid="designer-drop"
      >
        <strong>{{ dropLine(record).at }}:</strong> {{ dropLine(record).why }}
      </li>
      <li
        v-for="(record, index) in pending"
        :key="`p${index}`"
        data-testid="designer-pending-drop"
      >
        <strong>{{ dropLine(record).at }} — will not be built:</strong> {{ dropLine(record).why }}
      </li>
    </ul>
  </section>
</template>

<style scoped>
.drops {
  padding: var(--space-3);
  border: var(--border-width) solid var(--warning);
  border-radius: var(--radius-sm);
  background: var(--surface-sunken);
}

.drops-title {
  margin: 0 0 var(--space-2);
  color: var(--text-primary);
  font-size: var(--text-sm);
  font-weight: var(--weight-bold);
}

.drops-list {
  margin: 0;
  padding-left: var(--space-5);
  color: var(--text-primary);
  font-size: var(--text-sm);
  line-height: var(--line-normal);
}
</style>
