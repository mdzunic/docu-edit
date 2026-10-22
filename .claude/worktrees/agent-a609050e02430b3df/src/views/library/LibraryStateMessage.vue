// The library's centred state messages (SPEC-067): the "no results" branch
// (AC-54..56) with a clear-filters action, and the "library empty but trash
// has documents" note. Product-neutral — HomeView picks which branch via the
// `noResults` prop and wires the clear-filters intent.
<script setup lang="ts">
import UiButton from "../../components/ui/UiButton.vue";

const props = defineProps({
  noResults: { type: Boolean, default: false },
  trashCount: { type: Number, default: 0 },
});

const emit = defineEmits(["clear-filters"]);
</script>

<template>
  <div
    v-if="props.noResults"
    class="state-message no-results"
    data-testid="no-results"
  >
    <p>No documents match your search or filters.</p>
    <UiButton
      data-testid="clear-filters"
      @click="emit('clear-filters')"
    >
      Clear search and filters
    </UiButton>
  </div>
  <div
    v-else
    class="state-message"
    data-testid="library-empty"
  >
    <p>Your library is empty. {{ props.trashCount }} document(s) in the trash.</p>
  </div>
</template>

<style scoped>
.state-message {
  padding: var(--space-7) var(--space-5);
  text-align: center;
  color: var(--text-secondary);
  display: flex;
  flex-direction: column;
  align-items: center;
  gap: var(--space-4);
}

.no-results {
  border: var(--border-width) dashed var(--border-default);
  border-radius: var(--radius-md);
}
</style>
