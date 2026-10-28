<script setup lang="ts">
// SPEC-114 (experimental): Labs' controls for the library search index —
// "Rebuild index" forgets every stored record and builds the index again
// from scratch; "Clear index" deletes every record and stops building. The
// index is a cache of the library's text kept in this browser, so either
// costs time and nothing else. Loaded by LabsSection's ExperimentGate only
// while the library-search experiment is on.
import { computed, ref } from "vue";
import ExperimentBadge from "../../components/ExperimentBadge.vue";
import UiButton from "../../components/ui/UiButton.vue";
import { useLibrarySearch } from "../../composables/useLibrarySearch";

const search = useLibrarySearch();
const message = ref("");
const busy = ref(false);

const progressLine = computed(() => {
  const progress = search.progress.value;
  return progress.running ? `Indexed ${progress.done} of ${progress.total}` : "";
});

async function rebuild() {
  busy.value = true;
  message.value = "Rebuilding the index…";
  // Building runs in the background; the line above counts it.
  void search.rebuild().finally(() => {
    busy.value = false;
    message.value = "The index was rebuilt.";
  });
}

async function clear() {
  busy.value = true;
  await search.clearIndex();
  busy.value = false;
  message.value = "The index was cleared. It is built again the next time you search Everything.";
}
</script>

<template>
  <div
    class="search-labs"
    data-testid="prefs-library-search"
  >
    <p class="search-labs-copy">
      <ExperimentBadge experiment="library-search" />
      Library search keeps an index of your documents' text in this browser. It is a cache:
      rebuilding or clearing it loses nothing but the time to build it again.
    </p>
    <div class="search-labs-actions">
      <UiButton
        data-testid="search-rebuild-index"
        :disabled="busy"
        @click="rebuild"
      >
        Rebuild index
      </UiButton>
      <UiButton
        data-testid="search-clear-index"
        :disabled="busy"
        @click="clear"
      >
        Clear index
      </UiButton>
    </div>
    <p
      class="search-labs-status"
      role="status"
      data-testid="search-labs-status"
    >
      {{ progressLine || message }}
    </p>
  </div>
</template>

<style scoped>
.search-labs {
  display: flex;
  flex-direction: column;
  gap: var(--space-2);
}

.search-labs-copy,
.search-labs-status {
  margin: 0;
  color: var(--text-secondary);
  font-size: var(--text-sm);
  line-height: var(--line-normal);
}

.search-labs-actions {
  display: flex;
  gap: var(--space-3);
}
</style>
