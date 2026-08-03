<script setup lang="ts">
import { onUnmounted, ref } from "vue";
import { isStorageAvailable, onStorageStatusChange } from "../lib/documentStore";

// Renders nothing while storage works; once the store reports a failure
// (IndexedDB missing or quota exceeded) the banner appears and stays for the
// rest of the session — the status never flips back.
const available = ref(isStorageAvailable());
const unsubscribe = onStorageStatusChange(() => {
  available.value = isStorageAvailable();
});
onUnmounted(unsubscribe);
</script>

<template>
  <div
    v-if="!available"
    class="storage-warning"
    role="alert"
    data-testid="storage-warning"
  >
    Changes will not be saved in this browser
  </div>
</template>

<style scoped>
.storage-warning {
  flex: none;
  padding: 0.5rem 1rem;
  background: #fef7e0;
  color: #7a5d00;
  border-bottom: 1px solid #f0dfa6;
  font-size: 0.875rem;
}
</style>
