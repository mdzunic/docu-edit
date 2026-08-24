<script setup lang="ts">
// Storage usage panel (SPEC-067 AC-47..50): bytes used (with a percentage
// when the browser can estimate quota), the largest documents, and the
// trash shortcut when it would free space. The quota read lives here with
// the panel that displays it.
import { computed, onMounted, ref } from "vue";
import { formatFileSize } from "../../lib/format";

const props = defineProps({
  lib: { type: Object, required: true },
});

// Storage quota (AC-47): the browser's estimate, read once. Absent in
// engines without the Storage API, in which case the panel shows the used
// bytes without a percentage.
const quotaBytes = ref(0);
onMounted(async () => {
  try {
    if (navigator.storage?.estimate) {
      const estimate = await navigator.storage.estimate();
      quotaBytes.value = estimate.quota ?? 0;
    }
  } catch {
    // No estimate available — fall back to bytes-only display.
  }
});
const usagePercent = computed(() =>
  quotaBytes.value > 0
    ? Math.min(100, Math.round((props.lib.usage.value.totalBytes / quotaBytes.value) * 100))
    : null,
);
const usageHigh = computed(() => usagePercent.value !== null && usagePercent.value >= 80);
</script>

<template>
  <div
    class="storage-panel"
    data-testid="storage-panel"
  >
    <span
      class="storage-usage"
      data-testid="storage-usage"
    >
      {{ formatFileSize(props.lib.usage.value.totalBytes) }}<template v-if="usagePercent !== null"> of your browser's storage ({{ usagePercent }}%)</template> in use
    </span>
    <span
      v-if="props.lib.largestDocuments.value.length > 0"
      class="storage-largest"
    >
      Largest:
      <span
        v-for="doc in props.lib.largestDocuments.value.slice(0, 3)"
        :key="doc.id"
        class="largest-doc"
        data-testid="largest-doc"
      >{{ doc.title }} ({{ formatFileSize(doc.bytes) }})</span>
    </span>
    <button
      v-if="usageHigh || props.lib.trashCount.value > 0"
      type="button"
      class="storage-trash-link"
      data-testid="storage-trash-link"
      @click="props.lib.setShowTrash(true)"
    >
      {{ usageHigh ? "Storage is running low — empty the trash" : "Review the trash" }}
    </button>
  </div>
</template>

<style scoped>
.storage-panel {
  display: flex;
  align-items: center;
  gap: var(--space-4);
  flex-wrap: wrap;
  padding: var(--space-3) var(--space-4);
  margin-bottom: var(--space-5);
  border-radius: var(--radius-md);
  background: var(--surface-raised);
  color: var(--text-secondary);
  font-size: var(--text-sm);
}

.storage-largest {
  display: flex;
  gap: var(--space-3);
  flex-wrap: wrap;
}

.largest-doc:not(:last-child)::after {
  content: ",";
}

.storage-trash-link {
  font: inherit;
  margin-left: auto;
  border: none;
  background: transparent;
  color: var(--text-accent);
  cursor: pointer;
  text-decoration: underline;
}
</style>
