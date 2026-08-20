<script setup lang="ts">
import DownloadMenu from "./DownloadMenu.vue";

// The editor's app-bar-actions cluster (SPEC-010/006/060): the autosave state
// indicator, its failure-state "Download a copy" escape hatch, and the
// Download menu. Extracted from DocumentView's template so the view stays a
// wiring hub — it hands in the reactive state and re-emits the download pick.
const props = defineProps({
  /** "saved" | "saving" | "error" — drives the indicator and the escape link. */
  saveState: { type: String, required: true },
  /** The human-readable indicator text ("Saving…", etc.). */
  saveStateText: { type: String, required: true },
  /** The format the failure-state download link saves in. */
  downloadFormat: { type: String, default: "docx" },
});

const emit = defineEmits(["download"]);
</script>

<template>
  <div class="app-bar-cluster">
    <span
      class="save-state"
      data-testid="save-state"
      :data-save-state="props.saveState"
    >{{ props.saveStateText }}</span>
    <a
      v-if="props.saveState === 'error'"
      class="save-state-download"
      data-testid="save-state-download"
      href="#"
      @click.prevent="emit('download', props.downloadFormat)"
    >Download a copy</a>
    <DownloadMenu @download="(format) => emit('download', format)" />
  </div>
</template>

<style scoped>
.app-bar-cluster {
  display: flex;
  align-items: center;
  gap: var(--space-4);
}

.save-state {
  font-size: var(--text-md);
  color: var(--text-on-brand);
  white-space: nowrap;
}

.save-state-download {
  font-size: var(--text-md);
  color: var(--text-on-brand);
}
</style>
