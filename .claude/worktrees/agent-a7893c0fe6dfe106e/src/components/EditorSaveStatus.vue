<script setup lang="ts">
import SaveStateBadge from "./SaveStateBadge.vue";
import UiIconButton from "./ui/UiIconButton.vue";

// The editor's app-bar-actions cluster (SPEC-010/006/060): the autosave state
// indicator, its failure-state "Download a copy" escape hatch, and — at the
// bar's far right — the Close control that returns to the library. Extracted
// from DocumentView's template so the view stays a wiring hub — it hands in
// the reactive state and re-emits the download pick and the close request.
// Exporting a copy in a chosen format lives in File → Export (DocumentMenuBar),
// not here.
const props = defineProps({
  /** "saved" | "saving" | "error" — drives the indicator and the escape link. */
  saveState: { type: String, required: true },
  /** The human-readable indicator text ("Saving…", etc.). */
  saveStateText: { type: String, required: true },
  /** The format the failure-state download link saves in. */
  downloadFormat: { type: String, default: "docx" },
  /** True while the close flush is in flight (SPEC-060): the control dims,
   * reads aria-busy and blocks a second activation. */
  closing: { type: Boolean, default: false },
});

const emit = defineEmits(["download", "close"]);
</script>

<template>
  <div class="app-bar-cluster">
    <SaveStateBadge
      :save-state="props.saveState"
      :text="props.saveStateText"
    />
    <a
      v-if="props.saveState === 'error'"
      class="save-state-download"
      data-testid="save-state-download"
      href="#"
      @click.prevent="emit('download', props.downloadFormat)"
    >Download a copy</a>
    <UiIconButton
      class="close-button"
      icon="close"
      label="Close document"
      data-testid="close-document"
      :busy="props.closing"
      @click="emit('close')"
    />
  </div>
</template>

<style scoped>
.app-bar-cluster {
  display: flex;
  align-items: center;
  gap: var(--space-4);
}

.save-state-download {
  font-size: var(--text-md);
  color: var(--text-on-brand);
}

/* The icon button's default greys are tuned for light surfaces; on the brand
   bar it takes the on-brand foreground and a translucent wash so it reads in
   both themes. Matches UiIconButton's own hover/active specificity. */
.app-bar-cluster .close-button {
  color: var(--text-on-brand);
}

.app-bar-cluster .close-button:not(:disabled):hover {
  color: var(--text-on-brand);
  background: color-mix(in srgb, var(--text-on-brand) 12%, transparent);
}

.app-bar-cluster .close-button:not(:disabled):active {
  background: color-mix(in srgb, var(--text-on-brand) 20%, transparent);
}
</style>
