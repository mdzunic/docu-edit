<script setup lang="ts">
// The recovery bar (SPEC-010): shown when a document reopens over an autosave
// draft newer than its committed record — the previous session never closed
// cleanly. "Keep changes" promotes the recovered draft, "Discard changes"
// returns to the last clean state. Extracted from DocumentView by SPEC-103 so
// the deck shows the same bar over recovered slides; the markup, testids and
// copy ("Recovered unsaved changes from <time>", from useAutosave's
// recoveryLabel) are the text side's verbatim, so no e2e path changed.
//
// Runtime prop/emit declarations — see AppShell.vue's comment on the espree
// constraint.
defineProps({
  /** useAutosave's recoveryLabel: "Recovered unsaved changes from …". */
  label: { type: String, required: true },
});

const emit = defineEmits(["keep", "discard"]);
</script>

<template>
  <div
    class="notice-bar"
    role="status"
    data-testid="recovery-bar"
  >
    <span>{{ label }}</span>
    <button
      type="button"
      class="notice-button"
      data-testid="recovery-keep"
      @click="emit('keep')"
    >
      Keep changes
    </button>
    <button
      type="button"
      class="notice-button"
      data-testid="recovery-discard"
      @click="emit('discard')"
    >
      Discard changes
    </button>
  </div>
</template>

<style scoped>
/* Sticky so the notice stays visible while its editor's canvas scrolls —
   DocumentView's conflict and too-large bars share the look. */
.notice-bar {
  position: sticky;
  top: 0;
  z-index: 10;
  display: flex;
  align-items: center;
  gap: var(--space-4);
  padding: var(--space-3) var(--space-5);
  background: var(--surface-selected);
  border-bottom: var(--border-width) solid var(--accent-subtle);
  color: var(--text-accent);
  font-size: var(--text-md);
}

.notice-button {
  font: inherit;
  font-size: var(--text-md);
  padding: var(--space-1) var(--space-4);
  border: var(--border-width) solid var(--accent-subtle);
  border-radius: var(--radius-sm);
  background: var(--surface-base);
  color: var(--text-accent);
  cursor: pointer;
}
</style>
