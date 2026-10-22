<script setup lang="ts">
// The speaker-notes panel (SPEC-083): a plain multi-line editor bound to the
// current slide's notes, sitting below the stage. Textarea semantics (AC), a
// "Speaker notes for slide N" label, and a "Click to add speaker notes" prompt
// when empty. Editing is per-input (the parent coalesces bursts and marks the
// slide dirty). Runtime prop/emit declarations — espree lints .vue scripts.
import { computed } from "vue";

const props = defineProps({
  /** The current slide's notes as newline-joined plain text. */
  value: { type: String, default: "" },
  /** The 1-based slide index, for the panel/editor label. */
  slideNumber: { type: Number, default: 1 },
});

const emit = defineEmits(["input"]);

const label = computed(() => `Speaker notes for slide ${props.slideNumber}`);

function onInput(event = new Event("input")) {
  const el = event.target;
  if (el instanceof HTMLTextAreaElement) {
    emit("input", el.value);
  }
}
</script>

<template>
  <section
    class="notes-panel"
    data-testid="slides-notes-panel"
    :aria-label="label"
  >
    <textarea
      class="notes-editor"
      data-testid="slides-notes-editor"
      :aria-label="label"
      :value="value"
      placeholder="Click to add speaker notes"
      @input="onInput"
    />
  </section>
</template>

<style scoped>
.notes-panel {
  flex: 0 0 auto;
  display: flex;
  border-top: var(--border-width) solid var(--border-subtle);
  background: var(--surface-default);
}

.notes-editor {
  flex: 1 1 auto;
  /* Seven lines of notes with the panel's own scroll (SPEC-083 Reference). */
  height: calc(7 * 1.45em + 2 * var(--space-3));
  resize: none;
  padding: var(--space-3) var(--space-4);
  border: 0;
  background: transparent;
  color: var(--text-primary);
  font-family: var(--font-ui);
  font-size: var(--text-md);
  line-height: var(--line-normal);
  overflow-y: auto;
}

.notes-editor:focus-visible {
  outline: var(--space-1) solid var(--focus-ring);
  outline-offset: calc(-1 * var(--space-1));
}
</style>
