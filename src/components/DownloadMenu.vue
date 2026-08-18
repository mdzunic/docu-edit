<script setup lang="ts">
import { ref, useTemplateRef } from "vue";
import UiMenuSurface from "./ui/UiMenuSurface.vue";

// Runtime emit declaration (not the defineEmits<{...}>() generic) because
// ESLint parses .vue scripts with espree, which rejects TS-only syntax;
// vue-tsc still infers the emit signature from it. The payload is the
// chosen format id — "docx" for Word (.docx), "odt" for OpenDocument
// Text (.odt), "md" for Markdown (SPEC-050), "pdf" for PDF export
// (SPEC-034).
const emit = defineEmits(["download"]);

const open = ref(false);
const triggerEl = useTemplateRef("triggerEl");

// Outside-click and Escape dismissal (and focus return to the trigger) are
// owned by UiMenuSurface — the shared, tested implementation.
function choose(format = "") {
  open.value = false;
  emit("download", format);
}
</script>

<template>
  <div class="download">
    <button
      ref="triggerEl"
      type="button"
      class="download-button"
      data-testid="download-menu"
      aria-haspopup="menu"
      :aria-expanded="open"
      @click="open = !open"
    >
      Download
    </button>
    <UiMenuSurface
      :open="open"
      :anchor="triggerEl"
      class="download-list"
      role="menu"
      aria-label="Download formats"
      data-testid="download-menu-list"
      @dismiss="open = false"
    >
      <button
        type="button"
        role="menuitem"
        data-testid="download-docx"
        @click="choose('docx')"
      >
        Word (.docx)
      </button>
      <button
        type="button"
        role="menuitem"
        data-testid="download-odt"
        @click="choose('odt')"
      >
        OpenDocument Text (.odt)
      </button>
      <button
        type="button"
        role="menuitem"
        data-testid="download-md"
        @click="choose('md')"
      >
        Markdown (.md)
      </button>
      <button
        type="button"
        role="menuitem"
        data-testid="download-pdf"
        @click="choose('pdf')"
      >
        PDF (.pdf)
      </button>
    </UiMenuSurface>
  </div>
</template>

<style scoped>
.download {
  position: relative;
}

/* Sits on the dark app bar, so the idle state is a quiet outline button. */
.download-button {
  font: inherit;
  font-size: var(--text-md);
  color: inherit;
  background: transparent;
  border: var(--border-width) solid var(--text-on-brand);
  border-radius: var(--radius-sm);
  padding: var(--space-1) var(--space-4);
  cursor: pointer;
}

.download-button:hover {
  background: var(--state-hover);
}

.download-list {
  position: absolute;
  right: 0;
  top: calc(100% + var(--space-2));
  z-index: 30;
  min-width: 10rem;
  display: flex;
  flex-direction: column;
}

.download-list button {
  font: inherit;
  font-size: var(--text-md);
  color: var(--text-primary);
  text-align: left;
  background: none;
  border: none;
  border-radius: var(--radius-sm);
  padding: var(--space-2) var(--space-4);
  cursor: pointer;
  white-space: nowrap;
}

.download-list button:hover {
  background: var(--state-hover);
}
</style>
