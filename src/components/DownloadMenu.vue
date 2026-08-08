<script setup lang="ts">
import { onMounted, onUnmounted, ref } from "vue";

// Runtime emit declaration (not the defineEmits<{...}>() generic) because
// ESLint parses .vue scripts with espree, which rejects TS-only syntax;
// vue-tsc still infers the emit signature from it. The payload is the
// chosen format id — "docx" for Word (.docx), "odt" for OpenDocument
// Text (.odt), "pdf" for PDF export (SPEC-034).
const emit = defineEmits(["download"]);

const open = ref(false);

function closeMenu() {
  open.value = false;
}

// Any click that bubbles to the document (the component's root stops
// propagation) closes the menu — same pattern as HomeView's row menu.
onMounted(() => document.addEventListener("click", closeMenu));
onUnmounted(() => document.removeEventListener("click", closeMenu));

function choose(format = "") {
  open.value = false;
  emit("download", format);
}
</script>

<template>
  <div
    class="download"
    @click.stop
    @keydown.escape="closeMenu"
  >
    <button
      type="button"
      class="download-button"
      data-testid="download-menu"
      aria-haspopup="menu"
      :aria-expanded="open"
      @click="open = !open"
    >
      Download
    </button>
    <div
      v-if="open"
      class="download-list"
      role="menu"
      aria-label="Download formats"
      data-testid="download-menu-list"
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
        data-testid="download-pdf"
        @click="choose('pdf')"
      >
        PDF (.pdf)
      </button>
    </div>
  </div>
</template>

<style scoped>
.download {
  position: relative;
}

/* Sits on the dark app bar, so the idle state is a quiet outline button. */
.download-button {
  font: inherit;
  font-size: 0.875rem;
  color: inherit;
  background: rgb(255 255 255 / 0.1);
  border: 1px solid rgb(255 255 255 / 0.4);
  border-radius: 4px;
  padding: 0.125rem 0.75rem;
  cursor: pointer;
}

.download-button:hover {
  background: rgb(255 255 255 / 0.2);
}

.download-list {
  position: absolute;
  right: 0;
  top: calc(100% + 0.25rem);
  z-index: 10;
  min-width: 10rem;
  display: flex;
  flex-direction: column;
  background: #fff;
  border: 1px solid #d0d0d0;
  border-radius: 4px;
  box-shadow: 0 2px 6px rgb(0 0 0 / 0.2);
  padding: 0.25rem 0;
}

.download-list button {
  font: inherit;
  font-size: 0.875rem;
  color: #202124;
  text-align: left;
  background: none;
  border: none;
  padding: 0.375rem 1rem;
  cursor: pointer;
  white-space: nowrap;
}

.download-list button:hover {
  background: #f1f3f4;
}
</style>
