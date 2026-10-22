<script setup lang="ts">
// The deck's Download split-button (SPEC-080): a primary "Download" (own
// format) plus a ▾ menu offering "Download a copy as…" the other format,
// SPEC-088's HTML export and SPEC-104's "PDF (.pdf)…" / "Images (.png)…"
// (both open the print-and-export dialog, on its matching tab).
// Extracted from PresentationView so the view stays a wiring hub
// (PresentationView.size.spec.ts). Runtime prop/emit declarations per the
// espree constraint (see AppShell.vue).
import { ref } from "vue";
import UiButton from "../ui/UiButton.vue";

defineProps({
  /** The label of the other format, for the "Download a copy as…" item. */
  otherFormatLabel: { type: String, default: "" },
});

const emit = defineEmits(["download", "download-copy", "export-html", "export-pdf", "export-images"]);
const menuOpen = ref(false);

function downloadCopy() {
  menuOpen.value = false;
  emit("download-copy");
}

function exportHtml() {
  menuOpen.value = false;
  emit("export-html");
}

function exportPdf() {
  menuOpen.value = false;
  emit("export-pdf");
}

function exportImages() {
  menuOpen.value = false;
  emit("export-images");
}
</script>

<template>
  <div class="download-split">
    <UiButton
      data-testid="presentation-download"
      @click="emit('download')"
    >
      Download
    </UiButton>
    <UiButton
      variant="ghost"
      class="download-caret"
      data-testid="presentation-download-menu"
      aria-haspopup="menu"
      :aria-expanded="menuOpen"
      aria-label="More download options"
      @click="menuOpen = !menuOpen"
    >
      ▾
    </UiButton>
    <div
      v-if="menuOpen"
      class="download-menu"
      role="menu"
      data-testid="presentation-download-menu-popup"
    >
      <button
        type="button"
        role="menuitem"
        class="download-menu-item"
        data-testid="download-copy-as"
        @click="downloadCopy"
      >
        Download a copy as {{ otherFormatLabel }}…
      </button>
      <button
        type="button"
        role="menuitem"
        class="download-menu-item"
        data-testid="presentation-export-html"
        @click="exportHtml"
      >
        Export to HTML (.html)…
      </button>
      <button
        type="button"
        role="menuitem"
        class="download-menu-item"
        data-testid="presentation-export-pdf"
        @click="exportPdf"
      >
        PDF (.pdf)…
      </button>
      <button
        type="button"
        role="menuitem"
        class="download-menu-item"
        data-testid="presentation-export-images"
        @click="exportImages"
      >
        Images (.png)…
      </button>
    </div>
    <div
      v-if="menuOpen"
      class="menu-backdrop"
      @click="menuOpen = false"
    />
  </div>
</template>

<style scoped>
.download-split {
  position: relative;
  display: flex;
  gap: var(--space-1);
  align-items: center;
}

.download-caret {
  padding-inline: var(--space-2);
}

.download-menu {
  position: absolute;
  right: 0;
  top: calc(100% + var(--space-1));
  z-index: 40;
  min-width: 16rem;
  padding: var(--space-1);
  background: var(--surface-raised);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-md);
  box-shadow: var(--shadow-2);
}

.download-menu-item {
  display: block;
  width: 100%;
  text-align: left;
  padding: var(--space-2) var(--space-3);
  border: 0;
  border-radius: var(--radius-sm);
  background: transparent;
  color: var(--text-primary);
  cursor: pointer;
}

.download-menu-item:hover {
  background: var(--surface-sunken);
}

.download-menu-item:focus-visible {
  background: var(--surface-sunken);
  outline: var(--border-width) solid var(--focus-ring);
  outline-offset: calc(-1 * var(--border-width));
}

.menu-backdrop {
  position: fixed;
  inset: 0;
  z-index: 39;
}
</style>
