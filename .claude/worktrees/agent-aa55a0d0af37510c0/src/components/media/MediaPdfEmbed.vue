<script setup lang="ts">
// The PDF surface (SPEC-107): the BROWSER'S OWN viewer — an <object> over a
// blob: URL (created and revoked by useMediaDocument), with an <embed> and
// then a plain panel offering Download and Open-in-new-tab as the nested
// fallbacks. Rasterising a PDF page ourselves is explicitly out of scope;
// what that costs (no thumbnails, no toolbar control) is stated in the spec
// rather than hidden. Runtime prop declarations — see AppShell.vue.
const props = defineProps({
  /** blob: URL of the verbatim source bytes. */
  src: { type: String, required: true },
  fileName: { type: String, required: true },
});

const emit = defineEmits(["download"]);
</script>

<template>
  <div
    class="pdf-viewer"
    data-testid="media-pdf"
  >
    <object
      class="pdf-object"
      :data="props.src"
      type="application/pdf"
      :aria-label="props.fileName"
    >
      <embed
        class="pdf-object"
        :src="props.src"
        type="application/pdf"
      >
      <div
        class="pdf-fallback"
        data-testid="media-pdf-fallback"
      >
        <p class="pdf-fallback-message">
          Your browser cannot display PDFs here.
        </p>
        <div class="pdf-fallback-actions">
          <button
            type="button"
            data-testid="media-pdf-download"
            @click="emit('download')"
          >
            Download
          </button>
          <a
            :href="props.src"
            target="_blank"
            rel="noopener"
            data-testid="media-pdf-open-tab"
          >Open in new tab</a>
        </div>
      </div>
    </object>
  </div>
</template>

<style scoped>
.pdf-viewer {
  flex: 1 1 auto;
  min-width: 0;
  display: flex;
  background: var(--surface-sunken);
}

.pdf-object {
  flex: 1 1 auto;
  width: 100%;
  height: 100%;
  border: none;
}

.pdf-fallback {
  height: 100%;
  display: flex;
  flex-direction: column;
  align-items: center;
  justify-content: center;
  gap: var(--space-4);
  color: var(--text-primary);
  text-align: center;
  padding: var(--space-5);
}

.pdf-fallback-message {
  margin: 0;
}

.pdf-fallback-actions {
  display: flex;
  align-items: center;
  gap: var(--space-4);
}

.pdf-fallback-actions button {
  font: inherit;
  padding: var(--space-2) var(--space-5);
  border: var(--border-width) solid var(--border-strong);
  border-radius: var(--radius-sm);
  background: var(--surface-base);
  color: var(--text-primary);
  cursor: pointer;
}

.pdf-fallback-actions a {
  color: var(--text-accent);
}
</style>
