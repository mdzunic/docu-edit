<script setup lang="ts">
// The PDF surface (SPEC-107): the BROWSER'S OWN viewer — an <object> over a
// blob: URL (created and revoked by useMediaDocument), with an <embed> and
// then a plain panel offering Download and Open-in-new-tab as the nested
// fallbacks. Rasterising a PDF page ourselves is explicitly out of scope;
// what that costs (no thumbnails, no toolbar control) is stated in the spec
// rather than hidden. Runtime prop declarations — see AppShell.vue.
//
// SPEC-114: opened from a library search result (`find`), the query shows
// above the viewer with a Copy button, and the viewer's URL carries
// `#search=<query>` — a fragment only some browsers' viewers honour, which
// the bar says plainly rather than implying the jump will happen.
import { computed, ref } from "vue";

const props = defineProps({
  /** blob: URL of the verbatim source bytes. */
  src: { type: String, required: true },
  fileName: { type: String, required: true },
  /** A search's text to find in the viewer; "" for none. */
  find: { type: String, default: "" },
});

const emit = defineEmits(["download"]);

const viewerSrc = computed(() => (props.find ? `${props.src}#search=${encodeURIComponent(props.find)}` : props.src));
const copied = ref("");

function selectQuery(event = new FocusEvent("focus")) {
  if (event.target instanceof HTMLInputElement) {
    event.target.select();
  }
}

async function copyFind() {
  try {
    await navigator.clipboard.writeText(props.find);
    copied.value = "Copied";
  } catch {
    copied.value = "Copy it from the box";
  }
}
</script>

<template>
  <div
    class="pdf-viewer"
    data-testid="media-pdf"
  >
    <div
      v-if="props.find"
      class="pdf-find"
      data-testid="media-find"
    >
      <label class="pdf-find-label">
        Find in this PDF
        <input
          class="pdf-find-query"
          data-testid="media-find-query"
          :value="props.find"
          readonly
          @focus="selectQuery"
        >
      </label>
      <button
        type="button"
        data-testid="media-find-copy"
        @click="copyFind"
      >
        Copy
      </button>
      <span
        class="pdf-find-copied"
        role="status"
      >{{ copied }}</span>
      <p
        class="pdf-find-note"
        data-testid="media-find-note"
      >
        The viewer below was asked to jump to the first match, but only some browsers' PDF
        viewers honour that request. If yours did not, use its own find (Ctrl+F or ⌘F) with
        the copied text.
      </p>
    </div>
    <object
      class="pdf-object"
      :data="viewerSrc"
      type="application/pdf"
      :aria-label="props.fileName"
    >
      <embed
        class="pdf-object"
        :src="viewerSrc"
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
  flex-direction: column;
  background: var(--surface-sunken);
}

.pdf-find {
  display: flex;
  flex-wrap: wrap;
  align-items: center;
  gap: var(--space-2) var(--space-3);
  padding: var(--space-2) var(--space-4);
  border-bottom: var(--border-width) solid var(--border-subtle);
  background: var(--surface-base);
  color: var(--text-primary);
  font-size: var(--text-sm);
}

.pdf-find-label {
  display: inline-flex;
  align-items: center;
  gap: var(--space-2);
}

.pdf-find-query {
  font: inherit;
  min-height: var(--height-control);
  padding: 0 var(--space-2);
  border: var(--border-width) solid var(--border-strong);
  border-radius: var(--radius-sm);
  background: var(--surface-base);
  color: var(--text-primary);
}

.pdf-find button {
  font: inherit;
  min-height: var(--height-control);
  padding: 0 var(--space-3);
  border: var(--border-width) solid var(--border-strong);
  border-radius: var(--radius-sm);
  background: var(--surface-base);
  color: var(--text-primary);
  cursor: pointer;
}

.pdf-find-copied {
  color: var(--text-secondary);
}

.pdf-find-note {
  flex: 1 1 100%;
  margin: 0;
  color: var(--text-secondary);
}

.pdf-object {
  flex: 1 1 auto;
  width: 100%;
  height: 100%;
  min-height: 0;
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
