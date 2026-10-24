<script setup lang="ts">
// The page organiser's result preview (SPEC-111), in the workspace's centre
// while the tray is a PDF plan: the planned output itself, in the BROWSER'S
// OWN PDF viewer — SPEC-107's <object> over a blob: URL — since page cards
// carry no rendering of their own. It rebuilds a moment after each change
// (at most once every PREVIEW_DEBOUNCE_MS), never by itself past
// PREVIEW_MAX_PAGES pages, and whenever Refresh is pressed; beneath it, "What
// changed" lists what the merge could not carry. All logic is the
// usePdfMerge controller passed in. Runtime prop declarations (espree lints
// .vue scripts).
import { computed } from "vue";
import UiButton from "../ui/UiButton.vue";
import MediaPdfEmbed from "../media/MediaPdfEmbed.vue";
import MergeDrops from "./MergeDrops.vue";
import { PREVIEW_MAX_PAGES } from "../../views/assemble/usePdfMerge";
import { pageCountLabel } from "../../lib/assembly";

const props = defineProps({
  /** The usePdfMerge() controller. */
  merge: { type: Object, required: true },
  /** The output's title, for the viewer's label. */
  title: { type: String, default: "" },
});

const m = props.merge;
const p = m.preview;

const status = computed(() => {
  if (p.busy.value) {
    return "Building the preview…";
  }
  if (p.failure.value) {
    return `The preview could not be built: ${p.failure.value}`;
  }
  if (p.manual.value && p.stale.value) {
    return `This plan has ${pageCountLabel(m.pageCount.value)}; past ${PREVIEW_MAX_PAGES} the preview is built only when you press Refresh.`;
  }
  if (p.stale.value) {
    return p.url.value ? "The plan changed; the preview is updating." : "The preview is on its way.";
  }
  return p.url.value ? "The preview shows the planned PDF." : "";
});
</script>

<template>
  <section
    class="merge-preview-panel"
    aria-label="Result preview"
  >
    <div class="merge-preview-head">
      <h2 class="merge-preview-heading">
        Result preview
      </h2>
      <p
        class="merge-preview-status"
        data-testid="merge-preview-status"
        role="status"
      >
        {{ status }}
      </p>
      <UiButton
        size="sm"
        data-testid="merge-refresh"
        @click="p.refresh()"
      >
        Refresh
      </UiButton>
    </div>
    <div
      class="merge-preview"
      data-testid="merge-preview"
      :data-url="p.url.value || undefined"
    >
      <MediaPdfEmbed
        v-if="p.url.value"
        :src="p.url.value"
        :file-name="`${props.title || 'Combined PDF'}.pdf`"
        @download="m.download()"
      />
      <p
        v-else
        class="merge-preview-empty"
      >
        The planned PDF shows here once it is built. Pages are never rendered in the tray: this is where to check the result.
      </p>
    </div>
    <MergeDrops :groups="m.dropGroups.value" />
  </section>
</template>

<style scoped>
.merge-preview-panel {
  display: flex;
  flex-direction: column;
  gap: var(--space-3);
  min-width: 0;
  min-height: 0;
  font-size: var(--text-sm);
  color: var(--text-primary);
}

.merge-preview-head {
  display: flex;
  align-items: center;
  gap: var(--space-3);
}

.merge-preview-heading {
  margin: 0;
  font-size: var(--text-md);
}

.merge-preview-status {
  margin: 0;
  margin-right: auto;
  color: var(--text-secondary);
}

.merge-preview {
  flex: 1 1 0;
  min-height: 16rem;
  display: flex;
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-md);
  overflow: hidden;
  background: var(--surface-sunken);
}

.merge-preview-empty {
  margin: auto;
  max-width: 28rem;
  padding: var(--space-4);
  text-align: center;
  color: var(--text-secondary);
}
</style>
