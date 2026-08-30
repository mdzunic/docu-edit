<script setup lang="ts">
// The media viewer at /media/:id (SPEC-107): a wiring hub under the same
// rule as DocumentView and HomeView (MediaView.size.spec.ts holds the line
// budget). Load/kind-guard/blob-URL logic lives in useMediaDocument,
// fit/zoom/pan state in useImageViewport, and the surfaces in the
// components/media/* components — this file only wires them together under
// SPEC-058's app bar with SPEC-060's Close. There is deliberately NO editing
// surface here: no autosave session, no version history, no comments, no
// editor File menu — a media document's bytes are immutable.
import { computed } from "vue";
import { useRoute, useRouter } from "vue-router";
import AppShell from "../components/AppShell.vue";
import UiIconButton from "../components/ui/UiIconButton.vue";
import MediaImageCanvas from "../components/media/MediaImageCanvas.vue";
import MediaPdfEmbed from "../components/media/MediaPdfEmbed.vue";
import MediaDetailsPanel from "../components/media/MediaDetailsPanel.vue";
import { useTheme } from "../composables/useTheme";
import { formatBadgeLabel } from "../lib/format";
import { downloadMedia, mediaFileName } from "../lib/media/download";
import { useMediaDocument } from "./media/useMediaDocument";
import { useImageViewport } from "./media/useImageViewport";

const route = useRoute();
const router = useRouter();

// No menu bar here — the shared theme composable keeps the workspace themed
// by the app tokens and following the OS live (SPEC-059).
useTheme();

const mediaId = computed(() => String(route.params.id ?? ""));
const { media, blobUrl, storedAt, close } = useMediaDocument({ mediaId, router });

const title = computed(() => media.value?.title ?? "No document");
const badge = computed(() => (media.value ? formatBadgeLabel(media.value.format) : ""));
const fileName = computed(() => (media.value ? mediaFileName(media.value) : ""));
const isPdf = computed(() => media.value?.format === "pdf");

const viewport = useImageViewport();

function download() {
  if (media.value) {
    downloadMedia(media.value);
  }
}
</script>

<template>
  <AppShell
    :document-title="title"
    :format-badge="badge"
  >
    <template #app-bar-actions>
      <UiIconButton
        class="media-close"
        icon="close"
        label="Close document"
        data-testid="close-document"
        @click="close"
      />
    </template>
    <div
      v-if="media"
      class="media-workspace"
      data-testid="media-workspace"
    >
      <MediaPdfEmbed
        v-if="isPdf"
        :src="blobUrl"
        :file-name="fileName"
        @download="download"
      />
      <MediaImageCanvas
        v-else
        :viewport="viewport"
        :src="blobUrl"
        :alt="fileName"
      />
      <MediaDetailsPanel
        :media="media"
        :stored-at="storedAt"
      />
    </div>
    <div
      v-else
      class="media-missing"
      data-testid="media-not-found"
      role="alert"
    >
      <p>This media file doesn't exist in your library.</p>
    </div>
  </AppShell>
</template>

<style scoped>
.media-workspace {
  height: 100%;
  min-height: 0;
  display: flex;
}

.media-missing {
  height: 100%;
  display: flex;
  align-items: center;
  justify-content: center;
  color: var(--text-primary);
}

/* The icon button's default greys are tuned for light surfaces; on the brand
   bar it takes the on-brand foreground and a translucent wash (the same
   treatment the editor's Close control gets). */
.media-close {
  color: var(--text-on-brand);
}

.media-close:not(:disabled):hover {
  color: var(--text-on-brand);
  background: color-mix(in srgb, var(--text-on-brand) 12%, transparent);
}

.media-close:not(:disabled):active {
  background: color-mix(in srgb, var(--text-on-brand) 20%, transparent);
}
</style>
