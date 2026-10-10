<script setup lang="ts">
// The media viewer at /media/:id (SPEC-107): a wiring hub under the same
// rule as DocumentView and HomeView (MediaView.size.spec.ts holds the line
// budget). Load/kind-guard/blob-URL logic lives in useMediaDocument,
// fit/zoom/pan state in useImageViewport, the image editor's session
// (SPEC-112) in useImageEditor, and the surfaces in the components/media/*
// components — this file only wires them together under SPEC-058's app bar
// with SPEC-060's Close. There is no autosave session, version history,
// comments or editor File menu: a media document changes only through the
// image editor's explicit Save, Save as a copy and Flatten.
import { computed } from "vue";
import { useRoute, useRouter } from "vue-router";
import AppShell from "../components/AppShell.vue";
import MediaBarActions from "../components/media/MediaBarActions.vue";
import MediaImageCanvas from "../components/media/MediaImageCanvas.vue";
import MediaImageEditor from "../components/media/MediaImageEditor.vue";
import MediaPdfEmbed from "../components/media/MediaPdfEmbed.vue";
import MediaDetailsPanel from "../components/media/MediaDetailsPanel.vue";
import { useTheme } from "../composables/useTheme";
import { formatBadgeLabel } from "../lib/format";
import { downloadMedia, mediaFileName } from "../lib/media/download";
import { isEditableImage } from "../lib/media/editModel";
import { useMediaDocument } from "./media/useMediaDocument";
import { useImageViewport } from "./media/useImageViewport";
import { useImageEditor } from "./media/useImageEditor";

const route = useRoute();
const router = useRouter();

// No menu bar here — the shared theme composable keeps the workspace themed
// by the app tokens and following the OS live (SPEC-059).
useTheme();

const mediaId = computed(() => String(route.params.id ?? ""));
const { media, blobUrl, storedAt, close, reload } = useMediaDocument({ mediaId, router });
const editor = useImageEditor({ media, onStored: reload });

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
      <MediaBarActions
        :can-edit="media !== null && isEditableImage(media)"
        :editing="editor.active.value"
        @edit="editor.enter()"
        @close="close"
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
      <MediaImageEditor
        v-else-if="editor.active.value"
        :editor="editor"
      />
      <MediaImageCanvas
        v-else
        :viewport="viewport"
        :src="blobUrl"
        :alt="fileName"
      />
      <MediaDetailsPanel
        :media="editor.liveMedia.value ?? media"
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
</style>
