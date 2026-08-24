<script setup lang="ts">
// The presentation workspace at /slides/:id — deliberately a read-only
// preview, not an editor: the presentations product ships open/preserve/
// save-back first (the same arc SPEC-001..008 walked for text). It renders
// the parsed slide outline, offers the byte-preserving download, and
// redirects a text-document id to /doc/:id (DocumentView does the reverse),
// so a stale or hand-edited URL never opens a model in the wrong product.
import { computed, shallowRef, watch } from "vue";
import { useRoute, useRouter } from "vue-router";
import AppShell from "../components/AppShell.vue";
import UiButton from "../components/ui/UiButton.vue";
import { useTheme } from "../composables/useTheme";
import { documentKindOf, getPresentation, initLibrary } from "../lib/documentStore";
import { downloadPresentation } from "../lib/presentation/download";
import { formatBadgeLabel, presentationFormatLabel } from "../lib/format";

const route = useRoute();
const router = useRouter();

// No menu bar here — read the shared theme composable so the view is themed
// by the same tokens and follows the OS live (the SPEC-059 convention every
// menubar-less view uses).
useTheme();

const presentationId = computed(() => String(route.params.id ?? ""));

function loadPresentation() {
  return getPresentation(presentationId.value) ?? null;
}

const presentation = shallowRef(loadPresentation());

function syncFromStore() {
  // A text document's id belongs to the word processor — hand it over.
  if (documentKindOf(presentationId.value) === "text") {
    void router.replace(`/doc/${presentationId.value}`);
    return;
  }
  presentation.value = loadPresentation();
}

// main.ts awaits initLibrary() before mounting, so the store is hydrated in
// the running app; re-syncing after it resolves matters only when the view
// is mounted directly (tests) — the same pattern HomeView uses.
syncFromStore();
initLibrary().then(syncFromStore);
watch(presentationId, syncFromStore);

const title = computed(() => presentation.value?.title ?? "No document");
const badge = computed(() => (presentation.value ? formatBadgeLabel(presentation.value.format) : ""));
const formatName = computed(() =>
  presentation.value ? presentationFormatLabel(presentation.value.format) : "",
);
const slideCountLabel = computed(() => {
  const count = presentation.value?.slides.length ?? 0;
  return count === 1 ? "1 slide" : `${count} slides`;
});

function download() {
  if (presentation.value) {
    void downloadPresentation(presentation.value);
  }
}

function close() {
  // Mirrors the editor's close (useFileActions): the library scrolls to and
  // highlights the row that was just closed (SPEC-060).
  router.push({ path: "/", state: { closedDocId: presentationId.value } });
}
</script>

<template>
  <AppShell
    :document-title="title"
    :format-badge="badge"
  >
    <div
      v-if="!presentation"
      class="state-panel"
      data-testid="presentation-not-found"
    >
      <h2>Presentation not found</h2>
      <p>This link doesn't match a presentation in your library.</p>
      <UiButton
        data-testid="presentation-back"
        @click="router.push('/')"
      >
        Back to library
      </UiButton>
    </div>

    <div
      v-else
      class="presentation"
      data-testid="presentation-view"
    >
      <header class="deck-header">
        <div class="deck-facts">
          <h2 class="deck-title">
            {{ presentation.title }}
          </h2>
          <p class="deck-meta">
            <span data-testid="presentation-slide-count">{{ slideCountLabel }}</span>
            · {{ formatName }}
          </p>
        </div>
        <div class="deck-actions">
          <UiButton
            data-testid="presentation-download"
            @click="download"
          >
            Download
          </UiButton>
          <UiButton
            variant="ghost"
            data-testid="presentation-close"
            @click="close"
          >
            Close
          </UiButton>
        </div>
      </header>

      <p
        class="deck-note"
        data-testid="presentation-note"
      >
        Read-only preview — editing presentations isn't supported yet. Download
        returns the file with every part of the original preserved.
      </p>

      <ol class="slide-list">
        <li
          v-for="(slide, index) in presentation.slides"
          :key="index"
          class="slide-card"
          data-testid="presentation-slide"
        >
          <span class="slide-number">{{ index + 1 }}</span>
          <div class="slide-body">
            <h3
              v-if="slide.title"
              class="slide-title"
              data-testid="presentation-slide-title"
            >
              {{ slide.title }}
            </h3>
            <p
              v-for="(line, lineIndex) in slide.lines"
              :key="lineIndex"
              class="slide-line"
              data-testid="presentation-slide-line"
            >
              {{ line }}
            </p>
            <p
              v-if="!slide.title && slide.lines.length === 0"
              class="slide-empty"
            >
              (no text on this slide)
            </p>
          </div>
        </li>
      </ol>
    </div>
  </AppShell>
</template>

<style scoped>
.state-panel {
  height: 100%;
  display: flex;
  flex-direction: column;
  align-items: center;
  justify-content: center;
  gap: var(--space-4);
  color: var(--text-primary);
  text-align: center;
  padding: var(--space-6);
}

.state-panel h2 {
  margin: 0;
}

.state-panel p {
  margin: 0;
  color: var(--text-secondary);
}

.presentation {
  max-width: 48rem;
  margin: 0 auto;
  padding: var(--space-7) var(--space-5);
}

.deck-header {
  display: flex;
  align-items: flex-start;
  gap: var(--space-4);
  flex-wrap: wrap;
}

.deck-facts {
  flex: 1 1 auto;
  min-width: 0;
}

.deck-title {
  margin: 0;
  font-size: var(--text-xl);
  font-weight: var(--weight-bold);
  color: var(--text-primary);
  overflow-wrap: anywhere;
}

.deck-meta {
  margin: var(--space-1) 0 0;
  color: var(--text-secondary);
  font-size: var(--text-sm);
}

.deck-actions {
  display: flex;
  gap: var(--space-3);
}

.deck-note {
  margin: var(--space-4) 0 var(--space-5);
  padding: var(--space-3) var(--space-4);
  border-radius: var(--radius-md);
  background: var(--surface-raised);
  color: var(--text-secondary);
  font-size: var(--text-sm);
}

.slide-list {
  list-style: none;
  margin: 0;
  padding: 0;
  display: flex;
  flex-direction: column;
  gap: var(--space-4);
}

.slide-card {
  display: flex;
  gap: var(--space-4);
  padding: var(--space-4) var(--space-5);
  background: var(--surface-base);
  border: var(--border-width) solid var(--border-subtle);
  border-radius: var(--radius-md);
  box-shadow: var(--shadow-1);
}

.slide-number {
  flex: none;
  color: var(--text-secondary);
  font-size: var(--text-sm);
  font-weight: var(--weight-bold);
  min-width: 1.5rem;
  text-align: right;
}

.slide-body {
  min-width: 0;
}

.slide-title {
  margin: 0 0 var(--space-2);
  font-size: var(--text-lg);
  font-weight: var(--weight-bold);
  color: var(--text-primary);
  overflow-wrap: anywhere;
}

.slide-line {
  margin: 0 0 var(--space-1);
  color: var(--text-primary);
  overflow-wrap: anywhere;
}

.slide-empty {
  margin: 0;
  color: var(--text-secondary);
  font-style: italic;
}
</style>
