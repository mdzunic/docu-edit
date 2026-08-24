<script setup lang="ts">
// The presentation workspace at /slides/:id (SPEC-079): a WYSIWYG slide stage
// that scales to fit the workspace (letterboxed, correct aspect ratio) plus a
// keyboard-navigable filmstrip — the read-only outline and its note are gone.
// The view owns current-slide state, the stage's fit scaling (a
// ResizeObserver), global PageUp/PageDown navigation, and the on-open upgrade
// of preparation-milestone records. It still redirects a text-document id to
// /doc/:id (DocumentView does the reverse). Runtime declarations, not
// defineProps<{...}>() generics — espree lints .vue scripts.
import { computed, nextTick, onBeforeUnmount, onMounted, ref, shallowRef, useTemplateRef, watch } from "vue";
import { useRoute, useRouter } from "vue-router";
import AppShell from "../components/AppShell.vue";
import UiButton from "../components/ui/UiButton.vue";
import SlideStage from "../components/presentation/SlideStage.vue";
import SlideFilmstrip from "../components/presentation/SlideFilmstrip.vue";
import { useTheme } from "../composables/useTheme";
import { documentKindOf, getPresentation, initLibrary, storePresentation } from "../lib/documentStore";
import { downloadPresentation } from "../lib/presentation/download";
import { isLegacyPresentationModel } from "../lib/presentation/model";
import { upgradePresentationModel } from "../lib/presentation/open";
import { cmToPx } from "../lib/core/units";
import { formatBadgeLabel, presentationFormatLabel } from "../lib/format";

const route = useRoute();
const router = useRouter();

// No menu bar here — read the shared theme composable so the surrounding
// workspace is themed by the app tokens and follows the OS live (SPEC-059).
useTheme();

const presentationId = computed(() => String(route.params.id ?? ""));

// A legacy (preparation-milestone) model still holds outline-shaped slides,
// which the frame renderer cannot draw — hide it until loadPresentation has
// upgraded it, so the filmstrip/stage only ever see the frame model.
function initialModel() {
  const model = getPresentation(presentationId.value) ?? null;
  return model && isLegacyPresentationModel(model) ? null : model;
}

const presentation = shallowRef(initialModel());
const currentIndex = ref(0);

async function loadPresentation() {
  // A text document's id belongs to the word processor — hand it over.
  if (documentKindOf(presentationId.value) === "text") {
    void router.replace(`/doc/${presentationId.value}`);
    return;
  }
  let model = getPresentation(presentationId.value) ?? null;
  // Preparation-milestone records upgrade on open (AC-54..56).
  if (model && isLegacyPresentationModel(model)) {
    try {
      const upgraded = await upgradePresentationModel(model);
      storePresentation(upgraded);
      model = getPresentation(presentationId.value) ?? upgraded;
    } catch {
      // Reading stays tolerant: if the retained package cannot be re-parsed,
      // keep the record as-is rather than blanking the view.
    }
  }
  presentation.value = model;
  currentIndex.value = 0;
  void nextTick(recomputeScale);
}

// main.ts awaits initLibrary() before mounting, so the store is hydrated in
// the running app; re-syncing after it resolves matters only when the view is
// mounted directly (tests) — the same pattern HomeView uses.
void loadPresentation();
void initLibrary().then(loadPresentation);
watch(presentationId, () => {
  void loadPresentation();
});

const title = computed(() => presentation.value?.title ?? "No document");
const badge = computed(() => (presentation.value ? formatBadgeLabel(presentation.value.format) : ""));
const formatName = computed(() => (presentation.value ? presentationFormatLabel(presentation.value.format) : ""));
const slideTotal = computed(() => presentation.value?.slides.length ?? 0);
const slideCountLabel = computed(() => (slideTotal.value === 1 ? "1 slide" : `${slideTotal.value} slides`));

const currentSlide = computed(() => presentation.value?.slides[currentIndex.value] ?? null);
const liveLabel = computed(() =>
  presentation.value ? `Slide ${currentIndex.value + 1} of ${slideTotal.value}` : "",
);
const deckMedia = computed(() => presentation.value?.media ?? []);

function selectSlide(index = 0) {
  if (index >= 0 && index < slideTotal.value) {
    currentIndex.value = index;
  }
}

// --- fit-to-workspace scaling ----------------------------------------------

const stageContainer = useTemplateRef("stageContainer");
const stageScale = ref(1);
const resizeObserver = typeof ResizeObserver !== "undefined" ? new ResizeObserver(recomputeScale) : null;

function recomputeScale() {
  const container = stageContainer.value;
  const model = presentation.value;
  if (!container || !model) {
    return;
  }
  const availableW = container.clientWidth;
  const availableH = container.clientHeight;
  const surfaceW = cmToPx(model.slideWidthCm);
  const surfaceH = cmToPx(model.slideHeightCm);
  if (surfaceW <= 0 || surfaceH <= 0 || availableW <= 0 || availableH <= 0) {
    return;
  }
  stageScale.value = Math.min(availableW / surfaceW, availableH / surfaceH);
}

watch([presentation, currentSlide], () => {
  void nextTick(recomputeScale);
});

// --- global slide navigation (AC-48) ---------------------------------------

function onGlobalKeydown(event = new KeyboardEvent("keydown")) {
  if (event.key !== "PageUp" && event.key !== "PageDown") {
    return;
  }
  const target = event.target;
  if (target instanceof HTMLElement && (target.tagName === "INPUT" || target.tagName === "TEXTAREA" || target.isContentEditable)) {
    return;
  }
  event.preventDefault();
  selectSlide(currentIndex.value + (event.key === "PageDown" ? 1 : -1));
}

onMounted(() => {
  if (resizeObserver && stageContainer.value) {
    resizeObserver.observe(stageContainer.value);
  }
  window.addEventListener("resize", recomputeScale);
  window.addEventListener("keydown", onGlobalKeydown);
  void nextTick(recomputeScale);
});

onBeforeUnmount(() => {
  resizeObserver?.disconnect();
  window.removeEventListener("resize", recomputeScale);
  window.removeEventListener("keydown", onGlobalKeydown);
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

      <div class="deck-body">
        <SlideFilmstrip
          :slides="presentation.slides"
          :width-cm="presentation.slideWidthCm"
          :height-cm="presentation.slideHeightCm"
          :media="deckMedia"
          :current-index="currentIndex"
          @select="selectSlide"
        />
        <div
          ref="stageContainer"
          class="stage-workspace"
          data-testid="stage-workspace"
        >
          <SlideStage
            v-if="currentSlide"
            :slide="currentSlide"
            :width-cm="presentation.slideWidthCm"
            :height-cm="presentation.slideHeightCm"
            :media="deckMedia"
            :scale="stageScale"
            :live-label="liveLabel"
          />
        </div>
      </div>
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
  display: flex;
  flex-direction: column;
  height: 100%;
  min-height: 0;
}

.deck-header {
  display: flex;
  align-items: flex-start;
  gap: var(--space-4);
  flex-wrap: wrap;
  padding: var(--space-4) var(--space-5);
  border-bottom: var(--border-width) solid var(--border-subtle);
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

.deck-body {
  flex: 1 1 auto;
  min-height: 0;
  display: grid;
  grid-template-columns: minmax(180px, 220px) 1fr;
}

.deck-body :deep(.filmstrip) {
  border-right: var(--border-width) solid var(--border-subtle);
  max-height: 100%;
}

.stage-workspace {
  position: relative;
  min-width: 0;
  min-height: 0;
  display: flex;
  align-items: center;
  justify-content: center;
  padding: var(--space-5);
  background: var(--surface-sunken);
  overflow: hidden;
}
</style>
