<script setup lang="ts">
// The presentation workspace at /slides/:id (SPEC-079): a WYSIWYG slide stage
// that scales to fit the workspace (letterboxed, correct aspect ratio) plus a
// keyboard-navigable filmstrip — the read-only outline and its note are gone.
// The view owns current-slide state, the stage's fit scaling (a
// ResizeObserver), global PageUp/PageDown navigation, and the on-open upgrade
// of preparation-milestone records. It still redirects a text-document id to
// /doc/:id (DocumentView does the reverse). Runtime declarations, not
// defineProps<{...}>() generics — espree lints .vue scripts.
import { computed, nextTick, onBeforeUnmount, onMounted, ref, shallowRef, triggerRef, useTemplateRef, watch } from "vue";
import { useRoute, useRouter } from "vue-router";
import AppShell from "../components/AppShell.vue";
import UiButton from "../components/ui/UiButton.vue";
import SlideStage from "../components/presentation/SlideStage.vue";
import SlideFilmstrip from "../components/presentation/SlideFilmstrip.vue";
import PresentationToolbar from "../components/presentation/PresentationToolbar.vue";
import { useTheme } from "../composables/useTheme";
import { useToasts } from "../composables/useToasts";
import { documentKindOf, getPresentation, initLibrary, storePresentation, touchDocument } from "../lib/documentStore";
import { downloadPresentation } from "../lib/presentation/download";
import { isLegacyPresentationModel } from "../lib/presentation/model";
import { upgradePresentationModel } from "../lib/presentation/open";
import { cmToPx } from "../lib/core/units";
import { formatBadgeLabel, presentationFormatLabel } from "../lib/format";
import { usePresentationUndo } from "./presentation/usePresentationUndo";
import { useSlideCommands } from "./presentation/useSlideCommands";
import { usePresentationEditing } from "./presentation/usePresentationEditing";

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

const { addToast } = useToasts();

// --- structural commands + undo (SPEC-080) ---------------------------------

/** Re-render the shallowRef and persist write-through after a mutation. The
 * selection is clamped in case a delete dropped it past the tail. */
function commit() {
  const model = presentation.value;
  if (!model) {
    return;
  }
  if (currentIndex.value > model.slides.length - 1) {
    currentIndex.value = Math.max(0, model.slides.length - 1);
  }
  // A fresh array reference so the filmstrip/stage props re-render (commands
  // splice in place); the model object identity is kept so the store record
  // and undo snapshots stay in step.
  model.slides = model.slides.slice();
  touchDocument(model.id);
  triggerRef(presentation);
  void nextTick(recomputeScale);
}

const slideUndo = usePresentationUndo({ model: presentation, onApply: commit });
const canUndo = slideUndo.canUndo;
const canRedo = slideUndo.canRedo;
const slideCommands = useSlideCommands({
  model: presentation,
  currentIndex,
  record: slideUndo.record,
  commit,
});

function addSlideAction() {
  slideCommands.add(currentIndex.value);
}
function onFilmstripCommand(payload = { action: "", index: 0 }) {
  const { action, index } = payload;
  if (action === "new") {
    slideCommands.add(index);
  } else if (action === "duplicate") {
    slideCommands.duplicate(index);
  } else if (action === "delete") {
    slideCommands.remove(index);
  } else if (action === "move-up") {
    slideCommands.moveUp(index);
  } else if (action === "move-down") {
    slideCommands.moveDown(index);
  }
}
function onFilmstripReorder(payload = { from: 0, to: 0 }) {
  slideCommands.reorder(payload.from, payload.to);
}

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
  slideUndo.reset(); // the undo stack is per-open-view (AC-45)
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

// --- selection + text editing + toolbar (SPEC-081) -------------------------
const stageEdit = usePresentationEditing({
  model: presentation,
  currentSlide,
  record: slideUndo.record,
  commit,
  undo: slideUndo.undo,
  redo: slideUndo.redo,
  navigate: (delta) => selectSlide(currentIndex.value + delta),
});

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

// --- global keyboard: undo/redo, navigation, selection, edit (SPEC-080/081) -

onMounted(() => {
  if (resizeObserver && stageContainer.value) {
    resizeObserver.observe(stageContainer.value);
  }
  window.addEventListener("resize", recomputeScale);
  window.addEventListener("keydown", stageEdit.onGlobalKeydown);
  void nextTick(recomputeScale);
});

onBeforeUnmount(() => {
  resizeObserver?.disconnect();
  window.removeEventListener("resize", recomputeScale);
  window.removeEventListener("keydown", stageEdit.onGlobalKeydown);
});

// --- download (own format + "Download a copy as…" the other) ---------------

const downloadMenuOpen = ref(false);
const otherFormat = computed(() => (presentation.value?.format === "odp" ? "pptx" : "odp"));
const otherFormatLabel = computed(() => presentationFormatLabel(otherFormat.value));

function download() {
  if (presentation.value) {
    void downloadPresentation(presentation.value);
  }
}
async function downloadCopy() {
  downloadMenuOpen.value = false;
  const model = presentation.value;
  if (!model) {
    return;
  }
  const { droppedObjects } = await downloadPresentation(model, otherFormat.value);
  if (droppedObjects > 0) {
    addToast(`${droppedObjects} object${droppedObjects === 1 ? "" : "s"} couldn't be converted`);
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
            variant="ghost"
            data-testid="presentation-undo"
            :disabled="!canUndo"
            aria-label="Undo"
            @click="slideUndo.undo()"
          >
            Undo
          </UiButton>
          <UiButton
            variant="ghost"
            data-testid="presentation-redo"
            :disabled="!canRedo"
            aria-label="Redo"
            @click="slideUndo.redo()"
          >
            Redo
          </UiButton>
          <div class="download-split">
            <UiButton
              data-testid="presentation-download"
              @click="download"
            >
              Download
            </UiButton>
            <UiButton
              variant="ghost"
              class="download-caret"
              data-testid="presentation-download-menu"
              aria-haspopup="menu"
              :aria-expanded="downloadMenuOpen"
              aria-label="More download options"
              @click="downloadMenuOpen = !downloadMenuOpen"
            >
              ▾
            </UiButton>
            <div
              v-if="downloadMenuOpen"
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
            </div>
            <div
              v-if="downloadMenuOpen"
              class="menu-backdrop"
              @click="downloadMenuOpen = false"
            />
          </div>
          <UiButton
            variant="ghost"
            data-testid="presentation-close"
            @click="close"
          >
            Close
          </UiButton>
        </div>
      </header>

      <PresentationToolbar
        :state="stageEdit.editing.toolbarState.value ?? undefined"
        :enabled="stageEdit.editing.toolbarEnabled.value"
        :recent-colors="stageEdit.recentColors.value"
        @bold="stageEdit.editing.toggleBold()"
        @italic="stageEdit.editing.toggleItalic()"
        @underline="stageEdit.editing.toggleUnderline()"
        @font="stageEdit.editing.setFontFamily($event)"
        @size="stageEdit.editing.setFontSize($event)"
        @color="stageEdit.setColor($event)"
        @align="stageEdit.editing.setAlign($event)"
        @bullet="stageEdit.editing.toggleBulletCmd()"
        @indent="stageEdit.editing.indent(1)"
        @outdent="stageEdit.editing.indent(-1)"
      />

      <div class="deck-body">
        <SlideFilmstrip
          :slides="presentation.slides"
          :width-cm="presentation.slideWidthCm"
          :height-cm="presentation.slideHeightCm"
          :media="deckMedia"
          :current-index="currentIndex"
          @select="selectSlide"
          @add="addSlideAction"
          @command="onFilmstripCommand"
          @reorder="onFilmstripReorder"
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
            :selected-frame-id="stageEdit.selection.selectedFrameId.value ?? undefined"
            :editing-frame-id="stageEdit.editing.editingFrameId.value ?? undefined"
            @frame-select="stageEdit.onFrameSelect"
            @frame-start-edit="stageEdit.onStartEdit"
            @edit-mounted="stageEdit.editing.registerHost"
            @edit-unmounted="stageEdit.editing.registerHost(null)"
            @surface-pointerdown="stageEdit.onSurfaceDeselect"
          />
          <div
            v-else
            class="empty-stage"
            data-testid="presentation-empty"
          >
            <h3>This deck has no slides</h3>
            <UiButton
              data-testid="presentation-empty-add"
              @click="addSlideAction"
            >
              Add slide
            </UiButton>
          </div>
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
  align-items: center;
}

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

.empty-stage {
  display: flex;
  flex-direction: column;
  align-items: center;
  gap: var(--space-4);
  color: var(--text-secondary);
  text-align: center;
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
