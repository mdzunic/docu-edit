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
import PresentationToolbar from "../components/presentation/PresentationToolbar.vue";
import SlideInsertBar from "../components/presentation/SlideInsertBar.vue";
import DeckActions from "../components/presentation/DeckActions.vue";
import HtmlExportOptionsDialog from "../components/presentation/HtmlExportOptionsDialog.vue";
import NotesPanel from "../components/presentation/NotesPanel.vue";
import TransitionsPanel from "../components/presentation/TransitionsPanel.vue";
import AnimationsPanel from "../components/presentation/AnimationsPanel.vue";
import AnimationStageLayer from "../components/presentation/AnimationStageLayer.vue";
import SlideTransitionLayer from "../components/presentation/SlideTransitionLayer.vue";
import { useTheme } from "../composables/useTheme";
import { useToasts } from "../composables/useToasts";
import { documentKindOf, getPresentation, initLibrary, storePresentation } from "../lib/documentStore";
import { isLegacyPresentationModel } from "../lib/presentation/model";
import { upgradePresentationModel } from "../lib/presentation/open";
import { cmToPx } from "../lib/core/units";
import { formatBadgeLabel, presentationFormatLabel } from "../lib/format";
import { usePresentationUndo } from "./presentation/usePresentationUndo";
import { useSlideCommands } from "./presentation/useSlideCommands";
import { usePresentationEditing } from "./presentation/usePresentationEditing";
import { useNotes } from "./presentation/useNotes";
import { useTransitions } from "./presentation/useTransitions";
import { useAnimations } from "./presentation/useAnimations";
import { useSlideRender } from "./presentation/useSlideRender";
import { useDeckActions } from "./presentation/useDeckActions";

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

// The re-render/persist trio lives in a composable (SPEC-086 extraction) so the
// view stays within its size budget; it owns the shallowRef re-identification.
const { commit, refresh } = useSlideRender({ model: presentation, currentIndex, recomputeScale: () => recomputeScale() });

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
  // Reselect the slide the play view was showing on exit (SPEC-084 `?return`),
  // else the first slide.
  currentIndex.value = deckActions.returnIndex();
  slideUndo.reset(); // the undo stack is per-open-view (AC-45)
  void nextTick(recomputeScale);
}

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

// --- deck header actions: download + playback entry (SPEC-080/084) ---------
const deckActions = useDeckActions({
  model: presentation,
  presentationId,
  currentIndex,
  slideTotal,
  router,
  route,
  reportError: (message) => addToast(message),
});

// main.ts awaits initLibrary() before mounting, so re-syncing after it resolves
// matters only when the view is mounted directly (tests). These run after
// deckActions is defined so loadPresentation's return-selection read is safe.
void loadPresentation();
void initLibrary().then(loadPresentation);
watch(presentationId, () => {
  void loadPresentation();
});

// --- selection + text editing + toolbar (SPEC-081) -------------------------
const stageEdit = usePresentationEditing({
  model: presentation,
  currentSlide,
  currentIndex,
  record: slideUndo.record,
  commit,
  refresh,
  undo: slideUndo.undo,
  redo: slideUndo.redo,
  navigate: (delta) => selectSlide(currentIndex.value + delta),
  reportError: (message) => addToast(message),
});

// --- speaker notes (SPEC-083) ----------------------------------------------
const notes = useNotes({
  model: presentation,
  currentSlide,
  currentIndex,
  record: slideUndo.record,
  commit,
});

// --- slide transitions (SPEC-085) ------------------------------------------
const transitions = useTransitions({
  model: presentation,
  currentSlide,
  currentIndex,
  record: slideUndo.record,
  commit,
});

// --- frame animations (SPEC-086) -------------------------------------------
const animations = useAnimations({
  model: presentation,
  currentSlide,
  currentIndex,
  record: slideUndo.record,
  commit,
  selectedFrameId: stageEdit.selection.selectedFrameId,
  selectFrame: stageEdit.selection.select,
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
  window.addEventListener("keydown", notes.onGlobalKeydown);
  window.addEventListener("keydown", deckActions.onPresentShortcut);
  void nextTick(recomputeScale);
});

onBeforeUnmount(() => {
  resizeObserver?.disconnect();
  window.removeEventListener("resize", recomputeScale);
  window.removeEventListener("keydown", stageEdit.onGlobalKeydown);
  window.removeEventListener("keydown", notes.onGlobalKeydown);
  window.removeEventListener("keydown", deckActions.onPresentShortcut);
});

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
        <DeckActions
          :can-undo="canUndo"
          :can-redo="canRedo"
          :notes-open="notes.open.value"
          :can-present="deckActions.canPresent.value"
          :other-format-label="deckActions.otherFormatLabel.value"
          @undo="slideUndo.undo()"
          @redo="slideUndo.redo()"
          @toggle-notes="notes.toggle()"
          @present="deckActions.present()"
          @download="deckActions.download()"
          @download-copy="deckActions.downloadCopy()"
          @export-html="deckActions.openHtmlExport()"
          @close="deckActions.close()"
        />
      </header>

      <PresentationToolbar
        :state="stageEdit.editing.toolbarState.value ?? undefined"
        :enabled="stageEdit.editing.toolbarEnabled.value"
        :recent-colors="stageEdit.recentColors.value"
        :transition-open="transitions.open.value"
        :animate-open="animations.open.value"
        @transition-toggle="transitions.toggle()"
        @animate-toggle="animations.toggle()"
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

      <SlideInsertBar
        :selected-frame="stageEdit.selection.selectedFrame.value ?? undefined"
        :armed-kind="stageEdit.insertion.armed.value?.kind ?? ''"
        @insert-text-box="stageEdit.insertion.armTextBox()"
        @insert-shape="stageEdit.insertion.armShape($event)"
        @insert-image="stageEdit.insertion.insertImageFile($event)"
        @arrange="stageEdit.manipulation.arrange($event)"
        @ungroup="stageEdit.manipulation.ungroup()"
        @set-fill="stageEdit.setFill($event)"
        @set-stroke="stageEdit.setStroke($event)"
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
        <div class="stage-column">
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
              :guides="stageEdit.manipulation.guides.value"
              :draw-preview="stageEdit.insertion.preview.value ?? undefined"
              :armed="stageEdit.insertion.isArmed.value"
              @frame-select="stageEdit.onFrameSelect"
              @frame-start-edit="stageEdit.onStartEdit"
              @frame-pointerdown="stageEdit.onFramePointerDown"
              @handle-pointerdown="stageEdit.onHandlePointerDown"
              @edit-mounted="stageEdit.editing.registerHost"
              @edit-unmounted="stageEdit.editing.registerHost(null)"
              @surface-pointerdown="stageEdit.onSurfaceDeselect"
              @surface-draw="stageEdit.onSurfaceDraw"
              @surface-mounted="stageEdit.registerSurface"
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
            <AnimationStageLayer
              v-if="animations.open.value && currentSlide"
              :controller="animations"
              :slide="currentSlide"
              :width-cm="presentation.slideWidthCm"
              :height-cm="presentation.slideHeightCm"
              :media="deckMedia"
              :scale="stageScale"
            />
            <SlideTransitionLayer
              v-if="transitions.previewState.value"
              class="stage-preview"
              :outgoing="transitions.previewState.value.outgoing"
              :incoming="transitions.previewState.value.incoming"
              :transition="transitions.previewState.value.transition"
              :width-cm="presentation.slideWidthCm"
              :height-cm="presentation.slideHeightCm"
              :media="deckMedia"
              :scale="stageScale"
              @done="transitions.onPreviewDone()"
            />
          </div>
          <NotesPanel
            v-if="notes.open.value && currentSlide"
            :value="notes.notesText.value"
            :slide-number="currentIndex + 1"
            @input="notes.onInput"
          />
          <TransitionsPanel
            v-if="transitions.open.value && currentSlide"
            :controller="transitions"
            :slide-number="currentIndex + 1"
          />
          <AnimationsPanel
            v-if="animations.open.value && currentSlide"
            :controller="animations"
            :slide-number="currentIndex + 1"
          />
        </div>
      </div>
    </div>
    <HtmlExportOptionsDialog
      v-if="deckActions.htmlDialogOpen.value"
      @confirm="deckActions.confirmHtmlExport($event)"
      @cancel="deckActions.cancelHtmlExport()"
    />
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

.stage-preview {
  position: absolute;
  inset: 0;
  display: flex;
  align-items: center;
  justify-content: center;
  z-index: 5;
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

.stage-column {
  min-width: 0;
  min-height: 0;
  display: flex;
  flex-direction: column;
}

.stage-workspace {
  position: relative;
  flex: 1 1 auto;
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
