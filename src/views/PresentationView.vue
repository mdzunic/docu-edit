<script setup lang="ts">
// The presentation workspace at /slides/:id (SPEC-079): a WYSIWYG slide stage
// that scales to fit the workspace (letterboxed, correct aspect ratio) plus a
// keyboard-navigable filmstrip — the read-only outline and its note are gone.
// The view owns current-slide state, the stage's fit scaling (a
// ResizeObserver), global PageUp/PageDown navigation, and the on-open upgrade
// of preparation-milestone records. It still redirects a text-document id to
// /doc/:id (DocumentView does the reverse). Runtime declarations, not
// defineProps<{...}>() generics — espree lints .vue scripts.
import { computed, ref } from "vue";
import { useRoute, useRouter } from "vue-router";
import AppShell from "../components/AppShell.vue";
import UiButton from "../components/ui/UiButton.vue";
import DeckBody from "../components/presentation/DeckBody.vue";
import { stageFields } from "../components/presentation/slideRender";
import { now } from "../lib/clock";
import PresentationToolbar from "../components/presentation/PresentationToolbar.vue";
import SlideInsertBar from "../components/presentation/SlideInsertBar.vue";
import DeckHeader from "../components/presentation/DeckHeader.vue";
import PresentationDialogs from "../components/presentation/PresentationDialogs.vue";
import HeaderFooterDialog from "../components/presentation/HeaderFooterDialog.vue";
import ThemesPanel from "../components/presentation/ThemesPanel.vue";
import { useTheme } from "../composables/useTheme";
import { useToasts } from "../composables/useToasts";
import { formatBadgeLabel, presentationFormatLabel } from "../lib/format";
import { useDeckLoad } from "./presentation/useDeckLoad";
import { usePresentationUndo } from "./presentation/usePresentationUndo";
import { useSlideCommands } from "./presentation/useSlideCommands";
import { usePresentationEditing } from "./presentation/usePresentationEditing";
import { useNotes } from "./presentation/useNotes";
import { useEditingPanels } from "./presentation/useEditingPanels";
import { useDeckTheme } from "./presentation/useDeckTheme";
import { useFilmstripCommands } from "./presentation/useFilmstripCommands";
import { useSlideLayouts } from "./presentation/useSlideLayouts";
import { useSlideRender } from "./presentation/useSlideRender";
import { useWindowKeys } from "./presentation/useWindowKeys";
import { useDeckActions } from "./presentation/useDeckActions";
import { useDeckSetup } from "./presentation/useDeckSetup";
import { useHeaderFooter } from "./presentation/useHeaderFooter";

const route = useRoute();
const router = useRouter();

// No menu bar here — read the shared theme composable so the surrounding
// workspace is themed by the app tokens and follows the OS live (SPEC-059).
useTheme();

const presentationId = computed(() => String(route.params.id ?? ""));
const currentIndex = ref(0);

// Deck loading + legacy upgrade live in useDeckLoad (SPEC-096 extraction);
// start() runs after deckActions exists (it reads the ?return selection).
const deckLoad = useDeckLoad({
  presentationId,
  router,
  returnIndex: () => deckActions.returnIndex(),
  currentIndex,
  resetUndo: () => slideUndo.reset(),
});
const presentation = deckLoad.presentation;

const { addToast } = useToasts();

// --- structural commands + undo (SPEC-080) ---------------------------------

// The re-render/persist trio lives in a composable (SPEC-086 extraction) so the
// view stays within its size budget; it owns the shallowRef re-identification.
// SlideStageArea owns the fit scaling now (SPEC-089) and recomputes reactively
// from the current slide + deck size, so the render trio needs no scale hook.
const { commit, refresh, bumpMedia } = useSlideRender({ model: presentation, currentIndex, recomputeScale: () => {} });

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
// SPEC-093: each slide's resolved field band, read off the deck's shallowRef so
// a commit (triggerRef) recomputes footers/numbers/dates live (AC-27/30/31).
const deckFields = computed(() => {
  const model = presentation.value;
  return model ? model.slides.map((_, index) => stageFields(model, index, now())) : [];
});

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

deckLoad.start();

// --- selection + text editing + toolbar (SPEC-081) -------------------------
const stageEdit = usePresentationEditing({
  model: presentation,
  currentSlide,
  currentIndex,
  record: slideUndo.record,
  commit,
  bumpMedia,
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

// --- transition/animation/table/chart panels (SPEC-085/086/091/092) --------
const { transitions, animations, tables, charts, media } = useEditingPanels({
  model: presentation,
  currentSlide,
  currentIndex,
  record: slideUndo.record,
  commit,
  bumpMedia,
  reportError: (message) => addToast(message),
  stageEdit,
});

// --- deck theme: the Design panel (SPEC-094) -------------------------------
const themePanel = useDeckTheme({ model: presentation, record: slideUndo.record, commit });

// --- slide layouts + Reset slide (SPEC-089) --------------------------------
const layouts = useSlideLayouts({
  model: presentation,
  currentSlide,
  currentIndex,
  record: slideUndo.record,
  commit,
  addSlide: slideCommands.add,
});
function toggleLayout() {
  if (layouts.open.value) {
    layouts.close();
  } else {
    layouts.openForChange();
  }
}

// --- filmstrip structural commands (SPEC-080) ------------------------------
const filmstrip = useFilmstripCommands({
  commands: slideCommands,
  selectSlide,
  openLayoutForNew: layouts.openForNew,
  openLayoutForChange: layouts.openForChange,
});

// --- deck setup: slide size, first slide number, metadata (SPEC-090) -------
const deckSetup = useDeckSetup({
  model: presentation,
  currentIndex,
  record: slideUndo.record,
  commit,
  reportError: (message) => addToast(message),
});

// --- header & footer: deck footer/date fields (SPEC-093) -------------------
const headerFooter = useHeaderFooter({ model: presentation, currentSlide, currentIndex, record: slideUndo.record, commit });

// --- global keyboard: undo/redo, navigation, selection, edit (SPEC-080/081) -
useWindowKeys([stageEdit.onGlobalKeydown, notes.onGlobalKeydown, deckActions.onPresentShortcut]);

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
      <DeckHeader
        :deck-title="presentation.title"
        :slide-count-label="slideCountLabel"
        :format-name="formatName"
        :can-undo="canUndo"
        :can-redo="canRedo"
        :notes-open="notes.open.value"
        :can-present="deckActions.canPresent.value"
        :other-format-label="deckActions.otherFormatLabel.value"
        @undo="slideUndo.undo()"
        @redo="slideUndo.redo()"
        @toggle-notes="notes.toggle()"
        @deck-setup="deckSetup.open()"
        @present="deckActions.present()"
        @download="deckActions.download()"
        @download-copy="deckActions.downloadCopy()"
        @export-html="deckActions.openHtmlExport()"
        @close="deckActions.close()"
      />

      <PresentationToolbar
        :state="stageEdit.editing.toolbarState.value ?? undefined"
        :enabled="stageEdit.editing.toolbarEnabled.value"
        :recent-colors="stageEdit.recentColors.value"
        :transition-open="transitions.open.value"
        :animate-open="animations.open.value"
        :layout-open="layouts.open.value"
        :table-open="tables.open.value"
        :media-open="media.open.value"
        :header-footer-open="headerFooter.open.value"
        :theme-open="themePanel.open.value"
        :theme="presentation.theme"
        @theme-toggle="themePanel.toggle()"
        @transition-toggle="transitions.toggle()"
        @animate-toggle="animations.toggle()"
        @layout-toggle="toggleLayout()"
        @table-toggle="tables.toggle()"
        @media-toggle="media.toggle()"
        @header-footer-toggle="headerFooter.show()"
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
        :selected-frame="stageEdit.selection.anchorFrame.value ?? undefined"
        :selected-count="stageEdit.arrange.selectedCount.value"
        :has-group="stageEdit.selection.selectedFrameIds.value.some((id) => currentSlide?.frames.find((f) => f.id === id)?.kind === 'group')"
        :can-distribute="stageEdit.arrange.canDistribute.value"
        :distribute-reason="stageEdit.arrange.distributeReason.value"
        :align-reference="stageEdit.arrange.alignReference.value"
        :theme="presentation.theme"
        :armed-kind="stageEdit.insertion.armed.value?.kind ?? ''"
        @insert-text-box="stageEdit.insertion.armTextBox()"
        @insert-shape="stageEdit.insertion.armShape($event)"
        @insert-image="stageEdit.insertion.insertImageFile($event)"
        @insert-media="stageEdit.insertion.insertMediaFile($event)"
        @insert-table="(rows, cols) => stageEdit.insertion.insertTable(rows, cols)"
        @insert-chart="charts.pickChartType($event)"
        @arrange="stageEdit.manipulation.arrange($event)"
        @align="stageEdit.arrange.align($event)"
        @distribute="stageEdit.arrange.distribute($event)"
        @align-reference="stageEdit.arrange.alignReference.value = $event"
        @group="stageEdit.arrange.group()"
        @ungroup="stageEdit.arrange.ungroup()"
        @position-toggle="stageEdit.arrange.openPosition()"
        @set-fill="stageEdit.setFill($event)"
        @set-stroke="stageEdit.setStroke($event)"
      />

      <DeckBody
        :slides="presentation.slides"
        :width-cm="presentation.slideWidthCm"
        :height-cm="presentation.slideHeightCm"
        :current-slide="currentSlide ?? undefined"
        :current-index="currentIndex"
        :media="deckMedia"
        :live-label="liveLabel"
        :stage-edit="stageEdit"
        :animations="animations"
        :transitions="transitions"
        :notes="notes"
        :tables="tables"
        :charts="charts"
        :media-panel="media"
        :fields-by-slide="deckFields"
        :theme="themePanel.displayTheme.value"
        @select="selectSlide"
        @add="addSlideAction"
        @command="filmstrip.onCommand"
        @reorder="filmstrip.onReorder"
      />

      <ThemesPanel :controller="themePanel" />
    </div>
    <PresentationDialogs
      :deck-actions="deckActions"
      :layouts="layouts"
      :deck-setup="deckSetup"
      :slide-number="currentIndex + 1"
    />
    <HeaderFooterDialog
      v-if="headerFooter.open.value"
      :controller="headerFooter"
      :slide-number="currentIndex + 1"
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

</style>
