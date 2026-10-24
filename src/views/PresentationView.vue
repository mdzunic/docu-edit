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
import PresentationNotFound from "../components/presentation/PresentationNotFound.vue";
import DeckBody from "../components/presentation/DeckBody.vue";
import SlideSorter from "../components/presentation/SlideSorter.vue";
import PresentationRibbon from "../components/presentation/PresentationRibbon.vue";
import DeckAppBar from "../components/presentation/DeckAppBar.vue";
import PresentationDialogs from "../components/presentation/PresentationDialogs.vue";
import RecoveryBar from "../components/RecoveryBar.vue";
import ThemesPanel from "../components/presentation/ThemesPanel.vue";
import TextPanel from "../components/presentation/TextPanel.vue";
import DeckSearchResults from "../components/presentation/DeckSearchResults.vue";
import { useTheme } from "../composables/useTheme";
import { useToasts } from "../composables/useToasts";
import { useDeckLoad } from "./presentation/useDeckLoad";
import { usePresentationUndo } from "./presentation/usePresentationUndo";
import { useDeckNavigation } from "./presentation/useDeckNavigation";
import { usePresentationEditing } from "./presentation/usePresentationEditing";
import { useEditingPanels } from "./presentation/useEditingPanels";
import { useDeckTheme } from "./presentation/useDeckTheme";
import { useSlideRender } from "./presentation/useSlideRender";
import { useSlideSorter } from "./presentation/useSlideSorter";
import { usePresentationMenu } from "./presentation/usePresentationMenu";
import { usePreferencesDialog } from "../composables/usePreferencesDialog";
import { useWindowKeys } from "./presentation/useWindowKeys";
import { useDeckActions } from "./presentation/useDeckActions";
import { useSlideTextDepth } from "./presentation/useSlideTextDepth";
import { useDeckDialogs } from "./presentation/useDeckDialogs";
import { useDeckSaving } from "./presentation/useDeckSaving";
import { useDeckSearch } from "./presentation/useDeckSearch";
import { stageSearchHooks } from "./presentation/stageSearchHooks";
import { useSlideSpellCheck } from "./presentation/useSlideSpellCheck";

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

// The re-render trio lives in a composable (SPEC-086 extraction) so the view
// stays within its size budget; it owns the shallowRef re-identification.
// SlideStageArea owns the fit scaling now (SPEC-089) and recomputes reactively
// from the current slide + deck size, so the render trio needs no scale hook.
// Every finished edit reports to the deck's autosave and history (SPEC-103).
const { commit, refresh, bumpMedia } = useSlideRender({ model: presentation, currentIndex, recomputeScale: () => {}, onCommit: () => saving.noteEdit() });

const slideUndo = usePresentationUndo({ model: presentation, onApply: commit });
// Deck navigation — slide commands, derived facts, Add slide and the filmstrip
// dispatch — lives in useDeckNavigation (SPEC-099 extraction).
const { slideCommands, title, badge, formatName, slideTotal, slideCountLabel, currentSlide, liveLabel, deckMedia, deckFields, selectSlide, addSlideAction, onFilmstripCommand, onFilmstripReorder } =
  useDeckNavigation({ presentation, currentIndex, record: slideUndo.record, commit, openLayoutForNew: () => layouts.openForNew(), openLayoutForChange: () => layouts.openForChange() });

// The deck context every slide composable below takes (SPEC-098 — the
// DocumentView `editorContext` idea, for slides): each call is
// `use*({ ...deck, ...extras })`, and a composable reads the keys it needs.
const deck = { model: presentation, currentSlide, currentIndex, record: slideUndo.record, commit, reportError: (message = "") => addToast(message) };

// --- deck header actions: download + playback entry (SPEC-080/084) ---------
const deckActions = useDeckActions({
  ...deck, presentationId, slideTotal, router, route, onCommitPoint: () => saving.commitPoint(), flushForClose: () => saving.flushForClose(),
});

// --- autosave, crash recovery and version history (SPEC-103) ---------------
const saving = useDeckSaving({ ...deck, presentationId });

deckLoad.start();

// --- selection + text editing + toolbar (SPEC-081) -------------------------
const stageEdit = usePresentationEditing({
  ...deck, bumpMedia, refresh, undo: slideUndo.undo, redo: slideUndo.redo, navigate: (delta) => selectSlide(currentIndex.value + delta),
});

// --- notes/transition/animation/table/chart panels (SPEC-083/085/086/091/092)
const { notes, transitions, animations, tables, charts, media } = useEditingPanels({ ...deck, bumpMedia, stageEdit });

// --- deck theme: the Design panel (SPEC-094) -------------------------------
const themePanel = useDeckTheme(deck);

// --- slide text depth: lists, spacing, autofit, links (SPEC-097) -----------
const textDepth = useSlideTextDepth({ ...deck, selection: stageEdit.selection, editing: stageEdit.editing });

// --- slide sorter, sections, slide clipboard, Reuse slides (SPEC-099) -----
const sorter = useSlideSorter(deck);

// --- layouts, deck setup, header & footer, print & export, Redesign (SPEC-089/090/093/104/115)
const { layouts, deckSetup, headerFooter, deckExport, redesign } = useDeckDialogs({ ...deck, slideTotal, addSlide: slideCommands.add });

// --- find & replace and spell check on slides (SPEC-098) -------------------
const search = useDeckSearch({ ...deck, ...stageSearchHooks(stageEdit, notes, presentation) });
const spell = useSlideSpellCheck({ ...deck, ...stageEdit, notesText: notes.notesText });
const proofing = { search, spell }; // the stage's highlight + squiggle layer (one stable prop)

// --- the menu bar (SPEC-049's model, for slides): routes to every composable above, so it is built last
const menu = usePresentationMenu({
  undo: slideUndo, deckActions, deckSetup, notes, transitions, animations, layouts, tables, media,
  headerFooter, themePanel, stageEdit, preferences: usePreferencesDialog(), redesign, addSlide: addSlideAction,
});

// --- global keyboard: undo/redo, navigation, selection, edit (SPEC-080/081) -
useWindowKeys([search.onGlobalKeydown, sorter.onGlobalKeydown, stageEdit.onGlobalKeydown, notes.onGlobalKeydown, deckActions.onPresentShortcut, deckExport.onPrintShortcut]);
</script>

<template>
  <AppShell
    :document-title="title"
    :format-badge="badge"
  >
    <!-- The deck's whole chrome — menus, undo/redo, panel toggles and the
         strip — is one component, and it goes in the shell's menubar region
         rather than inside the canvas: the canvas scrolls, and a scroll
         container clips the menus' popups. -->
    <template
      v-if="presentation"
      #menubar
    >
      <PresentationRibbon
        :undo="slideUndo"
        :panels="{ notes, transitions, animations, layouts, tables, media, headerFooter, themePanel, sorter, history: saving.history, search }"
        :menu="menu"
        :stage-edit="stageEdit"
        :text-depth="textDepth"
        :charts="charts"
        :theme="presentation.theme"
        :slide="currentSlide ?? undefined"
      />
    </template>
    <!-- One identity line, like the word processor's — see DeckAppBar. -->
    <template
      v-if="presentation"
      #app-bar-actions
    >
      <DeckAppBar
        :slide-count-label="slideCountLabel"
        :format-name="formatName"
        :save-state="saving.saveState.value"
        :save-state-text="saving.saveStateText.value"
        :deck-actions="deckActions"
        :deck-export="deckExport"
        :spell="spell"
      />
    </template>
    <PresentationNotFound
      v-if="!presentation"
      @back="router.push('/')"
    />

    <div
      v-else
      class="presentation"
      data-testid="presentation-view"
    >
      <RecoveryBar
        v-if="saving.recovery.value"
        :label="saving.recoveryLabel.value"
        @keep="saving.keepRecovered()"
        @discard="saving.discardRecovered()"
      />
      <DeckSearchResults :controller="search" />
      <DeckBody
        v-if="!sorter.open.value"
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
        :selection="sorter.selection"
        :history="saving.history"
        :proofing="proofing"
        @select="selectSlide"
        @add="addSlideAction"
        @command="onFilmstripCommand"
        @reorder="onFilmstripReorder"
      />
      <SlideSorter
        v-else
        :controller="sorter"
        :width-cm="presentation.slideWidthCm"
        :height-cm="presentation.slideHeightCm"
        :current-index="currentIndex"
        :media="deckMedia"
        :theme="themePanel.displayTheme.value"
        :fields-by-slide="deckFields"
        @command="onFilmstripCommand"
      />
      <ThemesPanel :controller="themePanel" />
      <TextPanel
        :controller="textDepth"
        :slide-number="currentIndex + 1"
      />
    </div>
    <PresentationDialogs
      :deck-actions="deckActions"
      :layouts="layouts"
      :deck-setup="deckSetup"
      :deck-export="deckExport"
      :reuse="sorter.reuse"
      :header-footer="headerFooter"
      :slide-number="currentIndex + 1"
      :redesign="redesign"
    />
  </AppShell>
</template>

<style scoped>
.presentation {
  display: flex;
  flex-direction: column;
  height: 100%;
  min-height: 0;
}
</style>
