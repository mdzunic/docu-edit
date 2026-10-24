<script setup lang="ts">
// The deck's whole ribbon — SlideRibbon's undo/redo pair and panel toggles,
// the menu bar in its #menu slot and both toolbar strips (SlideToolbarHost) —
// extracted from PresentationView (SPEC-104) so the view, a wiring hub under
// a line budget that only ratchets down, makes room for the print-and-export
// wiring. Like PresentationDialogs and DeckAppBar it takes the controllers
// whole and calls them directly instead of having the view re-wire a dozen
// props and events. Runtime prop declarations per the espree constraint.
import SlideRibbon from "./SlideRibbon.vue";
import SlideToolbarHost from "./SlideToolbarHost.vue";
import PresentationMenuBar from "./PresentationMenuBar.vue";

const props = defineProps({
  /** The usePresentationUndo controller (canUndo/canRedo, undo/redo). */
  undo: { type: Object, required: true },
  /** The panel controllers whose open state each toggle shows and flips:
   * notes, transitions, animations, layouts, tables, media, headerFooter,
   * themePanel, sorter, SPEC-103's history (useDeckSaving's) and SPEC-098's
   * search (useDeckSearch — the Find toggle). */
  panels: { type: Object, required: true },
  /** The usePresentationMenu controller (menu-bar state + command runner). */
  menu: { type: Object, required: true },
  /** The usePresentationEditing controller, for both toolbar strips. */
  stageEdit: { type: Object, required: true },
  /** The useSlideTextDepth controller (the Text panel toggle + toolbars). */
  textDepth: { type: Object, required: true },
  /** The useSlideCharts controller (Insert ▸ Chart). */
  charts: { type: Object, default: null },
  /** The deck theme (SPEC-094), for the colour pickers. */
  theme: { type: Object, default: null },
  /** The current slide (Insert: whether a selected frame is a group). */
  slide: { type: Object, default: null },
});

const p = props.panels;
</script>

<template>
  <SlideRibbon
    :can-undo="undo.canUndo.value"
    :can-redo="undo.canRedo.value"
    :notes-open="p.notes.open.value"
    :transition-open="p.transitions.open.value"
    :animate-open="p.animations.open.value"
    :layout-open="p.layouts.open.value"
    :table-open="p.tables.open.value"
    :media-open="p.media.open.value"
    :header-footer-open="p.headerFooter.open.value"
    :theme-open="p.themePanel.open.value"
    :text-open="textDepth.panelOpen.value"
    :sorter-open="p.sorter.open.value"
    :history-open="p.history?.open.value ?? false"
    :find-open="p.search?.open.value ?? false"
    @undo="undo.undo()"
    @redo="undo.redo()"
    @notes-toggle="p.notes.toggle()"
    @theme-toggle="p.themePanel.toggle()"
    @text-toggle="textDepth.togglePanel()"
    @transition-toggle="p.transitions.toggle()"
    @animate-toggle="p.animations.toggle()"
    @layout-toggle="p.layouts.toggle()"
    @table-toggle="p.tables.toggle()"
    @media-toggle="p.media.toggle()"
    @header-footer-toggle="p.headerFooter.show()"
    @sorter-toggle="p.sorter.toggle()"
    @history-toggle="p.history?.toggle()"
    @find-toggle="p.search?.toggle()"
  >
    <template #menu>
      <PresentationMenuBar
        :state="menu.state.value"
        @command="menu.run"
      />
    </template>
    <!-- Both strips' bindings live in SlideToolbarHost (SPEC-097). -->
    <template #home>
      <SlideToolbarHost
        :stage-edit="stageEdit"
        :text-depth="textDepth"
        :theme="theme"
      />
    </template>
    <template #insert>
      <SlideToolbarHost
        strip="insert"
        :stage-edit="stageEdit"
        :text-depth="textDepth"
        :theme="theme"
        :charts="charts"
        :slide="slide"
      />
    </template>
  </SlideRibbon>
</template>
