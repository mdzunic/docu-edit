<script setup lang="ts">
// The presentation view's modal dialogs (HTML export options, the layout
// gallery, deck setup, SPEC-104's print and export, and — since SPEC-103 paid
// for its wiring with them — SPEC-099's Reuse slides and SPEC-093's Header &
// footer, and SPEC-115's Redesign), collected out of PresentationView so the view keeps a one-line
// hookup (PresentationView.size.spec.ts). Each is shown when its controller
// reports open; the controllers own all state and commands. Runtime prop
// declarations per the espree constraint.
import HtmlExportOptionsDialog from "./HtmlExportOptionsDialog.vue";
import LayoutGalleryDialog from "./LayoutGalleryDialog.vue";
import DeckSetupDialog from "./DeckSetupDialog.vue";
import DeckPrintDialog from "./DeckPrintDialog.vue";
import ReuseSlidesDialog from "./ReuseSlidesDialog.vue";
import HeaderFooterDialog from "./HeaderFooterDialog.vue";
import ExperimentGate from "../ExperimentGate.vue";

defineProps({
  deckActions: { type: Object, required: true },
  layouts: { type: Object, required: true },
  deckSetup: { type: Object, required: true },
  /** SPEC-104: the useDeckExport controller. */
  deckExport: { type: Object, default: null },
  /** SPEC-099: the slide sorter's Reuse slides controller (`sorter.reuse`). */
  reuse: { type: Object, default: null },
  /** SPEC-093: the useHeaderFooter controller. */
  headerFooter: { type: Object, default: null },
  slideNumber: { type: Number, default: 1 },
  /** SPEC-115: the useDeckRedesign controller (Slide › Redesign…). */
  redesign: { type: Object, default: null },
});

// SPEC-115: the deck designer in redesign mode, loaded through the
// deck-designer experiment's gate only while Redesign is open.
const loadDeckDesigner = () => import("../../views/design/DeckDesignerDialog.vue");
</script>

<template>
  <HtmlExportOptionsDialog
    v-if="deckActions.htmlDialogOpen.value"
    :media-note="deckActions.htmlMediaNote.value"
    @confirm="deckActions.confirmHtmlExport($event)"
    @cancel="deckActions.cancelHtmlExport()"
  />
  <LayoutGalleryDialog
    v-if="layouts.open.value"
    :controller="layouts"
    :slide-number="slideNumber"
  />
  <DeckSetupDialog
    v-if="deckSetup.isOpen.value"
    :controller="deckSetup"
  />
  <DeckPrintDialog
    v-if="deckExport?.isOpen.value"
    :controller="deckExport"
  />
  <ReuseSlidesDialog
    v-if="reuse?.open.value"
    :controller="reuse"
  />
  <HeaderFooterDialog
    v-if="headerFooter?.open.value"
    :controller="headerFooter"
    :slide-number="slideNumber"
  />
  <ExperimentGate
    v-if="redesign?.open.value"
    experiment="deck-designer"
    :load="loadDeckDesigner"
    mode="redesign"
    :host="redesign.host"
    @close="redesign.close()"
  />
</template>
