<script setup lang="ts">
// The presentation view's modal dialogs (HTML export options, the layout
// gallery, deck setup), collected out of PresentationView so the view keeps a
// one-line hookup (PresentationView.size.spec.ts). Each is shown when its
// controller reports open; the controllers own all state and commands. Runtime
// prop declarations per the espree constraint.
import HtmlExportOptionsDialog from "./HtmlExportOptionsDialog.vue";
import LayoutGalleryDialog from "./LayoutGalleryDialog.vue";
import DeckSetupDialog from "./DeckSetupDialog.vue";

defineProps({
  deckActions: { type: Object, required: true },
  layouts: { type: Object, required: true },
  deckSetup: { type: Object, required: true },
  slideNumber: { type: Number, default: 1 },
});
</script>

<template>
  <HtmlExportOptionsDialog
    v-if="deckActions.htmlDialogOpen.value"
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
</template>
