<script setup lang="ts">
// SPEC-115: the library New menu's "Presentation from a description…"
// (new-deck-from-description) — the deck designer's entry point that creates
// a new deck and changes nothing that exists. Loaded only by
// LibraryHeaderActions' ExperimentGate while the deck-designer experiment is
// on, so with the flag off neither this nor anything it imports is fetched.
//
// The dialog is teleported to <body> because this button sits in the
// library's sticky rail, whose stacking context would put the dialog under the
// sticky list header and quick look. Its layer sits one below the app's other
// dialogs (z-index 50), so SPEC-113's consent dialog — which a request opens
// on top of this one — and Labs (from the badge) still show above it.
import { ref } from "vue";
import ExperimentBadge from "../../components/ExperimentBadge.vue";
import UiButton from "../../components/ui/UiButton.vue";
import DeckDesignerDialog from "./DeckDesignerDialog.vue";

const open = ref(false);
</script>

<template>
  <span class="new-deck-designer">
    <UiButton
      data-testid="new-deck-from-description"
      @click="open = true"
    >
      Presentation from a description…
    </UiButton>
    <ExperimentBadge experiment="deck-designer" />
    <Teleport to="body">
      <div
        v-if="open"
        class="designer-layer"
      >
        <DeckDesignerDialog
          mode="new"
          @close="open = false"
        />
      </div>
    </Teleport>
  </span>
</template>

<style scoped>
.new-deck-designer {
  display: inline-flex;
  align-items: center;
  gap: var(--space-2);
}

.designer-layer {
  position: relative;
  z-index: 49;
}
</style>
