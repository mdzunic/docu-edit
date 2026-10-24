<script setup lang="ts">
// SPEC-115: a deck menu's "Redesign…" (deck-redesign), badged as
// experimental. Loaded only by the slides menu bar's ExperimentGate while the
// deck-designer experiment is on. It only emits `select`: the menu routes that
// to the view's redesign controller, and the dialog itself is hosted by
// PresentationDialogs — a menu popup unmounts as it closes, so it cannot host
// one. Disabled on a deck with no slides (there is nothing to arrange).
// Runtime prop declarations per the espree constraint (see AppShell.vue).
import ExperimentBadge from "../../components/ExperimentBadge.vue";
import MenuItem from "../../components/MenuItem.vue";

defineProps({
  disabled: { type: Boolean, default: false },
});

const emit = defineEmits(["select"]);
</script>

<template>
  <div class="redesign-item">
    <MenuItem
      label="Redesign…"
      testid="deck-redesign"
      :disabled="disabled"
      @select="emit('select')"
    />
    <ExperimentBadge experiment="deck-designer" />
  </div>
</template>

<style scoped>
.redesign-item {
  display: flex;
  align-items: center;
  gap: var(--space-2);
  padding-right: var(--space-2);
}

.redesign-item > :first-child {
  flex: 1 1 auto;
}
</style>
