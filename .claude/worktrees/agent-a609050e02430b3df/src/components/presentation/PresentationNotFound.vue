<script setup lang="ts">
// The presentation route's "not found" state (extracted from PresentationView,
// SPEC-099, to pay for the slide sorter's wiring): the deck id names nothing in
// the library, so the view offers the way back. Emits `back`; the view owns
// the routing.
import UiButton from "../ui/UiButton.vue";

const emit = defineEmits(["back"]);
</script>

<template>
  <div
    class="state-panel"
    data-testid="presentation-not-found"
  >
    <h2>Presentation not found</h2>
    <p>This link doesn't match a presentation in your library.</p>
    <UiButton
      data-testid="presentation-back"
      @click="emit('back')"
    >
      Back to library
    </UiButton>
  </div>
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
</style>
