<script setup lang="ts">
// The presenter view's speaker notes (SPEC-100): the current slide's SPEC-083
// `notes` read-only, one paragraph per entry — editing them stays in the
// editor's notes panel — at a scale the smaller/larger controls step through
// (the view persists it per device). Runtime declarations, not
// defineProps<{...}>() generics (espree lints .vue scripts).
import UiButton from "../ui/UiButton.vue";

defineProps({
  notes: { type: Array, default: () => [] },
  /** A multiplier of the base notes size (PRESENTER_NOTES_SCALES). */
  scale: { type: Number, default: 1 },
  canShrink: { type: Boolean, default: true },
  canGrow: { type: Boolean, default: true },
});

const emit = defineEmits(["smaller", "larger"]);
</script>

<template>
  <section
    class="notes-panel"
    aria-labelledby="presenter-notes-heading"
  >
    <div class="notes-head">
      <h2
        id="presenter-notes-heading"
        class="notes-title"
      >
        Notes
      </h2>
      <UiButton
        size="sm"
        variant="ghost"
        data-testid="presenter-notes-smaller"
        aria-label="Smaller notes text"
        :disabled="!canShrink"
        @click="emit('smaller')"
      >
        A−
      </UiButton>
      <UiButton
        size="sm"
        variant="ghost"
        data-testid="presenter-notes-larger"
        aria-label="Larger notes text"
        :disabled="!canGrow"
        @click="emit('larger')"
      >
        A+
      </UiButton>
    </div>
    <div
      class="notes-body"
      data-testid="presenter-notes"
      :data-scale="scale"
      :style="{ '--notes-scale': String(scale) }"
    >
      <p
        v-for="(line, index) in notes"
        :key="index"
      >
        {{ line }}
      </p>
      <p
        v-if="notes.length === 0"
        class="notes-empty"
      >
        No notes for this slide
      </p>
    </div>
  </section>
</template>

<style scoped>
.notes-panel {
  display: flex;
  flex-direction: column;
  min-height: 0;
  gap: var(--space-2);
}

.notes-head {
  display: flex;
  align-items: center;
  gap: var(--space-1);
}

.notes-title {
  flex: 1;
  margin: 0;
  color: var(--text-secondary);
  font-size: var(--text-sm);
  font-weight: var(--weight-bold);
}

.notes-body {
  flex: 1;
  min-height: 0;
  overflow-y: auto;
  padding: var(--space-3);
  border-radius: var(--radius-md);
  background: var(--surface-sunken);
  color: var(--text-primary);
  font-size: calc(var(--text-md) * var(--notes-scale, 1));
  line-height: var(--line-normal);
}

.notes-body p {
  margin: 0 0 var(--space-2);
}

.notes-empty {
  color: var(--text-secondary);
}
</style>
