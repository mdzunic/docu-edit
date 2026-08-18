<script setup lang="ts">
import { computed } from "vue";
import EmbedFontsDialog from "./EmbedFontsDialog.vue";
import { asFontReportEntries, fontStatusLabel } from "../views/document/useFonts";

// The "Fonts" panel (SPEC-048 AC-7/8/10): every family the document uses,
// each marked available / substituted (naming the twin) / missing. Also hosts
// the "Embed fonts in this file" action and its dialog (AC-15/16/17/18); the
// embed controller is passed in from useFonts. Runtime prop/emit declarations,
// since espree parses .vue scripts.
const props = defineProps({
  entries: { type: Array, default: () => [] },
  controls: { type: Object, default: () => ({}) },
});

const emit = defineEmits(["close"]);

const entriesTyped = computed(() => asFontReportEntries(props.entries));
</script>

<template>
  <aside
    class="fonts-pane"
    aria-label="Fonts in this document"
    data-testid="fonts-pane"
  >
    <div class="fonts-pane-header">
      <span class="fonts-pane-title">Fonts</span>
      <button
        type="button"
        class="fonts-close"
        data-testid="fonts-pane-close"
        aria-label="Close fonts panel"
        @click="emit('close')"
      >
        ×
      </button>
    </div>

    <p
      class="fonts-summary"
      data-testid="fonts-summary"
    >
      {{ entriesTyped.length }} font{{ entriesTyped.length === 1 ? "" : "s" }} used
    </p>

    <button
      v-if="props.controls.available"
      type="button"
      class="fonts-embed-button"
      data-testid="fonts-embed"
      :disabled="props.controls.busy"
      @click="props.controls.open()"
    >
      Embed fonts in this file…
    </button>

    <ul
      class="fonts-list"
      data-testid="fonts-list"
    >
      <li
        v-for="entry in entriesTyped"
        :key="entry.family"
        class="fonts-item"
        :data-testid="`fonts-item-${entry.family}`"
      >
        <span class="fonts-family">{{ entry.family }}</span>
        <span
          class="fonts-status"
          :class="`fonts-status-${entry.status}`"
          :data-testid="`fonts-status-${entry.family}`"
        >
          {{ fontStatusLabel(entry) }}
        </span>
      </li>
    </ul>

    <EmbedFontsDialog
      v-if="props.controls.dialogOpen"
      :face-count="props.controls.faceCount"
      :size-text="props.controls.sizeText"
      :skipped="props.controls.skipped"
      :busy="props.controls.busy"
      @confirm="props.controls.confirm()"
      @close="props.controls.close()"
    />
  </aside>
</template>

<style scoped>
.fonts-pane {
  flex: none;
  align-self: flex-start;
  position: sticky;
  top: 0;
  width: 20rem;
  max-height: 100vh;
  display: flex;
  flex-direction: column;
  min-height: 0;
  overflow-y: auto;
  background: var(--surface-raised);
  border-left: var(--border-width) solid var(--border-default);
  box-sizing: border-box;
  font-size: var(--text-md);
}

.fonts-pane-header {
  flex: none;
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: var(--space-3);
  padding: var(--space-3) var(--space-4);
}

.fonts-pane-title {
  font-size: var(--text-sm);
  font-weight: var(--weight-bold);
  letter-spacing: 0.04em;
  text-transform: uppercase;
  color: var(--text-secondary);
}

.fonts-close {
  border: none;
  background: transparent;
  font-size: var(--text-lg);
  line-height: 1;
  cursor: pointer;
  color: var(--text-secondary);
}

.fonts-summary {
  flex: none;
  margin: 0;
  padding: 0 var(--space-4) var(--space-3);
  color: var(--text-secondary);
}

.fonts-embed-button {
  margin: 0 var(--space-4) var(--space-3);
  padding: var(--space-2) var(--space-3);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-sm);
  background: var(--surface-base);
  font-size: var(--text-sm);
  cursor: pointer;
}

.fonts-embed-button:disabled {
  opacity: 0.6;
  cursor: default;
}

.fonts-list {
  list-style: none;
  margin: 0;
  padding: 0 var(--space-3) var(--space-3);
  display: flex;
  flex-direction: column;
  gap: var(--space-2);
}

.fonts-item {
  display: flex;
  align-items: baseline;
  justify-content: space-between;
  gap: var(--space-3);
  background: var(--surface-base);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-md);
  padding: var(--space-3) var(--space-3);
}

.fonts-family {
  font-weight: var(--weight-bold);
  color: var(--text-primary);
}

.fonts-status {
  flex: none;
  font-size: var(--text-xs);
  white-space: nowrap;
}

.fonts-status-available {
  color: var(--success);
}

.fonts-status-substituted {
  color: var(--warning);
}

.fonts-status-missing {
  color: var(--text-danger);
}
</style>
