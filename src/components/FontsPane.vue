<script setup lang="ts">
import { computed } from "vue";
import { asFontReportEntries, fontStatusLabel } from "../views/document/useFonts";

// The "Fonts" panel (SPEC-048 AC-7/8/10): every family the document uses,
// each marked available / substituted (naming the twin) / missing. Pure
// disclosure — docked the same way AccessibilityPane is. Runtime prop/emit
// declarations, not the generic form, since espree parses .vue scripts.
const props = defineProps({
  entries: { type: Array, default: () => [] },
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
  background: #f8f9fa;
  border-left: 1px solid #dadce0;
  box-sizing: border-box;
  font-size: 0.8125rem;
}

.fonts-pane-header {
  flex: none;
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 0.5rem;
  padding: 0.5rem 0.75rem;
}

.fonts-pane-title {
  font-size: 0.75rem;
  font-weight: 600;
  letter-spacing: 0.04em;
  text-transform: uppercase;
  color: #5f6368;
}

.fonts-close {
  border: none;
  background: transparent;
  font-size: 1rem;
  line-height: 1;
  cursor: pointer;
  color: #5f6368;
}

.fonts-summary {
  flex: none;
  margin: 0;
  padding: 0 0.75rem 0.5rem;
  color: #5f6368;
}

.fonts-list {
  list-style: none;
  margin: 0;
  padding: 0 0.5rem 0.5rem;
  display: flex;
  flex-direction: column;
  gap: 0.35rem;
}

.fonts-item {
  display: flex;
  align-items: baseline;
  justify-content: space-between;
  gap: 0.5rem;
  background: #fff;
  border: 1px solid #dadce0;
  border-radius: 6px;
  padding: 0.4rem 0.6rem;
}

.fonts-family {
  font-weight: 600;
  color: #202124;
}

.fonts-status {
  flex: none;
  font-size: 0.72rem;
  white-space: nowrap;
}

.fonts-status-available {
  color: #188038;
}

.fonts-status-substituted {
  color: #b06000;
}

.fonts-status-missing {
  color: #c5221f;
}
</style>
