<script setup lang="ts">
import { computed, useId } from "vue";
import { EXPERIMENTS } from "../lib/experiments";
import { usePreferencesDialog } from "../composables/usePreferencesDialog";

// SPEC-113: the "Experimental" chip every enabled experimental surface
// carries, so a user always knows which part of the product is provisional.
// Hover or focus shows a tooltip saying what that means, with a link to the
// Labs section where the experiment is switched off. Runtime prop
// declarations for the espree constraint AppShell.vue explains.
const props = defineProps({
  /** The experiment this surface belongs to — named in the tooltip. Optional
   * for a surface shared by several experiments. */
  experiment: { type: String, default: "" },
});

const tooltipId = useId();
const preferencesDialog = usePreferencesDialog();

const title = computed(() => EXPERIMENTS.find((entry) => entry.id === props.experiment)?.title ?? "");
</script>

<template>
  <span class="experiment-badge-wrap">
    <span
      class="experiment-badge"
      data-testid="experiment-badge"
      tabindex="0"
      :aria-describedby="tooltipId"
    >Experimental</span>
    <span
      :id="tooltipId"
      class="experiment-badge-tooltip"
      role="tooltip"
      data-testid="experiment-badge-tooltip"
    >
      {{ title ? `${title} is experimental` : "This is experimental" }}: it may change or be removed, and is
      not covered by this product's round-trip guarantees.
      <button
        type="button"
        class="experiment-badge-link"
        data-testid="experiment-badge-labs"
        @click="preferencesDialog.openLabs()"
      >
        Manage it in Labs
      </button>
    </span>
  </span>
</template>

<style scoped>
.experiment-badge-wrap {
  position: relative;
  display: inline-flex;
  vertical-align: middle;
}

.experiment-badge {
  display: inline-flex;
  align-items: center;
  padding: 0 var(--space-2);
  border: var(--border-width) solid var(--accent);
  border-radius: var(--radius-sm);
  background: var(--accent-subtle);
  color: var(--text-accent);
  font-family: var(--font-ui);
  font-size: var(--text-xs);
  font-weight: var(--weight-bold);
  letter-spacing: var(--tracking-caption);
  line-height: var(--line-tight);
  text-transform: uppercase;
  white-space: nowrap;
  cursor: default;
}

/* Shown while the chip or anything inside the tooltip is hovered or
   focused, so the link inside stays reachable by pointer and keyboard. */
.experiment-badge-tooltip {
  position: absolute;
  top: 100%;
  left: 0;
  z-index: 60;
  min-width: 16rem;
  /* Flush under the chip: a gap would drop :hover on the way to the link. */
  padding: var(--space-3);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-sm);
  background: var(--surface-raised);
  box-shadow: var(--shadow-2);
  color: var(--text-primary);
  font-family: var(--font-ui);
  font-size: var(--text-sm);
  font-weight: var(--weight-regular);
  line-height: var(--line-normal);
  text-transform: none;
  white-space: normal;
  visibility: hidden;
}

.experiment-badge-wrap:hover .experiment-badge-tooltip,
.experiment-badge-wrap:focus-within .experiment-badge-tooltip {
  visibility: visible;
}

.experiment-badge-link {
  display: block;
  margin-top: var(--space-2);
  padding: 0;
  border: none;
  background: none;
  color: var(--text-accent);
  font: inherit;
  text-decoration: underline;
  cursor: pointer;
}
</style>
