<script setup lang="ts">
// The print preview panel (SPEC-034 AC-1..5, AC-27/28/31): a docked side
// panel over the paginated canvas — the canvas itself IS the preview (the
// body's is-print-preview class hides screen-only affordances and pages
// outside the selected range). Deliberately NOT built on BaseDialog: that
// shell is a centred modal with a dimming backdrop and aria-modal="true",
// all three of which would be wrong here — the panel docks top-right, must
// leave the pages behind it visible and interactive-looking, and is not
// modal (so the missing aria-modal is correct, not an oversight). Runtime
// prop/emit declarations, not the generic form — this script is parsed by
// espree (see AppShell.vue).
import { nextTick, onMounted, useTemplateRef } from "vue";

defineProps({
  /** Printable pages; 0 for a content-empty document (AC-27). */
  pageCount: { type: Number, required: true },
  /** "all" or "custom". */
  rangeMode: { type: String, required: true },
  rangeText: { type: String, required: true },
  /** The current range validation error, or "" (AC-28). */
  error: { type: String, default: "" },
  canPrint: { type: Boolean, required: true },
  /** "Print background colours" (SPEC-074 AC-25). */
  printBackground: { type: Boolean, default: false },
  /** SPEC-136: what an empty job says (the sheet's "This sheet has nothing to print."). */
  emptyMessage: { type: String, default: "This document has no content to print." },
  /** SPEC-136: the background-colours option (a sheet prints its fills always). */
  showBackgroundOption: { type: Boolean, default: true },
});

const emit = defineEmits([
  "update:range-mode",
  "update:range-text",
  "update:print-background",
  "print",
  "close",
]);

const printButton = useTemplateRef("printButton");

onMounted(async () => {
  await nextTick();
  printButton.value?.focus();
});

function chooseCustom() {
  emit("update:range-mode", "custom");
}

function handleRangeInput(event = new Event("input")) {
  const target = event.target;
  emit("update:range-text", target instanceof HTMLInputElement ? target.value : "");
}

function handleBackgroundToggle(event = new Event("change")) {
  const target = event.target;
  emit("update:print-background", target instanceof HTMLInputElement && target.checked);
}
</script>

<template>
  <div
    class="print-preview-panel"
    role="dialog"
    aria-label="Print preview"
    data-testid="print-preview"
    @keydown.escape.stop="emit('close')"
  >
    <header class="panel-header">
      <h2 class="panel-title">
        Print preview
      </h2>
      <button
        type="button"
        class="panel-close"
        aria-label="Close print preview"
        data-testid="print-preview-close"
        @click="emit('close')"
      >
        ×
      </button>
    </header>
    <p
      class="panel-count"
      data-testid="print-page-count"
    >
      {{ pageCount }} {{ pageCount === 1 ? "page" : "pages" }}
    </p>
    <p
      v-if="pageCount === 0"
      class="panel-empty"
      data-testid="print-no-pages"
    >
      {{ emptyMessage }}
    </p>
    <fieldset
      class="panel-range"
      :disabled="pageCount === 0"
    >
      <legend>Pages</legend>
      <label class="range-option">
        <input
          type="radio"
          name="print-range"
          value="all"
          :checked="rangeMode === 'all'"
          data-testid="print-range-all"
          @change="emit('update:range-mode', 'all')"
        >
        All pages
      </label>
      <label class="range-option">
        <input
          type="radio"
          name="print-range"
          value="custom"
          :checked="rangeMode === 'custom'"
          data-testid="print-range-custom"
          @change="chooseCustom"
        >
        Custom
      </label>
      <input
        v-if="rangeMode === 'custom'"
        type="text"
        class="range-input"
        aria-label="Page range"
        placeholder="e.g. 1-3, 5"
        data-testid="print-range-input"
        :value="rangeText"
        @input="handleRangeInput"
      >
      <p
        v-if="error"
        class="range-error"
        role="alert"
        data-testid="print-range-error"
      >
        {{ error }}
      </p>
    </fieldset>
    <!-- SPEC-136: a product's own options (the sheet's print scope). -->
    <slot name="options" />
    <!-- A page background is a screen decoration until asked for on paper
         (SPEC-074 AC-25); page borders always print, so they get no switch. -->
    <label
      v-if="showBackgroundOption"
      class="panel-option"
    >
      <input
        type="checkbox"
        :checked="printBackground"
        data-testid="print-background"
        @change="handleBackgroundToggle"
      >
      Print background colours
    </label>
    <footer class="panel-actions">
      <!-- SPEC-136: a product's own actions (the sheet's Page setup…, Download PDF). -->
      <slot name="actions" />
      <button
        ref="printButton"
        type="button"
        class="panel-print"
        data-testid="print-confirm"
        :disabled="!canPrint"
        @click="emit('print')"
      >
        Print
      </button>
      <button
        type="button"
        class="panel-cancel"
        data-testid="print-cancel"
        @click="emit('close')"
      >
        Cancel
      </button>
    </footer>
  </div>
</template>

<style scoped>
.panel-option {
  display: flex;
  align-items: center;
  gap: var(--space-3);
  margin: var(--space-4) 0 0;
  font-size: var(--text-md);
}

/* Docked to the top-right of the canvas so the paginated pages stay
   visible beside it — they are the preview. */
.print-preview-panel {
  position: fixed;
  top: 6rem;
  right: 1.5rem;
  z-index: 120;
  width: 17rem;
  padding: var(--space-5);
  background: var(--surface-base);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-md);
  box-shadow: var(--shadow-3);
  color: var(--text-primary);
  font-size: var(--text-md);
}

.panel-header {
  display: flex;
  align-items: center;
  justify-content: space-between;
}

.panel-title {
  margin: 0;
  font-size: var(--text-lg);
  font-weight: var(--weight-bold);
}

.panel-close {
  border: none;
  background: none;
  font-size: var(--text-xl);
  line-height: 1;
  color: var(--text-secondary);
  cursor: pointer;
  padding: var(--space-1) var(--space-3);
}

.panel-count {
  margin: var(--space-3) 0 0;
  color: var(--text-secondary);
}

.panel-empty {
  margin: var(--space-3) 0 0;
  color: var(--text-danger);
}

.panel-range {
  margin: var(--space-4) 0 0;
  padding: var(--space-3) var(--space-4) var(--space-4);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-md);
}

.range-option {
  display: block;
  padding: var(--space-1) 0;
  cursor: pointer;
}

.range-input {
  font: inherit;
  width: 100%;
  box-sizing: border-box;
  margin-top: var(--space-3);
  padding: var(--space-2) var(--space-3);
  border: var(--border-width) solid var(--border-strong);
  border-radius: var(--radius-sm);
}

.range-error {
  margin: var(--space-3) 0 0;
  color: var(--text-danger);
}

.panel-actions {
  display: flex;
  flex-wrap: wrap;
  gap: var(--space-3);
  margin-top: var(--space-5);
}

.panel-print {
  font: inherit;
  padding: var(--space-3) var(--space-6);
  border: none;
  border-radius: var(--radius-sm);
  background: var(--accent);
  color: var(--text-on-brand);
  cursor: pointer;
}

.panel-print:disabled {
  background: var(--accent-subtle);
  cursor: default;
}

.panel-cancel {
  font: inherit;
  padding: var(--space-3) var(--space-5);
  border: var(--border-width) solid var(--border-strong);
  border-radius: var(--radius-sm);
  background: var(--surface-base);
  cursor: pointer;
}
</style>
