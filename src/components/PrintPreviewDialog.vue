<script setup lang="ts">
// The print preview panel (SPEC-034 AC-1..5, AC-27/28/31): a docked side
// panel over the paginated canvas — the canvas itself IS the preview (the
// body's is-print-preview class hides screen-only affordances and pages
// outside the selected range). Runtime prop/emit declarations, not the
// generic form — this script is parsed by espree (see AppShell.vue).
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
});

const emit = defineEmits(["update:range-mode", "update:range-text", "print", "close"]);

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
      This document has no content to print.
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
    <footer class="panel-actions">
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
/* Docked to the top-right of the canvas so the paginated pages stay
   visible beside it — they are the preview. */
.print-preview-panel {
  position: fixed;
  top: 6rem;
  right: 1.5rem;
  z-index: 120;
  width: 17rem;
  padding: 1rem;
  background: #fff;
  border: 1px solid #d0d0d0;
  border-radius: 8px;
  box-shadow: 0 4px 16px rgb(0 0 0 / 0.25);
  color: #202124;
  font-size: 0.875rem;
}

.panel-header {
  display: flex;
  align-items: center;
  justify-content: space-between;
}

.panel-title {
  margin: 0;
  font-size: 1rem;
  font-weight: 600;
}

.panel-close {
  border: none;
  background: none;
  font-size: 1.25rem;
  line-height: 1;
  color: #5f6368;
  cursor: pointer;
  padding: 0.125rem 0.375rem;
}

.panel-count {
  margin: 0.5rem 0 0;
  color: #5f6368;
}

.panel-empty {
  margin: 0.5rem 0 0;
  color: #b3261e;
}

.panel-range {
  margin: 0.75rem 0 0;
  padding: 0.5rem 0.75rem 0.75rem;
  border: 1px solid #e0e0e0;
  border-radius: 6px;
}

.range-option {
  display: block;
  padding: 0.125rem 0;
  cursor: pointer;
}

.range-input {
  font: inherit;
  width: 100%;
  box-sizing: border-box;
  margin-top: 0.375rem;
  padding: 0.25rem 0.5rem;
  border: 1px solid #bdbdbd;
  border-radius: 4px;
}

.range-error {
  margin: 0.375rem 0 0;
  color: #b3261e;
}

.panel-actions {
  display: flex;
  gap: 0.5rem;
  margin-top: 1rem;
}

.panel-print {
  font: inherit;
  padding: 0.375rem 1.25rem;
  border: none;
  border-radius: 4px;
  background: #1a73e8;
  color: #fff;
  cursor: pointer;
}

.panel-print:disabled {
  background: #a8c7fa;
  cursor: default;
}

.panel-cancel {
  font: inherit;
  padding: 0.375rem 1rem;
  border: 1px solid #bdbdbd;
  border-radius: 4px;
  background: #fff;
  cursor: pointer;
}
</style>
