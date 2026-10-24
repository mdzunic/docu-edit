<script setup lang="ts">
// The PDF's page controls (SPEC-110), at the top of the inspector when the
// output is PDF: how the tray is cut into pages — one per page (the
// default), several per page in SPEC-109's grid (n-up: 2, 4, 6, 9 or 16), or
// every image on one sheet (single-page, SPEC-109's own layouts) — whether
// each page carries its number, and how many pages that makes. The sheet
// panel below sets up EVERY page: one sheet per assembly, never per page.
// All logic is the useAssembly controller's `pdf`. Runtime prop
// declarations (espree lints .vue scripts).
import { N_UP_CHOICES, PAGE_MODES } from "../../lib/media/imagesToPdf";

const props = defineProps({
  /** The useAssembly() controller. */
  assembly: { type: Object, required: true },
});

const a = props.assembly;
const pdf = a.pdf;

const PAGE_MODE_LABELS = {
  "one-per-page": "One image per page",
  "n-up": "Several per page",
  "single-page": "All on one page",
};

// A control whose input the controller refuses shows the value in force.
function commit(event = new Event("change"), apply = (value = "") => Boolean(value), current = () => "") {
  const el = event.target;
  if (el instanceof HTMLSelectElement) {
    apply(el.value);
    el.value = current();
  }
}

function checkedOf(event = new Event("change")) {
  return event.target instanceof HTMLInputElement && event.target.checked;
}
</script>

<template>
  <section
    class="assembly-pages-panel"
    aria-label="Pages"
    data-testid="assembly-pages-panel"
  >
    <div class="pages-row">
      <label class="pages-field">
        <span>Pages</span>
        <select
          data-testid="assembly-page-mode"
          :value="a.output.pageMode"
          @change="commit($event, pdf.setPageMode, () => a.output.pageMode)"
        >
          <option
            v-for="mode in PAGE_MODES"
            :key="mode"
            :value="mode"
          >
            {{ PAGE_MODE_LABELS[mode] }}
          </option>
        </select>
      </label>
      <label
        v-if="a.output.pageMode === 'n-up'"
        class="pages-field"
      >
        <span>Images per page</span>
        <select
          data-testid="assembly-n-up"
          :value="String(a.output.nUp)"
          @change="commit($event, pdf.setNUp, () => String(a.output.nUp))"
        >
          <option
            v-for="n in N_UP_CHOICES"
            :key="n"
            :value="String(n)"
          >
            {{ n }}
          </option>
        </select>
      </label>
      <p
        class="pages-count"
        data-testid="assembly-page-count"
        aria-live="polite"
      >
        {{ pdf.pageCountLabel.value }}
      </p>
    </div>
    <label class="pages-check">
      <input
        type="checkbox"
        data-testid="assembly-page-numbers"
        :checked="a.output.pageNumbers"
        @change="pdf.setPageNumbers(checkedOf($event))"
      >
      <span>Page numbers</span>
    </label>
    <p class="pages-note">
      The sheet below applies to every page.
    </p>
    <p
      v-if="pdf.error.value"
      class="pages-error"
      data-testid="assembly-pages-error"
      role="alert"
    >
      {{ pdf.error.value }}
    </p>
  </section>
</template>

<style scoped>
.assembly-pages-panel {
  display: flex;
  flex-direction: column;
  gap: var(--space-2);
  padding: var(--space-2) var(--space-3) var(--space-3);
  border: var(--border-width) solid var(--border-subtle);
  border-radius: var(--radius-md);
}

.pages-row {
  display: flex;
  flex-wrap: wrap;
  align-items: flex-end;
  gap: var(--space-3);
}

.pages-field {
  display: flex;
  flex-direction: column;
  gap: var(--space-1);
}

.pages-field span,
.pages-note {
  color: var(--text-secondary);
}

.pages-field select {
  font: inherit;
  min-height: var(--height-control);
  padding: 0 var(--space-2);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-sm);
  background: var(--surface-base);
  color: var(--text-primary);
}

.pages-field select:focus-visible,
.pages-check input:focus-visible {
  outline: var(--space-1) solid var(--focus-ring);
  outline-offset: 0;
}

.pages-check {
  display: flex;
  align-items: center;
  gap: var(--space-2);
}

.pages-count,
.pages-note,
.pages-error {
  margin: 0;
}

.pages-count {
  align-self: center;
  font-weight: var(--weight-medium);
}

.pages-error {
  color: var(--text-danger);
}
</style>
