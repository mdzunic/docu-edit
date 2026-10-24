<script setup lang="ts">
// The assembly inspector (SPEC-109): the layout controls — layout, columns
// (grid and masonry), the gap between cells (in the user's unit) and the
// cell background — then, in the overlay layout, the selected image's own
// placement (x/y as a percentage of the content box, scale, opacity, blend),
// above SPEC-108's SheetSetupPanel, whose fit and align apply within each
// cell (its estimate is hidden: the output bar states this output's own).
// For a PDF (SPEC-110) the inspector turns from one sheet to pages: the page
// controls come first, the layout controls only where a layout applies (all
// on one page; n-up is SPEC-109's grid, with its gap), and the one sheet
// setup below serves every page.
// All logic is the useAssembly controller passed in; a rejected input snaps
// back to the value in force and the reason shows below. Runtime prop
// declarations (espree lints .vue scripts).
import { computed, reactive, watchEffect } from "vue";
import SheetSetupPanel from "./SheetSetupPanel.vue";
import AssemblyPagesPanel from "./AssemblyPagesPanel.vue";
import { BLEND_MODES, LAYOUT_KINDS } from "../../lib/core/assemblyLayout";
import { SHEET_BACKGROUND_CHOICES } from "../../lib/assemblySheet";
import { COLUMN_CHOICES } from "../../views/assembly/useAssembly";

const props = defineProps({
  /** The useAssembly() controller. */
  assembly: { type: Object, required: true },
});

const a = props.assembly;

const LAYOUT_LABELS = {
  grid: "Grid",
  horizontal: "One row",
  vertical: "One column",
  masonry: "Masonry",
  overlay: "Overlay",
};

const BLEND_LABELS = {
  normal: "Normal",
  multiply: "Multiply",
  screen: "Screen",
  overlay: "Overlay",
};

const TRANSFORM_FIELDS = [
  { field: "xPercent", label: "Across (%)", testid: "overlay-x", min: 0, max: 100, step: 0.1 },
  { field: "yPercent", label: "Down (%)", testid: "overlay-y", min: 0, max: 100, step: 0.1 },
  { field: "scalePercent", label: "Scale (%)", testid: "overlay-scale", min: 1, max: 1000, step: 1 },
  { field: "opacity", label: "Opacity", testid: "overlay-opacity", min: 0, max: 100, step: 1 },
];

const drafts = reactive({ gap: "" });
watchEffect(() => {
  drafts.gap = a.gapDisplay.value;
});

function valueOf(event = new Event("change")) {
  const el = event.target;
  return el instanceof HTMLSelectElement || el instanceof HTMLInputElement ? el.value : "";
}

// A control whose input the controller refuses — or clamps — shows the value
// actually in force afterwards.
function commit(event = new Event("change"), apply = (value = "") => Boolean(value), current = () => "") {
  const el = event.target;
  if (el instanceof HTMLSelectElement || el instanceof HTMLInputElement) {
    apply(el.value);
    el.value = current();
  }
}

const cellChoice = () =>
  SHEET_BACKGROUND_CHOICES.some((choice) => choice.value === a.layout.cellBackground) ? a.layout.cellBackground : "custom";

// A PDF of one image per page, or n-up, lays out by its own grid.
const layoutApplies = computed(() => !a.isPdf.value || a.output.pageMode === "single-page");
// One per page fills the content box: there is no gap between cells.
const gapApplies = computed(() => !a.isOverlay.value && !(a.isPdf.value && a.output.pageMode === "one-per-page"));

function transformValue(field = "") {
  const entry = a.overlayItem.value;
  return entry ? String(entry.transform[field]) : "";
}
</script>

<template>
  <div
    class="assembly-inspector"
    data-testid="assembly-inspector"
  >
    <AssemblyPagesPanel
      v-if="a.isPdf.value"
      :assembly="a"
    />

    <section
      class="inspector-layout"
      aria-label="Layout"
    >
      <div
        v-if="layoutApplies"
        class="inspector-row"
      >
        <label class="inspector-field">
          <span>Layout</span>
          <select
            data-testid="assembly-layout"
            :value="a.layout.layout"
            @change="commit($event, a.setLayout, () => a.layout.layout)"
          >
            <option
              v-for="kind in LAYOUT_KINDS"
              :key="kind"
              :value="kind"
            >
              {{ LAYOUT_LABELS[kind] }}
            </option>
          </select>
        </label>
        <label
          v-if="a.layout.layout === 'grid' || a.layout.layout === 'masonry'"
          class="inspector-field"
        >
          <span>Columns</span>
          <select
            data-testid="assembly-columns"
            :value="String(a.layout.columns)"
            @change="commit($event, a.setColumns, () => String(a.layout.columns))"
          >
            <option
              v-for="choice in COLUMN_CHOICES"
              :key="choice"
              :value="String(choice)"
            >
              {{ choice === "auto" ? "Auto" : choice }}
            </option>
            <option
              v-if="!COLUMN_CHOICES.includes(a.layout.columns)"
              :value="String(a.layout.columns)"
            >
              {{ a.layout.columns }}
            </option>
          </select>
        </label>
      </div>
      <div class="inspector-row">
        <label
          v-if="gapApplies"
          class="inspector-field"
        >
          <span>Gap ({{ a.sheet.unit.value }})</span>
          <input
            v-model="drafts.gap"
            type="text"
            inputmode="decimal"
            data-testid="assembly-gap"
            @change="commit($event, a.setGap, () => a.gapDisplay.value)"
          >
        </label>
        <label class="inspector-field">
          <span>Cell background</span>
          <select
            data-testid="assembly-cell-background"
            :value="cellChoice()"
            @change="commit($event, a.setCellBackground, cellChoice)"
          >
            <option
              v-for="choice in SHEET_BACKGROUND_CHOICES"
              :key="choice.value"
              :value="choice.value"
            >
              {{ choice.label }}
            </option>
          </select>
        </label>
        <label
          v-if="cellChoice() === 'custom'"
          class="inspector-field"
        >
          <span>Colour</span>
          <input
            type="color"
            data-testid="assembly-cell-color"
            :value="a.layout.cellBackground"
            @change="a.setCellBackground(valueOf($event))"
          >
        </label>
      </div>
    </section>

    <section
      v-if="a.isOverlay.value"
      class="inspector-overlay"
      aria-label="Placement of the selected image"
      data-testid="overlay-inspector"
    >
      <template v-if="a.overlayItem.value">
        <p class="inspector-caption">
          {{ a.overlayItem.value.title }}
        </p>
        <div class="inspector-row">
          <label
            v-for="item in TRANSFORM_FIELDS"
            :key="item.field"
            class="inspector-field"
          >
            <span>{{ item.label }}</span>
            <input
              type="number"
              :data-testid="item.testid"
              :min="item.min"
              :max="item.max"
              :step="item.step"
              :value="transformValue(item.field)"
              @change="commit($event, (value) => a.setTransformField(item.field, value), () => transformValue(item.field))"
            >
          </label>
          <label class="inspector-field">
            <span>Blend</span>
            <select
              data-testid="overlay-blend"
              :value="a.overlayItem.value.transform.blend"
              @change="commit($event, (value) => a.setTransformField('blend', value), () => transformValue('blend'))"
            >
              <option
                v-for="blend in BLEND_MODES"
                :key="blend"
                :value="blend"
              >
                {{ BLEND_LABELS[blend] }}
              </option>
            </select>
          </label>
        </div>
      </template>
      <p
        v-else
        class="inspector-caption"
        data-testid="overlay-inspector-empty"
      >
        Select an image in the tray or the preview to place it.
      </p>
    </section>

    <p
      v-if="a.error.value"
      class="inspector-error"
      data-testid="assembly-error"
      role="alert"
    >
      {{ a.error.value }}
    </p>

    <SheetSetupPanel
      :controller="a.sheet"
      :estimate="false"
    />
  </div>
</template>

<style scoped>
.assembly-inspector {
  display: flex;
  flex-direction: column;
  gap: var(--space-4);
  font-size: var(--text-sm);
  color: var(--text-primary);
}

.inspector-layout,
.inspector-overlay {
  display: flex;
  flex-direction: column;
  gap: var(--space-3);
}

.inspector-overlay {
  padding: var(--space-2) var(--space-3) var(--space-3);
  border: var(--border-width) solid var(--border-subtle);
  border-radius: var(--radius-md);
}

.inspector-row {
  display: flex;
  flex-wrap: wrap;
  align-items: flex-end;
  gap: var(--space-3);
}

.inspector-field {
  display: flex;
  flex-direction: column;
  gap: var(--space-1);
}

.inspector-field span,
.inspector-caption {
  color: var(--text-secondary);
}

.inspector-caption,
.inspector-error {
  margin: 0;
}

.inspector-field select,
.inspector-field input {
  font: inherit;
  min-height: var(--height-control);
  padding: 0 var(--space-2);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-sm);
  background: var(--surface-base);
  color: var(--text-primary);
}

.inspector-field input[type="text"],
.inspector-field input[type="number"] {
  width: 5rem;
}

.inspector-field select:focus-visible,
.inspector-field input:focus-visible {
  outline: var(--space-1) solid var(--focus-ring);
  outline-offset: 0;
}

.inspector-error {
  color: var(--text-danger);
}
</style>
