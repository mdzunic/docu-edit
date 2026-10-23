<script setup lang="ts">
// The assembly sheet's setup panel (SPEC-108) — the one control every
// assembly surface embeds: paper (SPEC-027's catalogue, Custom, or Fit to
// content), orientation, Custom width/height, margins (SPEC-066's presets
// plus four per-edge fields and a link toggle), fit, alignment, background,
// resolution, the live output estimate and the rejected-input reason. All
// logic lives in the useSheetSetup controller passed in; this component is
// presentation + event forwarding only, in the TransitionsPanel.vue mould.
// Every control is native, so the keyboard works without help. Lengths are
// shown and typed in the user's measurement unit; the controller keeps
// centimetres. Runtime prop declarations (espree lints .vue scripts).
import { reactive, watchEffect } from "vue";
import { SHEET_ALIGN_LABELS, SHEET_BACKGROUND_CHOICES, SHEET_FIT_LABELS } from "../../composables/useSheetSetup";

const props = defineProps({
  /** The useSheetSetup() controller object (refs + command functions). */
  controller: { type: Object, required: true },
  /** Whether to show the output estimate — off where the surface states
   * its own (SPEC-109's output bar, whose format can be JPEG). */
  estimate: { type: Boolean, default: true },
});

const c = props.controller;

const EDGES = [
  { edge: "top", key: "topCm", label: "Top" },
  { edge: "right", key: "rightCm", label: "Right" },
  { edge: "bottom", key: "bottomCm", label: "Bottom" },
  { edge: "left", key: "leftCm", label: "Left" },
];

// The text fields edit drafts, re-read from the setup whenever it (or the
// unit) changes. A rejected entry leaves the setup untouched, so the typed
// text stays beside its reason until the user fixes it — or until another
// accepted change re-reads every field from the setup.
const drafts = reactive(Object.fromEntries(["width", "height", ...EDGES.map((item) => item.edge)].map((name) => [name, ""])));
watchEffect(() => {
  const setup = c.setup.value;
  const sheet = c.geometry.value;
  drafts.width = sheet ? c.display(sheet.widthCm) : "";
  drafts.height = sheet ? c.display(sheet.heightCm) : "";
  for (const { edge, key } of EDGES) {
    drafts[edge] = c.display(setup.margins[key]);
  }
});

function selectValue(event = new Event("change")) {
  const el = event.target;
  return el instanceof HTMLSelectElement || el instanceof HTMLInputElement ? el.value : "";
}

// A select whose choice the controller refuses snaps back to the setup's
// value, so the control never shows a sheet that is not the one in force.
function choose(event = new Event("change"), apply = (value = "") => Boolean(value), current = () => "") {
  const el = event.target;
  if (el instanceof HTMLSelectElement && !apply(el.value)) {
    el.value = current();
  }
}

function onLink(event = new Event("change")) {
  const el = event.target;
  if (el instanceof HTMLInputElement && !c.setMarginsLinked(el.checked)) {
    el.checked = c.marginsLinked.value;
  }
}

function invalid(field = "") {
  return c.errorField.value === field ? "true" : undefined;
}
</script>

<template>
  <section
    class="sheet-setup"
    data-testid="sheet-setup"
    aria-label="Sheet setup"
  >
    <div class="sheet-row">
      <label class="sheet-field">
        <span>Paper</span>
        <select
          data-testid="sheet-paper"
          :value="c.setup.value.paper"
          :aria-invalid="invalid('paper')"
          @change="choose($event, c.setPaper, () => c.setup.value.paper)"
        >
          <option
            v-for="paper in c.paperOptions"
            :key="paper"
            :value="paper"
          >
            {{ paper }}
          </option>
        </select>
      </label>
      <div
        class="sheet-orientation"
        role="group"
        aria-label="Orientation"
      >
        <button
          type="button"
          class="sheet-toggle"
          data-testid="sheet-portrait"
          :aria-pressed="c.setup.value.orientation === 'portrait' ? 'true' : 'false'"
          :disabled="c.isFitToContent.value"
          @click="c.setOrientation('portrait')"
        >
          Portrait
        </button>
        <button
          type="button"
          class="sheet-toggle"
          data-testid="sheet-landscape"
          :aria-pressed="c.setup.value.orientation === 'landscape' ? 'true' : 'false'"
          :disabled="c.isFitToContent.value"
          @click="c.setOrientation('landscape')"
        >
          Landscape
        </button>
      </div>
    </div>

    <div class="sheet-row">
      <label class="sheet-field">
        <span>Width ({{ c.unit.value }})</span>
        <input
          v-model="drafts.width"
          type="text"
          inputmode="decimal"
          data-testid="sheet-width"
          :disabled="!c.isCustom.value"
          :aria-invalid="invalid('width')"
          @change="c.setCustomSize('width', drafts.width)"
        >
      </label>
      <label class="sheet-field">
        <span>Height ({{ c.unit.value }})</span>
        <input
          v-model="drafts.height"
          type="text"
          inputmode="decimal"
          data-testid="sheet-height"
          :disabled="!c.isCustom.value"
          :aria-invalid="invalid('height')"
          @change="c.setCustomSize('height', drafts.height)"
        >
      </label>
    </div>

    <fieldset class="sheet-margins">
      <legend>Margins ({{ c.unit.value }})</legend>
      <div class="sheet-row">
        <label class="sheet-field">
          <span>Preset</span>
          <select
            data-testid="sheet-margin-preset"
            :value="c.marginPreset.value"
            :aria-invalid="invalid('margin-preset')"
            @change="choose($event, c.setMarginPreset, () => c.marginPreset.value)"
          >
            <option
              v-for="name in c.marginPresetOptions"
              :key="name"
              :value="name"
              :disabled="name === 'Custom'"
            >
              {{ name }}
            </option>
          </select>
        </label>
        <label class="sheet-check">
          <input
            type="checkbox"
            data-testid="sheet-margin-link"
            :checked="c.marginsLinked.value"
            :aria-invalid="invalid('margin-link')"
            @change="onLink"
          >
          <span>Same on all sides</span>
        </label>
      </div>
      <div class="sheet-row sheet-edges">
        <label
          v-for="item in EDGES"
          :key="item.edge"
          class="sheet-field"
        >
          <span>{{ item.label }}</span>
          <input
            v-model="drafts[item.edge]"
            type="text"
            inputmode="decimal"
            :data-testid="`sheet-margin-${item.edge}`"
            :aria-invalid="invalid(`margin-${item.edge}`)"
            @change="c.setMargin(item.edge, drafts[item.edge])"
          >
        </label>
      </div>
    </fieldset>

    <div class="sheet-row">
      <label class="sheet-field">
        <span>Fit</span>
        <select
          data-testid="sheet-fit"
          :value="c.setup.value.fit"
          @change="choose($event, c.setFit, () => c.setup.value.fit)"
        >
          <option
            v-for="fit in c.fitOptions"
            :key="fit"
            :value="fit"
          >
            {{ SHEET_FIT_LABELS[fit] }}
          </option>
        </select>
      </label>
      <label class="sheet-field">
        <span>Align</span>
        <select
          data-testid="sheet-align"
          :value="c.setup.value.align"
          :disabled="c.setup.value.fit === 'stretch'"
          @change="choose($event, c.setAlign, () => c.setup.value.align)"
        >
          <option
            v-for="align in c.alignOptions"
            :key="align"
            :value="align"
          >
            {{ SHEET_ALIGN_LABELS[align] }}
          </option>
        </select>
      </label>
    </div>

    <div class="sheet-row">
      <label class="sheet-field">
        <span>Background</span>
        <select
          data-testid="sheet-background"
          :value="c.backgroundChoice.value"
          :aria-invalid="invalid('background')"
          @change="choose($event, c.setBackground, () => c.backgroundChoice.value)"
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
        v-if="c.backgroundChoice.value === 'custom'"
        class="sheet-field"
      >
        <span>Colour</span>
        <input
          type="color"
          data-testid="sheet-background-color"
          :value="c.setup.value.background"
          @change="c.setBackground(selectValue($event))"
        >
      </label>
      <label class="sheet-field">
        <span>Resolution</span>
        <select
          data-testid="sheet-dpi"
          :value="String(c.setup.value.dpi)"
          @change="choose($event, c.setDpi, () => String(c.setup.value.dpi))"
        >
          <option
            v-for="dpi in c.dpiChoices"
            :key="dpi"
            :value="String(dpi)"
          >
            {{ dpi }} dpi
          </option>
        </select>
      </label>
    </div>

    <p
      v-if="c.backgroundNotice.value"
      class="sheet-notice"
      data-testid="sheet-background-notice"
      role="status"
    >
      {{ c.backgroundNotice.value }}
    </p>
    <p
      v-if="props.estimate"
      class="sheet-estimate"
      data-testid="sheet-estimate"
      aria-live="polite"
    >
      Estimated output: {{ c.estimate.value }}
    </p>
    <p
      v-if="c.error.value"
      class="sheet-error"
      data-testid="sheet-error"
      role="alert"
    >
      {{ c.error.value }}
    </p>
  </section>
</template>

<style scoped>
.sheet-setup {
  display: flex;
  flex-direction: column;
  gap: var(--space-3);
  font-size: var(--text-sm);
  color: var(--text-primary);
}

.sheet-row {
  display: flex;
  flex-wrap: wrap;
  align-items: flex-end;
  gap: var(--space-3);
}

.sheet-field {
  display: flex;
  flex-direction: column;
  gap: var(--space-1);
}

.sheet-field span,
.sheet-margins legend {
  color: var(--text-secondary);
}

.sheet-field input[type="text"] {
  width: 5rem;
}

.sheet-field select,
.sheet-field input {
  font: inherit;
  min-height: var(--height-control);
  padding: 0 var(--space-2);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-sm);
  background: var(--surface-base);
  color: var(--text-primary);
}

.sheet-field input:disabled,
.sheet-field select:disabled {
  color: var(--text-disabled);
  background: var(--surface-sunken);
}

.sheet-field [aria-invalid="true"] {
  border-color: var(--danger);
}

.sheet-check {
  display: flex;
  align-items: center;
  gap: var(--space-2);
  min-height: var(--height-control);
}

.sheet-margins {
  margin: 0;
  padding: var(--space-2) var(--space-3) var(--space-3);
  border: var(--border-width) solid var(--border-subtle);
  border-radius: var(--radius-md);
  display: flex;
  flex-direction: column;
  gap: var(--space-2);
}

.sheet-orientation {
  display: flex;
}

.sheet-toggle {
  font: inherit;
  min-height: var(--height-control);
  padding: 0 var(--space-3);
  border: var(--border-width) solid var(--border-default);
  background: var(--surface-base);
  color: var(--text-primary);
  cursor: pointer;
}

.sheet-toggle:first-child {
  border-radius: var(--radius-sm) 0 0 var(--radius-sm);
}

.sheet-toggle:last-child {
  border-radius: 0 var(--radius-sm) var(--radius-sm) 0;
}

.sheet-toggle[aria-pressed="true"] {
  background: var(--surface-selected);
  color: var(--text-accent);
  font-weight: var(--weight-medium);
}

.sheet-toggle:disabled {
  color: var(--text-disabled);
  cursor: default;
}

.sheet-toggle:focus-visible,
.sheet-field select:focus-visible,
.sheet-field input:focus-visible {
  outline: var(--space-1) solid var(--focus-ring);
  outline-offset: 0;
}

.sheet-notice,
.sheet-estimate,
.sheet-error {
  margin: 0;
}

.sheet-notice {
  color: var(--text-secondary);
  font-weight: var(--weight-medium);
}

.sheet-estimate {
  color: var(--text-secondary);
}

.sheet-error {
  color: var(--text-danger);
}
</style>
