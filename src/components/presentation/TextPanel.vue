<script setup lang="ts">
// The Text panel (SPEC-097 AC-34..36): paragraph spacing (line spacing in
// CustomSpacingDialog.vue's multiple / exactly / at least vocabulary, rebuilt
// as native controls because that dialog is bound to the text model; space
// before/after; letter spacing) and the text body of the selected or edited
// text/shape frame (autofit, anchor, insets, columns). Every control reflects
// the selection — a mixed value shows blank — and commits on `change`, one
// undo entry each, through the useSlideTextDepth controller passed in. The
// slide link dialog shares that controller and mounts here, beside the panel
// rather than inside a ribbon strip (a strip may be hidden by its tab).
//
// A dumb :controller-taking component like ThemesPanel.vue; runtime prop
// declarations and untyped helpers (espree) — the option lists and event
// readers live in slideTextControls.ts.
import { computed } from "vue";
import SlideLinkDialog from "./SlideLinkDialog.vue";
import {
  ANCHOR_OPTIONS,
  AUTOFIT_OPTIONS,
  INSET_SIDES,
  LINE_SPACING_MODE_OPTIONS,
  eventNumber,
  eventString,
  inputValue,
} from "./slideTextControls";

const props = defineProps({
  /** The useSlideTextDepth() controller object (refs + command functions). */
  controller: { type: Object, required: true },
  /** The 1-based number of the current slide, for the panel's accessible name. */
  slideNumber: { type: Number, default: 1 },
});

const c = props.controller;

const isMultiple = computed(() => c.lineSpacingMode.value !== "exact" && c.lineSpacingMode.value !== "atLeast");
const insets = computed(() => c.insetsCm.value ?? {});

function onKeydown(event = new KeyboardEvent("keydown")) {
  // The stage's window keys (arrow nudges, Escape-deselect) must not also act
  // on keys typed into the panel's controls.
  event.stopPropagation();
  // Enter commits a number field (its change event, fired by the blur). Its
  // default action is cancelled first: the commit re-renders and refocuses an
  // open text session's contenteditable host, and the browser would otherwise
  // deliver this Enter there and split the paragraph at the caret.
  const target = event.target;
  if (event.key === "Enter" && target instanceof HTMLInputElement) {
    event.preventDefault();
    target.blur();
  }
}
function onMode(event = new Event("change")) {
  const mode = eventString(event);
  if (mode !== "") c.setLineSpacingMode(mode);
}
function onNumber(command = "", event = new Event("change")) {
  const value = eventNumber(event);
  if (!Number.isFinite(value)) return;
  c[command](value);
}
function onInset(side = "left", event = new Event("change")) {
  const value = eventNumber(event);
  if (Number.isFinite(value)) c.setInset(side, value);
}
function onSelect(command = "", event = new Event("change")) {
  const value = eventString(event);
  if (value !== "") c[command](value);
}
</script>

<template>
  <section
    v-if="c.panelOpen.value"
    class="text-panel"
    data-testid="slides-text-panel"
    :aria-label="`Text for slide ${slideNumber}`"
    @keydown="onKeydown"
  >
    <header class="panel-header">
      <h3>Text</h3>
      <p
        v-if="!c.paragraphsEnabled.value && !c.bodyEnabled.value"
        class="hint"
      >
        Select a text box or shape, or edit its text.
      </p>
    </header>

    <fieldset
      class="group"
      :disabled="!c.paragraphsEnabled.value"
    >
      <legend>Spacing</legend>
      <label class="field">
        <span class="field-label">Line spacing</span>
        <select
          data-testid="text-line-spacing-mode"
          :value="c.lineSpacingMode.value"
          @change="onMode"
        >
          <option
            value=""
            disabled
            hidden
          >
            Mixed
          </option>
          <option
            v-for="option in LINE_SPACING_MODE_OPTIONS"
            :key="option.value"
            :value="option.value"
          >
            {{ option.label }}
          </option>
        </select>
      </label>
      <label class="field">
        <span class="field-label">{{ isMultiple ? "Lines" : "Points" }}</span>
        <input
          type="number"
          data-testid="text-line-spacing"
          :min="isMultiple ? 0.25 : 4"
          :max="isMultiple ? 5 : 1000"
          :step="isMultiple ? 0.05 : 0.5"
          :value="inputValue(c.lineSpacingValue.value)"
          :aria-label="isMultiple ? 'Line spacing in lines' : 'Line spacing in points'"
          @change="onNumber('setLineSpacingValue', $event)"
        >
      </label>
      <label class="field">
        <span class="field-label">Before (pt)</span>
        <input
          type="number"
          min="0"
          max="500"
          step="1"
          data-testid="text-space-before"
          :value="inputValue(c.spaceBeforePt.value)"
          @change="onNumber('setSpaceBefore', $event)"
        >
      </label>
      <label class="field">
        <span class="field-label">After (pt)</span>
        <input
          type="number"
          min="0"
          max="500"
          step="1"
          data-testid="text-space-after"
          :value="inputValue(c.spaceAfterPt.value)"
          @change="onNumber('setSpaceAfter', $event)"
        >
      </label>
      <label class="field">
        <span class="field-label">Letter spacing (pt)</span>
        <input
          type="number"
          min="-10"
          max="10"
          step="0.1"
          data-testid="text-letter-spacing"
          :value="inputValue(c.letterSpacingPt.value)"
          @change="onNumber('setLetterSpacingPt', $event)"
        >
      </label>
    </fieldset>

    <fieldset
      class="group"
      :disabled="!c.bodyEnabled.value"
    >
      <legend>Text box</legend>
      <label class="field">
        <span class="field-label">Autofit</span>
        <select
          data-testid="text-autofit"
          :value="c.autofit.value"
          @change="onSelect('setAutofit', $event)"
        >
          <option
            v-for="option in AUTOFIT_OPTIONS"
            :key="option.value"
            :value="option.value"
          >
            {{ option.label }}
          </option>
        </select>
      </label>
      <label class="field">
        <span class="field-label">Anchor</span>
        <select
          data-testid="text-anchor"
          :value="c.anchor.value"
          @change="onSelect('setAnchor', $event)"
        >
          <option
            v-for="option in ANCHOR_OPTIONS"
            :key="option.value"
            :value="option.value"
          >
            {{ option.label }}
          </option>
        </select>
      </label>
      <label
        v-for="inset in INSET_SIDES"
        :key="inset.side"
        class="field"
      >
        <span class="field-label">{{ inset.label }} inset (cm)</span>
        <input
          type="number"
          min="0"
          max="50"
          step="0.05"
          :data-testid="`text-inset-${inset.side}`"
          :value="inputValue(insets[inset.side])"
          @change="onInset(inset.side, $event)"
        >
      </label>
      <label class="field">
        <span class="field-label">Columns</span>
        <input
          type="number"
          min="1"
          max="16"
          step="1"
          data-testid="text-columns"
          :value="inputValue(c.columns.value?.count)"
          @change="onNumber('setColumnCount', $event)"
        >
      </label>
      <label class="field">
        <span class="field-label">Column gap (cm)</span>
        <input
          type="number"
          min="0"
          max="50"
          step="0.05"
          data-testid="text-column-gap"
          :value="inputValue(c.columns.value?.gapCm)"
          @change="onNumber('setColumnGap', $event)"
        >
      </label>
    </fieldset>
  </section>
  <SlideLinkDialog
    v-if="c.linkDialog.value"
    :controller="controller"
  />
</template>

<style scoped>
.text-panel {
  flex: 0 0 auto;
  display: flex;
  flex-wrap: wrap;
  gap: var(--space-3) var(--space-4);
  padding: var(--space-3) var(--space-4);
  border-top: var(--border-width) solid var(--border-subtle);
  background: var(--surface-raised);
  color: var(--text-primary);
  font-size: var(--text-sm);
  max-height: 45%;
  overflow-y: auto;
}

.panel-header {
  flex: 1 0 100%;
  display: flex;
  align-items: baseline;
  gap: var(--space-3);
}

.panel-header h3 {
  margin: 0;
}

.hint {
  margin: 0;
  color: var(--text-secondary);
}

.group {
  display: flex;
  flex-wrap: wrap;
  align-items: flex-end;
  gap: var(--space-2) var(--space-3);
  margin: 0;
  padding: var(--space-2) var(--space-3);
  border: var(--border-width) solid var(--border-subtle);
  border-radius: var(--radius-sm);
}

.group legend {
  padding: 0 var(--space-1);
  color: var(--text-secondary);
}

.field {
  display: flex;
  flex-direction: column;
  gap: var(--space-1);
}

.field-label {
  color: var(--text-secondary);
}

.field input,
.field select {
  height: var(--height-control);
  padding: 0 var(--space-2);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-sm);
  background: var(--surface-default);
  color: var(--text-primary);
  font: inherit;
}

.field input {
  width: 5.5rem;
}
</style>
