<script setup lang="ts">
// The slide-transition side panel (SPEC-085): the effect list (None + six),
// direction, duration (seconds), advance-on-click, auto-advance, Preview and
// Apply-to-all for the current slide, plus the two inline confirmations
// (replacing a preserved out-of-subset transition, and the batched Apply-to-all
// replacement). All logic lives in the `useTransitions` controller passed in;
// this component is presentation + event forwarding only. Runtime prop
// declarations, not defineProps<{...}>() generics (espree lints .vue scripts).
import UiButton from "../ui/UiButton.vue";

const props = defineProps({
  /** The useTransitions() controller object (refs + command functions). */
  controller: { type: Object, required: true },
  slideNumber: { type: Number, default: 1 },
});

const EFFECTS = [
  { value: "none", label: "None" },
  { value: "fade", label: "Fade" },
  { value: "push", label: "Push" },
  { value: "wipe", label: "Wipe" },
  { value: "cover", label: "Cover" },
  { value: "uncover", label: "Uncover" },
  { value: "split", label: "Split" },
];
const DIRECTIONS = [
  { value: "left", label: "Left" },
  { value: "right", label: "Right" },
  { value: "up", label: "Up" },
  { value: "down", label: "Down" },
];

const c = props.controller;

function onEffect(value = "none") {
  c.setEffect(value);
}
function onDirection(event = new Event("change")) {
  const el = event.target;
  if (el instanceof HTMLSelectElement) c.setDirection(el.value);
}
function onDuration(event = new Event("change")) {
  const el = event.target;
  if (el instanceof HTMLInputElement) c.setDurationSeconds(Number(el.value));
}
function onAdvanceClick(event = new Event("change")) {
  const el = event.target;
  if (el instanceof HTMLInputElement) c.setAdvanceOnClick(el.checked);
}
function onAutoToggle(event = new Event("change")) {
  const el = event.target;
  if (el instanceof HTMLInputElement) c.setAutoAdvanceEnabled(el.checked);
}
function onAutoSeconds(event = new Event("change")) {
  const el = event.target;
  if (el instanceof HTMLInputElement) c.setAutoAdvanceSeconds(Number(el.value));
}
</script>

<template>
  <section
    class="transitions-panel"
    data-testid="slides-transition-panel"
    :aria-label="`Transition for slide ${slideNumber}`"
  >
    <!-- Keyed on formNonce so a change the user made in the DOM but then
         cancelled (a moved radio / typed input) is force-remounted back to the
         model, which stays unset (AC-8). -->
    <div
      :key="c.formNonce.value"
      class="transition-controls"
    >
      <fieldset
        class="effect-list"
        data-testid="transition-effect"
      >
        <legend>Effect</legend>
        <label
          v-for="effect in EFFECTS"
          :key="effect.value"
          class="effect-option"
        >
          <input
            type="radio"
            name="transition-effect"
            :value="effect.value"
            :data-testid="`transition-effect-${effect.value}`"
            :checked="c.selectedType.value === effect.value"
            @change="onEffect(effect.value)"
          >
          <span>{{ effect.label }}</span>
        </label>
        <p
          v-if="c.isPreserved.value"
          class="preserved-note"
          data-testid="transition-preserved-note"
        >
          Original transition kept — this editor can't show it.
        </p>
      </fieldset>

      <label class="field">
        <span>Direction</span>
        <select
          data-testid="transition-direction"
          :disabled="!c.directionEnabled.value"
          :value="c.direction.value"
          @change="onDirection"
        >
          <option
            v-for="d in DIRECTIONS"
            :key="d.value"
            :value="d.value"
          >
            {{ d.label }}
          </option>
        </select>
      </label>

      <label class="field">
        <span>Duration (seconds)</span>
        <input
          type="number"
          data-testid="transition-duration"
          min="0.1"
          max="5"
          step="0.1"
          :value="c.durationSeconds.value"
          @change="onDuration"
        >
      </label>

      <label class="check">
        <input
          type="checkbox"
          data-testid="transition-advance-click"
          :checked="c.advanceOnClick.value"
          @change="onAdvanceClick"
        >
        <span>Advance on click</span>
      </label>

      <div class="auto-advance">
        <label class="check">
          <input
            type="checkbox"
            data-testid="transition-advance-auto"
            :checked="c.autoAdvanceEnabled.value"
            @change="onAutoToggle"
          >
          <span>Auto-advance after</span>
        </label>
        <input
          type="number"
          data-testid="transition-advance-auto-seconds"
          min="0"
          step="0.1"
          :disabled="!c.autoAdvanceEnabled.value"
          :value="c.autoAdvanceSeconds.value"
          @change="onAutoSeconds"
        >
        <span>s</span>
      </div>
    </div>

    <p class="helper">
      Clicks always advance while presenting here.
    </p>
    <p class="helper">
      The "Advance on click" flag is honoured only by the exported viewer's
      kiosk reading, not by in-editor playback.
    </p>

    <div class="actions">
      <UiButton
        variant="ghost"
        data-testid="transition-preview"
        @click="c.preview()"
      >
        Preview
      </UiButton>
      <UiButton
        variant="secondary"
        data-testid="transition-apply-all"
        :disabled="!c.canApplyAll.value"
        @click="c.applyToAll()"
      >
        Apply to all slides
      </UiButton>
    </div>
    <p
      v-if="!c.canApplyAll.value"
      class="helper"
      data-testid="transition-apply-all-hint"
    >
      Choose an effect first to apply it to all slides.
    </p>

    <div
      v-if="c.pendingReplace.value"
      class="confirm"
      role="alertdialog"
      aria-label="Replace preserved transition"
      data-testid="transition-replace-confirm"
    >
      <p>Replace the original transition this editor can't show?</p>
      <div class="confirm-actions">
        <UiButton
          variant="secondary"
          data-testid="transition-replace-cancel"
          @click="c.cancelReplace()"
        >
          Cancel
        </UiButton>
        <UiButton
          variant="primary"
          data-testid="transition-replace-confirm-button"
          @click="c.confirmReplace()"
        >
          Replace
        </UiButton>
      </div>
    </div>

    <div
      v-if="c.pendingApplyAllCount.value > 0"
      class="confirm"
      role="alertdialog"
      aria-label="Replace preserved transitions on other slides"
      data-testid="transition-apply-all-confirm"
    >
      <p>
        Replace the original transition on {{ c.pendingApplyAllCount.value }}
        slide(s) this editor can't show?
      </p>
      <div class="confirm-actions">
        <UiButton
          variant="secondary"
          data-testid="transition-apply-all-cancel"
          @click="c.cancelApplyAll()"
        >
          Cancel
        </UiButton>
        <UiButton
          variant="primary"
          data-testid="transition-apply-all-confirm-button"
          @click="c.confirmApplyAll()"
        >
          Replace on all
        </UiButton>
      </div>
    </div>
  </section>
</template>

<style scoped>
.transitions-panel {
  flex: 0 0 auto;
  display: flex;
  flex-direction: column;
  gap: var(--space-3);
  padding: var(--space-4);
  border-top: var(--border-width) solid var(--border-subtle);
  background: var(--surface-default);
  color: var(--text-primary);
  font-size: var(--text-sm);
  max-height: 40%;
  overflow-y: auto;
}

.effect-list {
  border: var(--border-width) solid var(--border-subtle);
  border-radius: var(--radius-sm);
  padding: var(--space-2) var(--space-3);
  display: flex;
  flex-wrap: wrap;
  gap: var(--space-2) var(--space-4);
}

.effect-list legend {
  font-weight: var(--weight-medium);
  padding: 0 var(--space-1);
}

.effect-option {
  display: inline-flex;
  align-items: center;
  gap: var(--space-1);
}

.preserved-note {
  flex-basis: 100%;
  margin: var(--space-1) 0 0;
  color: var(--text-secondary);
  font-style: italic;
}

.field {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: var(--space-3);
}

.field select,
.field input,
.auto-advance input {
  min-width: 6rem;
}

.check {
  display: inline-flex;
  align-items: center;
  gap: var(--space-2);
}

.auto-advance {
  display: flex;
  align-items: center;
  gap: var(--space-2);
}

.helper {
  margin: 0;
  color: var(--text-secondary);
}

.actions {
  display: flex;
  gap: var(--space-3);
}

.confirm {
  border: var(--border-width) solid var(--border-strong, var(--border-subtle));
  border-radius: var(--radius-sm);
  padding: var(--space-3);
  background: var(--surface-sunken);
  display: flex;
  flex-direction: column;
  gap: var(--space-2);
}

.confirm p {
  margin: 0;
}

.confirm-actions {
  display: flex;
  gap: var(--space-3);
  justify-content: flex-end;
}
</style>
