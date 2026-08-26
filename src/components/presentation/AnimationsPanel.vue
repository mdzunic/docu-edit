<script setup lang="ts">
// The animations side panel (SPEC-086): an ordered effect list (each row: order
// number, target frame name, effect, trigger glyph), the selected row's field
// controls (effect, direction, trigger, duration, delay), Add/Remove/Move
// up-down/Play buttons, and the inline confirm for replacing a slide's
// preserved out-of-subset timing. All logic lives in the `useAnimations`
// controller passed in; this component is presentation + event forwarding only.
// Every control is keyboard-operable (native buttons/select/number inputs) and
// carries the spec's testids. Runtime prop declarations, not
// defineProps<{...}>() generics (espree lints .vue scripts).
import UiButton from "../ui/UiButton.vue";

const props = defineProps({
  /** The useAnimations() controller object (refs + command functions). */
  controller: { type: Object, required: true },
  slideNumber: { type: Number, default: 1 },
});

const EFFECTS = [
  { value: "appear", label: "Appear" },
  { value: "fadeIn", label: "Fade in" },
  { value: "flyIn", label: "Fly in" },
  { value: "wipeIn", label: "Wipe in" },
  { value: "zoomIn", label: "Zoom in" },
  { value: "fadeOut", label: "Fade out" },
  { value: "flyOut", label: "Fly out" },
];
const DIRECTIONS = [
  { value: "left", label: "Left" },
  { value: "right", label: "Right" },
  { value: "up", label: "Up" },
  { value: "down", label: "Down" },
];
const TRIGGERS = [
  { value: "click", label: "On click" },
  { value: "withPrevious", label: "With previous" },
  { value: "afterPrevious", label: "After previous" },
];
const c = props.controller;

function effectLabel(value = "") {
  return EFFECTS.find((e) => e.value === value)?.label ?? value;
}
/** A compact glyph per trigger for the row summary (AC-6). */
function triggerGlyph(value = "") {
  if (value === "withPrevious") return "⇊";
  if (value === "afterPrevious") return "↳";
  return "▷";
}
function onEffect(event = new Event("change")) {
  const el = event.target;
  if (el instanceof HTMLSelectElement) c.setEffect(el.value);
}
function onDirection(event = new Event("change")) {
  const el = event.target;
  if (el instanceof HTMLSelectElement) c.setDirection(el.value);
}
function onTrigger(event = new Event("change")) {
  const el = event.target;
  if (el instanceof HTMLSelectElement) c.setTrigger(el.value);
}
function onDuration(event = new Event("change")) {
  const el = event.target;
  if (el instanceof HTMLInputElement) c.setDurationSeconds(Number(el.value));
}
function onDelay(event = new Event("change")) {
  const el = event.target;
  if (el instanceof HTMLInputElement) c.setDelaySeconds(Number(el.value));
}
</script>

<template>
  <section
    class="animations-panel"
    data-testid="slides-animation-panel"
    :aria-label="`Animations for slide ${slideNumber}`"
  >
    <ul
      class="effect-list"
      data-testid="animation-list"
      role="listbox"
      :aria-label="`Animations for slide ${slideNumber}`"
    >
      <li
        v-for="row in c.rows.value"
        :key="row.index"
      >
        <button
          type="button"
          class="row"
          role="option"
          :class="{ 'is-selected': row.index === c.selectedIndex.value }"
          :aria-selected="row.index === c.selectedIndex.value"
          data-testid="animation-row"
          @click="c.selectRow(row.index)"
        >
          <span class="row-num">{{ row.number }}</span>
          <span class="row-name">{{ row.frameName }}</span>
          <span class="row-effect">{{ effectLabel(row.effect) }}</span>
          <span
            class="row-trigger"
            :title="row.trigger"
            aria-hidden="true"
          >{{ triggerGlyph(row.trigger) }}</span>
        </button>
      </li>
      <li
        v-if="c.rows.value.length === 0"
        class="empty"
        data-testid="animation-list-empty"
      >
        No animations yet.
      </li>
    </ul>

    <div class="toolbar-row">
      <UiButton
        variant="secondary"
        data-testid="animation-add"
        :disabled="!c.canAdd.value"
        @click="c.add()"
      >
        Add
      </UiButton>
      <UiButton
        variant="ghost"
        data-testid="animation-remove"
        :disabled="!c.hasSelection.value"
        @click="c.remove()"
      >
        Remove
      </UiButton>
      <UiButton
        variant="ghost"
        data-testid="animation-move-up"
        :disabled="!c.canMoveUp.value"
        @click="c.moveUp()"
      >
        Move up
      </UiButton>
      <UiButton
        variant="ghost"
        data-testid="animation-move-down"
        :disabled="!c.canMoveDown.value"
        @click="c.moveDown()"
      >
        Move down
      </UiButton>
      <UiButton
        variant="ghost"
        data-testid="animation-play"
        :disabled="c.rows.value.length === 0"
        @click="c.play()"
      >
        Play
      </UiButton>
    </div>

    <p
      v-if="!c.canAdd.value && c.addHelp.value"
      class="helper"
      data-testid="animation-add-help"
    >
      {{ c.addHelp.value }}
    </p>

    <div
      v-if="c.hasSelection.value"
      class="controls"
    >
      <label class="field">
        <span>Effect</span>
        <select
          data-testid="animation-effect"
          :value="c.selectedEffect.value"
          @change="onEffect"
        >
          <option
            v-for="e in EFFECTS"
            :key="e.value"
            :value="e.value"
          >
            {{ e.label }}
          </option>
        </select>
      </label>

      <label class="field">
        <span>Direction</span>
        <select
          data-testid="animation-direction"
          :disabled="!c.directionEnabled.value"
          :value="c.selectedDirection.value"
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
        <span>Trigger</span>
        <select
          data-testid="animation-trigger"
          :disabled="c.triggerLocked.value"
          :value="c.selectedTrigger.value"
          @change="onTrigger"
        >
          <option
            v-for="t in TRIGGERS"
            :key="t.value"
            :value="t.value"
          >
            {{ t.label }}
          </option>
        </select>
      </label>

      <label class="field">
        <span>Duration (seconds)</span>
        <input
          type="number"
          data-testid="animation-duration"
          min="0.1"
          max="5"
          step="0.1"
          :value="c.durationSeconds.value"
          @change="onDuration"
        >
      </label>

      <label class="field">
        <span>Delay (seconds)</span>
        <input
          type="number"
          data-testid="animation-delay"
          min="0"
          max="10"
          step="0.1"
          :value="c.delaySeconds.value"
          @change="onDelay"
        >
      </label>
    </div>

    <div
      v-if="c.pendingReplaceFrameId.value"
      class="confirm"
      role="alertdialog"
      aria-label="Replace preserved animations"
      data-testid="animation-replace-confirm"
    >
      <p>Replace the original animations this editor can't show?</p>
      <div class="confirm-actions">
        <UiButton
          variant="secondary"
          data-testid="animation-replace-cancel"
          @click="c.cancelReplace()"
        >
          Cancel
        </UiButton>
        <UiButton
          variant="primary"
          data-testid="animation-replace-confirm-button"
          @click="c.confirmReplace()"
        >
          Replace
        </UiButton>
      </div>
    </div>
  </section>
</template>

<style scoped>
.animations-panel {
  flex: 0 0 auto;
  display: flex;
  flex-direction: column;
  gap: var(--space-3);
  padding: var(--space-4);
  border-top: var(--border-width) solid var(--border-subtle);
  background: var(--surface-default);
  color: var(--text-primary);
  font-size: var(--text-sm);
  max-height: 45%;
  overflow-y: auto;
}

.effect-list {
  list-style: none;
  margin: 0;
  padding: 0;
  border: var(--border-width) solid var(--border-subtle);
  border-radius: var(--radius-sm);
  max-height: 10rem;
  overflow-y: auto;
}

.row {
  display: flex;
  align-items: center;
  gap: var(--space-2);
  width: 100%;
  padding: var(--space-1) var(--space-2);
  border: none;
  background: transparent;
  color: inherit;
  text-align: left;
  cursor: pointer;
  font-size: inherit;
}

.row:hover {
  background: var(--surface-sunken);
}

.row.is-selected {
  background: var(--accent-soft, var(--surface-sunken));
}

.row-num {
  min-width: 1.25rem;
  font-variant-numeric: tabular-nums;
  color: var(--text-secondary);
}

.row-name {
  flex: 1 1 auto;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.row-effect {
  color: var(--text-secondary);
}

.row-trigger {
  min-width: 1rem;
  text-align: center;
}

.empty {
  padding: var(--space-2);
  color: var(--text-secondary);
  font-style: italic;
}

.toolbar-row {
  display: flex;
  flex-wrap: wrap;
  gap: var(--space-2);
}

.controls {
  display: flex;
  flex-direction: column;
  gap: var(--space-2);
}

.field {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: var(--space-3);
}

.field select,
.field input {
  min-width: 7rem;
}

.helper {
  margin: 0;
  color: var(--text-secondary);
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
